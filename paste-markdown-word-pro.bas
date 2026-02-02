' ==========================================================================
' MODULE:       MarkdownToWord_Pro
' DESCRIPTION:  Converts Markdown syntax on the clipboard to formatted Word
'               text. Optimized for large documents (2000+ lines).
'
' DEPENDENCIES:
'   1. Microsoft VBScript Regular Expressions 5.5
'      (Go to Tools > References > Check "Microsoft VBScript Regular Expressions 5.5")
'
' KEY FEATURES:
'   - Reverse Iteration:    Prevents "Index Drift" where formatting misses text.
'   - Non-Destructive:      Deletes Markdown markers (###) but styles existing text.
'   - Memory Management:    Clears Undo Stack to prevent RAM overflow on large files.
'   - Stability Checks:     Disables SpellCheck/Pagination to prevent freezing.
'   - Tested on 2500-3000 lines inputs, takes typically 200-300 seconds
' ==========================================================================

Option Explicit

' --------------------------------------------------------------------------
' Helper: Ensure "Code" style exists for Monospace blocks
' --------------------------------------------------------------------------
Private Sub EnsureCodeStyle()
    Dim sty As Style
    On Error Resume Next
    Set sty = ActiveDocument.Styles("Code")
    On Error GoTo 0
    
    ' Create style if it doesn't exist
    If sty Is Nothing Then
        Set sty = ActiveDocument.Styles.Add(Name:="Code", Type:=wdStyleTypeParagraph)
        With sty
            .Font.Name = "Consolas"
            .Shading.BackgroundPatternColor = wdColorGray15
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.LeftIndent = 0 ' Optional: Tweaks for code look
        End With
    End If
End Sub

' --------------------------------------------------------------------------
' Main Routine
' --------------------------------------------------------------------------
Public Sub PasteMarkdown()
    Dim rng As Range
    Dim startPos As Long
    Dim tStart As Single
    
    ' -- 1. PERFORMANCE & STABILITY SETTINGS --
    ' Store current settings to restore them later
    Dim wasSpellCheck As Boolean, wasGrammar As Boolean
    wasSpellCheck = Options.CheckSpellingAsYouType
    wasGrammar = Options.CheckGrammarAsYouType
    
    ' Disable resource-heavy background tasks
    Application.ScreenUpdating = False
    Application.Options.Pagination = False
    Options.CheckSpellingAsYouType = False
    Options.CheckGrammarAsYouType = False
    
    tStart = Timer
    EnsureCodeStyle

    ' -- 2. PASTE FROM CLIPBOARD --
    startPos = Selection.Range.Start
    
    ' Use error handling in case clipboard is empty or non-text
    On Error Resume Next
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    If Err.Number <> 0 Then
        MsgBox "Clipboard is empty or invalid.", vbExclamation
        GoTo Cleanup
    End If
    On Error GoTo 0
    
    ' Define range: From insertion point to end of doc
    Set rng = ActiveDocument.Range(Start:=startPos, End:=ActiveDocument.Range.End)

    ' -- 3. NORMALIZE LINE ENDINGS --
    ' Ensure consistent paragraph marks (^p) instead of manual line breaks (^l)
    With rng.Find
        .ClearFormatting
        .Text = "^l"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    
    ' -- 4. PHASE 1: CODE BLOCKS (Top-Down) --
    ' Detects ``` blocks first so we don't accidentally format content inside them.
    ProcessCodeBlocks rng

    ' -- 5. PHASE 2: FORMATTING (Bottom-Up) --
    ' Applies Headings, Lists, and Inline styles.
    ProcessFormattingReverse rng

    ' -- 6. CLEANUP & FINAL FLUSH --
    ' Remove empty paragraphs created by marker deletions
    RemoveEmptyParagraphs rng
    
    ' FINAL MEMORY CLEAR:
    ' Clears the cleanup actions from the Undo Stack.
    ActiveDocument.UndoClear

Cleanup:
    ' Restore user settings
    Options.CheckSpellingAsYouType = wasSpellCheck
    Options.CheckGrammarAsYouType = wasGrammar
    Application.ScreenUpdating = True
    Application.Options.Pagination = True
    Application.StatusBar = False
    
    MsgBox "Processing Complete." & vbCrLf & _
           "Time: " & Format(Timer - tStart, "0.0") & " seconds.", vbInformation
End Sub

' --------------------------------------------------------------------------
' Logic: Process Code Fences (```)
' Direction: Forward (Top-Down)
' Reason: We need to match opening ``` with closing ``` sequentially.
' --------------------------------------------------------------------------
Private Sub ProcessCodeBlocks(rng As Range)
    Dim para As Paragraph
    Dim inCode As Boolean
    Dim i As Long
    Dim txt As String
    
    ' Regex to strictly match lines starting with ```
    Dim reFence As Object
    Set reFence = CreateObject("VBScript.RegExp")
    reFence.pattern = "^```"
    reFence.Global = False
    
    inCode = False
    Application.StatusBar = "Phase 1: Analyzing Code Blocks..."
    
    For i = 1 To rng.Paragraphs.Count
        ' Break if index exceeds count (due to deletions)
        If i > rng.Paragraphs.Count Then Exit For
        
        Set para = rng.Paragraphs(i)
        txt = para.Range.Text
        
        ' Optimization: Only check Regex if backtick exists
        If InStr(1, txt, "`") > 0 Then
            If reFence.Test(txt) Then
                ' Found a fence: Delete the line, toggle "inCode" state
                para.Range.Delete
                inCode = Not inCode
                
                ' Step back one index because paragraphs shifted up
                i = i - 1
                GoTo NextIter
            End If
        End If
        
        ' Style the block contents
        If inCode Then
            para.Style = ActiveDocument.Styles("Code")
            para.Range.NoProofing = True ' Performance: Skip spellcheck in code
        End If

NextIter:
        ' Keep UI responsive every 200 lines
        If i Mod 200 = 0 Then DoEvents
    Next i
End Sub

' --------------------------------------------------------------------------
' Logic: Process Markdown Syntax (Headers, Lists, Inline)
' Direction: Reverse (Bottom-Up)
' Reason: Prevents "Index Drift". Modifying paragraph 2000 won't affect the
'         index of paragraph 100, keeping the loop stable.
' --------------------------------------------------------------------------
Private Sub ProcessFormattingReverse(rng As Range)
    Dim i As Long
    Dim para As Paragraph
    Dim paraRng As Range
    Dim txt As String
    Dim total As Long
    
    ' -- COMPILE REGEX OBJECTS ONCE (Memory Optimization) --
    Dim reHead As Object, reQuote As Object
    Dim reListU As Object, reListO As Object
    Dim reInline As Object, reLink As Object, reStrike As Object
    
    Set reHead = CreateObject("VBScript.RegExp"): reHead.pattern = "^(#{1,6})\s+"
    Set reQuote = CreateObject("VBScript.RegExp"): reQuote.pattern = "^>\s*"
    Set reListU = CreateObject("VBScript.RegExp"): reListU.pattern = "^(\s*)([-\*\+])\s+"
    Set reListO = CreateObject("VBScript.RegExp"): reListO.pattern = "^(\s*)(\d+)\.\s+"
    
    Set reInline = CreateObject("VBScript.RegExp"): reInline.Global = True
    Set reLink = CreateObject("VBScript.RegExp"): reLink.pattern = "\[(.+?)\]\((https?:\/\/[^\s\)]+)\)": reLink.Global = True
    Set reStrike = CreateObject("VBScript.RegExp"): reStrike.pattern = "~~(.+?)~~": reStrike.Global = True

    total = rng.Paragraphs.Count
    
    ' Loop backwards from end to start
    For i = total To 1 Step -1
        Set para = rng.Paragraphs(i)
        
        ' Skip "Code" blocks entirely
        If para.Style <> "Code" Then
            Set paraRng = para.Range
            txt = paraRng.Text
            
            ' -- STATUS & MEMORY FLUSH --
            If i Mod 100 = 0 Then
                Application.StatusBar = "Phase 2: Formatting line " & i & " of " & total
                ' Clear Undo Stack periodically to free RAM
                ActiveDocument.UndoClear
                DoEvents
            End If
            
            ' -- FAST-FAIL CHECKS --
            ' Only run regex if relevant characters exist (100x speedup on plain text)
            If Len(txt) > 1 Then
                
                ' 1. Block Elements (Check first char matches pattern start)
                If InStr(1, txt, "#") > 0 Then
                    If Left(LTrim(txt), 1) = "#" Then ApplyHeading paraRng, reHead
                End If
                
                If InStr(1, txt, ">") > 0 Then
                    If Left(LTrim(txt), 1) = ">" Then ApplyQuote paraRng, reQuote
                End If
                
                If InStr(1, txt, "-") > 0 Or InStr(1, txt, "*") > 0 Or InStr(1, txt, "+") > 0 Then
                    ApplyListU paraRng, reListU
                ElseIf IsNumeric(Left(LTrim(txt), 1)) Then
                    ApplyListO paraRng, reListO
                End If
                
                ' 2. Inline Elements
                ' Update 'txt' as block operations might have changed it
                txt = paraRng.Text
                
                If InStr(1, txt, "`") > 0 Then RunRegexInlineCode paraRng, reInline
                
                If InStr(1, txt, "*") > 0 Or InStr(1, txt, "_") > 0 Then
                    RunRegexInlineFormat paraRng, reInline
                End If
                
                If InStr(1, txt, "[") > 0 Then RunRegexLink paraRng, reLink
                If InStr(1, txt, "~") > 0 Then RunRegexStrike paraRng, reStrike
                
            End If
        End If
    Next i
End Sub

' --------------------------------------------------------------------------
' Formatting Helpers (Non-Destructive)
' --------------------------------------------------------------------------

' Deletes "### " marker, applies Heading Style
Private Sub ApplyHeading(rng As Range, re As Object)
    Dim m As Object
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        Dim lvl As Long
        lvl = Len(Trim(m.Value))
        If lvl > 6 Then lvl = 6
        
        ActiveDocument.Range(rng.Start, rng.Start + m.Length).Delete
        rng.Style = ActiveDocument.Styles("Heading " & lvl)
    End If
End Sub

' Deletes "> " marker, applies Quote Style
Private Sub ApplyQuote(rng As Range, re As Object)
    Dim m As Object
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        ActiveDocument.Range(rng.Start, rng.Start + m.Length).Delete
        rng.Style = ActiveDocument.Styles("Quote")
    End If
End Sub

' Deletes "- " marker, applies Bullet List + Indentation
Private Sub ApplyListU(rng As Range, re As Object)
    Dim m As Object
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        
        Dim leadingSpaces As Long
        leadingSpaces = Len(m.SubMatches(0))
        
        ActiveDocument.Range(rng.Start, rng.Start + m.Length).Delete
        rng.ListFormat.ApplyBulletDefault
        
        ' Indent logic: 2 spaces = 1 indent
        Dim k As Long
        For k = 1 To (leadingSpaces \ 2)
            rng.ListFormat.ListIndent
        Next k
    End If
End Sub

' Deletes "1. " marker, applies Number List + Indentation
Private Sub ApplyListO(rng As Range, re As Object)
    Dim m As Object
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        
        Dim leadingSpaces As Long
        leadingSpaces = Len(m.SubMatches(0))
        
        ActiveDocument.Range(rng.Start, rng.Start + m.Length).Delete
        rng.ListFormat.ApplyNumberDefault
        
        Dim k As Long
        For k = 1 To (leadingSpaces \ 2)
            rng.ListFormat.ListIndent
        Next k
    End If
End Sub

' Formats `code spans`
Private Sub RunRegexInlineCode(rng As Range, re As Object)
    Dim mCol As Object, m As Object
    Dim startP As Long, markLen As Long, rWhole As Range
    re.pattern = "(`+)(.+?)\1"
    
    Do
        Set mCol = re.Execute(rng.Text)
        If mCol.Count = 0 Then Exit Do
        Set m = mCol(mCol.Count - 1)
        
        markLen = Len(m.SubMatches(0))
        startP = rng.Start + m.FirstIndex
        Set rWhole = ActiveDocument.Range(startP, startP + m.Length)
        
        ' Delete markers from Ends first to keep Start index valid
        rWhole.Characters.Last.Delete markLen
        rWhole.Characters.First.Delete markLen
        rWhole.Font.Name = "Consolas"
        rWhole.Shading.BackgroundPatternColor = wdColorGray15
    Loop
End Sub

' Formats **Bold** and *Italic*
Private Sub RunRegexInlineFormat(rng As Range, re As Object)
    Dim patterns As Variant, i As Long
    patterns = Array("(\*\*\*|___)(.+?)\1", "(\*\*|__)(.+?)\1", "(\*|_)(.+?)\1")
    Dim types As Variant
    types = Array(0, 1, 2) ' 0=Both, 1=Bold, 2=Italic
    
    For i = 0 To 2
        re.pattern = patterns(i)
        ApplyFormatSimple rng, re, CLng(types(i))
    Next i
End Sub

' Formats Links [Text](URL)
Private Sub RunRegexLink(rng As Range, re As Object)
    Dim mCol As Object, m As Object
    Dim startP As Long, rWhole As Range
    Dim txt As String, url As String
    
    Do
        Set mCol = re.Execute(rng.Text)
        If mCol.Count = 0 Then Exit Do
        Set m = mCol(mCol.Count - 1)
        
        startP = rng.Start + m.FirstIndex
        Set rWhole = ActiveDocument.Range(startP, startP + m.Length)
        
        txt = m.SubMatches(0)
        url = m.SubMatches(1)
        
        rWhole.Text = txt
        ActiveDocument.Hyperlinks.Add Anchor:=rWhole, Address:=url, TextToDisplay:=txt
    Loop
End Sub

' Formats ~~Strikethrough~~
Private Sub RunRegexStrike(rng As Range, re As Object)
    ApplyFormatSimple rng, re, 4
End Sub

' Generic Helper for Font Styles
Private Sub ApplyFormatSimple(rng As Range, re As Object, fmtType As Long)
    Dim mCol As Object, m As Object
    Dim startP As Long, markLen As Long, rWhole As Range, rMarker As Range
    
    Do
        Set mCol = re.Execute(rng.Text)
        If mCol.Count = 0 Then Exit Do
        Set m = mCol(mCol.Count - 1)
        
        markLen = Len(m.SubMatches(0))
        startP = rng.Start + m.FirstIndex
        Set rWhole = ActiveDocument.Range(startP, startP + m.Length)
        
        Set rMarker = ActiveDocument.Range(rWhole.End - markLen, rWhole.End): rMarker.Delete
        Set rMarker = ActiveDocument.Range(rWhole.Start, rWhole.Start + markLen): rMarker.Delete
        
        Select Case fmtType
            Case 0: rWhole.Font.Bold = True: rWhole.Font.Italic = True
            Case 1: rWhole.Font.Bold = True
            Case 2: rWhole.Font.Italic = True
            Case 4: rWhole.Font.StrikeThrough = True
        End Select
    Loop
End Sub

' Cleanup Helper: Deletes empty paragraphs (leftover from marker deletion)
Private Sub RemoveEmptyParagraphs(rng As Range)
    Dim i As Long
    Dim p As Paragraph
    
    For i = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(i)
        ' Check if paragraph is effectively empty (just CR)
        If Len(p.Range.Text) <= 1 And p.Style <> "Code" Then
            p.Range.Delete
        End If
    Next i
End Sub

