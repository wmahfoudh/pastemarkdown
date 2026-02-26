' ==========================================================================
' MODULE:       OutlookMarkdownPaster
' DESCRIPTION:  Pastes Clipboard Markdown into Outlook Email as Formatted Text
' DEPENDENCIES: 1. Microsoft Word 16.0 Object Library
'               2. Microsoft VBScript Regular Expressions 5.5
' ==========================================================================

Option Explicit

' --------------------------------------------------------------------------
' PUBLIC MACRO: Link this to your Ribbon Button
' --------------------------------------------------------------------------
Public Sub PasteMarkdownToEmail()
    Dim insp As Outlook.Inspector
    Dim doc As Word.Document
    Dim sel As Word.Selection
    Dim rng As Word.Range
    Dim startPos As Long
    Dim tStart As Single
    
    ' 1. Validate we are in an email editor
    Set insp = Application.ActiveInspector
    If insp Is Nothing Then Exit Sub
    If insp.EditorType <> olEditorWord Then Exit Sub
    
    ' 2. Get the Word Editor (The email body)
    Set doc = insp.WordEditor
    Set sel = doc.Windows(1).Selection
    
    ' Performance settings
    Dim wasSpell As Boolean
    wasSpell = doc.Application.Options.CheckSpellingAsYouType
    doc.Application.ScreenUpdating = False
    doc.Application.Options.CheckSpellingAsYouType = False
    
    tStart = Timer
    EnsureCodeStyle doc
    
    ' 3. Paste
    startPos = sel.Range.Start
    On Error Resume Next
    sel.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    If Err.Number <> 0 Then
        MsgBox "Clipboard is empty or invalid.", vbExclamation
        GoTo Cleanup
    End If
    On Error GoTo 0
    
    ' 4. Process Range
    Set rng = doc.Range(Start:=startPos, End:=doc.Range.End)

    ' Normalize Line Breaks
    With rng.Find
        .ClearFormatting
        .Text = "^l"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Run Logic
    ProcessCodeBlocks rng
    ProcessFormattingReverse rng
    RemoveEmptyParagraphs rng
    
    ' Outlook specific cleanup
    doc.UndoClear

Cleanup:
    doc.Application.Options.CheckSpellingAsYouType = wasSpell
    doc.Application.ScreenUpdating = True
    
    ' Optional: Remove this MsgBox if you want silent operation
    ' MsgBox "Done (" & Format(Timer - tStart, "0.0") & "s)", vbInformation
End Sub

' --------------------------------------------------------------------------
' HELPER: Styles
' --------------------------------------------------------------------------
Private Sub EnsureCodeStyle(doc As Word.Document)
    Dim sty As Word.Style
    On Error Resume Next
    Set sty = doc.Styles("Code")
    On Error GoTo 0
    
    If sty Is Nothing Then
        Set sty = doc.Styles.Add(Name:="Code", Type:=wdStyleTypeParagraph)
        With sty
            .Font.Name = "Consolas"
            .Shading.BackgroundPatternColor = wdColorGray15
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .Font.Color = wdColorBlack ' Outlook often defaults to blue, force black
            .Font.Size = 10
        End With
    End If
End Sub

' --------------------------------------------------------------------------
' LOGIC: Code Blocks (Top-Down)
' --------------------------------------------------------------------------
Private Sub ProcessCodeBlocks(rng As Word.Range)
    Dim para As Word.Paragraph
    Dim inCode As Boolean
    Dim i As Long
    Dim txt As String
    Dim doc As Word.Document
    Set doc = rng.Parent
    
    Dim reFence As Object
    Set reFence = CreateObject("VBScript.RegExp")
    reFence.Pattern = "^```"
    reFence.Global = False
    
    inCode = False
    
    For i = 1 To rng.Paragraphs.Count
        If i > rng.Paragraphs.Count Then Exit For
        
        Set para = rng.Paragraphs(i)
        txt = para.Range.Text
        
        If InStr(1, txt, "`") > 0 Then
            If reFence.Test(txt) Then
                para.Range.Delete
                inCode = Not inCode
                i = i - 1
                GoTo NextIter
            End If
        End If
        
        If inCode Then
            para.Style = doc.Styles("Code")
            para.Range.NoProofing = True
        End If
NextIter:
    Next i
End Sub

' --------------------------------------------------------------------------
' LOGIC: Formatting (Bottom-Up)
' --------------------------------------------------------------------------
Private Sub ProcessFormattingReverse(rng As Word.Range)
    Dim i As Long
    Dim para As Word.Paragraph
    Dim paraRng As Word.Range
    Dim txt As String
    Dim total As Long
    Dim doc As Word.Document
    Set doc = rng.Parent
    
    ' REGEX SETUP
    Dim reHead As Object, reQuote As Object, reListU As Object, reListO As Object
    Dim reInline As Object, reLink As Object, reStrike As Object
    
    Set reHead = CreateObject("VBScript.RegExp"): reHead.Pattern = "^(#{1,6})\s+"
    Set reQuote = CreateObject("VBScript.RegExp"): reQuote.Pattern = "^>\s*"
    Set reListU = CreateObject("VBScript.RegExp"): reListU.Pattern = "^(\s*)([-\*\+])\s+"
    Set reListO = CreateObject("VBScript.RegExp"): reListO.Pattern = "^(\s*)(\d+)\.\s+"
    Set reInline = CreateObject("VBScript.RegExp"): reInline.Global = True
    Set reLink = CreateObject("VBScript.RegExp"): reLink.Pattern = "\[(.+?)\]\((https?:\/\/[^\s\)]+)\)": reLink.Global = True
    Set reStrike = CreateObject("VBScript.RegExp"): reStrike.Pattern = "~~(.+?)~~": reStrike.Global = True

    total = rng.Paragraphs.Count
    
    For i = total To 1 Step -1
        Set para = rng.Paragraphs(i)
        
        If para.Style <> "Code" Then
            Set paraRng = para.Range
            txt = paraRng.Text
            
            If i Mod 50 = 0 Then doc.UndoClear ' Memory management
            
            If Len(txt) > 1 Then
                ' Block Elements
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
                
                ' Inline Elements
                txt = paraRng.Text ' Update text after block changes
                If InStr(1, txt, "`") > 0 Then RunRegexInlineCode paraRng, reInline
                If InStr(1, txt, "*") > 0 Or InStr(1, txt, "_") > 0 Then RunRegexInlineFormat paraRng, reInline
                If InStr(1, txt, "[") > 0 Then RunRegexLink paraRng, reLink
                If InStr(1, txt, "~") > 0 Then RunRegexStrike paraRng, reStrike
            End If
        End If
    Next i
End Sub

' --------------------------------------------------------------------------
' HELPERS (Using rng.Parent to find Document)
' --------------------------------------------------------------------------
Private Sub ApplyHeading(rng As Word.Range, re As Object)
    Dim m As Object, doc As Word.Document: Set doc = rng.Parent
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        Dim lvl As Long: lvl = Len(Trim(m.Value))
        If lvl > 6 Then lvl = 6
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.Style = doc.Styles("Heading " & lvl)
    End If
End Sub

Private Sub ApplyQuote(rng As Word.Range, re As Object)
    Dim m As Object, doc As Word.Document: Set doc = rng.Parent
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.Style = doc.Styles("Quote")
    End If
End Sub

Private Sub ApplyListU(rng As Word.Range, re As Object)
    Dim m As Object, doc As Word.Document: Set doc = rng.Parent
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        Dim ls As Long: ls = Len(m.SubMatches(0))
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.ListFormat.ApplyBulletDefault
        Dim k As Long: For k = 1 To (ls \ 2): rng.ListFormat.ListIndent: Next k
    End If
End Sub

Private Sub ApplyListO(rng As Word.Range, re As Object)
    Dim m As Object, doc As Word.Document: Set doc = rng.Parent
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        Dim ls As Long: ls = Len(m.SubMatches(0))
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.ListFormat.ApplyNumberDefault
        Dim k As Long: For k = 1 To (ls \ 2): rng.ListFormat.ListIndent: Next k
    End If
End Sub

Private Sub RunRegexInlineCode(rng As Word.Range, re As Object)
    Dim mCol As Object, m As Object, startP As Long, markLen As Long, rWhole As Word.Range, doc As Word.Document: Set doc = rng.Parent
    re.Pattern = "(`+)(.+?)\1"
    Do
        Set mCol = re.Execute(rng.Text)
        If mCol.Count = 0 Then Exit Do
        Set m = mCol(mCol.Count - 1)
        markLen = Len(m.SubMatches(0))
        startP = rng.Start + m.FirstIndex
        Set rWhole = doc.Range(startP, startP + m.Length)
        rWhole.Characters.Last.Delete markLen
        rWhole.Characters.First.Delete markLen
        rWhole.Font.Name = "Consolas"
        rWhole.Shading.BackgroundPatternColor = wdColorGray15
    Loop
End Sub

Private Sub RunRegexInlineFormat(rng As Word.Range, re As Object)
    Dim patterns As Variant: patterns = Array("(\*\*\*|___)(.+?)\1", "(\*\*|__)(.+?)\1", "(\*|_)(.+?)\1")
    Dim types As Variant: types = Array(0, 1, 2)
    Dim i As Long
    For i = 0 To 2
        re.Pattern = patterns(i)
        ApplyFormatSimple rng, re, CLng(types(i))
    Next i
End Sub

Private Sub RunRegexLink(rng As Word.Range, re As Object)
    Dim mCol As Object, m As Object, startP As Long, rWhole As Word.Range, doc As Word.Document: Set doc = rng.Parent
    Do
        Set mCol = re.Execute(rng.Text)
        If mCol.Count = 0 Then Exit Do
        Set m = mCol(mCol.Count - 1)
        startP = rng.Start + m.FirstIndex
        Set rWhole = doc.Range(startP, startP + m.Length)
        Dim txt As String: txt = m.SubMatches(0)
        Dim url As String: url = m.SubMatches(1)
        rWhole.Text = txt
        doc.Hyperlinks.Add Anchor:=rWhole, Address:=url, TextToDisplay:=txt
    Loop
End Sub

Private Sub RunRegexStrike(rng As Word.Range, re As Object)
    ApplyFormatSimple rng, re, 4
End Sub

Private Sub ApplyFormatSimple(rng As Word.Range, re As Object, fmtType As Long)
    Dim mCol As Object, m As Object, startP As Long, markLen As Long, rWhole As Word.Range, rMarker As Word.Range, doc As Word.Document: Set doc = rng.Parent
    Do
        Set mCol = re.Execute(rng.Text)
        If mCol.Count = 0 Then Exit Do
        Set m = mCol(mCol.Count - 1)
        markLen = Len(m.SubMatches(0))
        startP = rng.Start + m.FirstIndex
        Set rWhole = doc.Range(startP, startP + m.Length)
        Set rMarker = doc.Range(rWhole.End - markLen, rWhole.End): rMarker.Delete
        Set rMarker = doc.Range(rWhole.Start, rWhole.Start + markLen): rMarker.Delete
        Select Case fmtType
            Case 0: rWhole.Font.Bold = True: rWhole.Font.Italic = True
            Case 1: rWhole.Font.Bold = True
            Case 2: rWhole.Font.Italic = True
            Case 4: rWhole.Font.Strikethrough = True
        End Select
    Loop
End Sub

Private Sub RemoveEmptyParagraphs(rng As Word.Range)
    Dim i As Long, p As Word.Paragraph
    For i = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(i)
        If Len(p.Range.Text) <= 1 And p.Style <> "Code" Then p.Range.Delete
    Next i
End Sub
