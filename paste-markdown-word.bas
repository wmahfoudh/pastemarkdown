' ==========================================================================
' MODULE:       WordMarkdownPaster
' DESCRIPTION: Pastes Clipboard Markdown into Word as Formatted Text
' DEPENDENCIES: Microsoft Word
'
' Notes:
' - Uses late-bound VBScript.RegExp, so no explicit Regex reference is required.
' - Designed to run directly inside Microsoft Word.
' ==========================================================================

Option Explicit

' --------------------------------------------------------------------------
' CONFIGURATION
' --------------------------------------------------------------------------
Private Const REMOVE_EMPTY_PARAGRAPHS_AFTER_PASTE As Boolean = True
Private Const CODE_STYLE_NAME As String = "Code"

' --------------------------------------------------------------------------
' PUBLIC MACRO: Link this to your Word Ribbon Button / Quick Access Toolbar
' --------------------------------------------------------------------------
Public Sub PasteMarkdownToWord()
    Dim doc As Word.Document
    Dim sel As Word.Selection
    Dim rng As Word.Range
    Dim startPos As Long
    Dim endPos As Long
    
    If Application.Documents.Count = 0 Then Exit Sub
    
    Set doc = ActiveDocument
    Set sel = Selection
    
    Dim wasSpell As Boolean
    Dim wasGrammar As Boolean
    Dim wasScreenUpdating As Boolean
    
    wasSpell = Application.Options.CheckSpellingAsYouType
    wasGrammar = Application.Options.CheckGrammarAsYouType
    wasScreenUpdating = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    Application.Options.CheckSpellingAsYouType = False
    Application.Options.CheckGrammarAsYouType = False
    
    On Error GoTo FailSafe
    
    EnsureCodeStyle doc
    
    ' Paste clipboard as plain text
    startPos = sel.Range.Start
    
    On Error Resume Next
    sel.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Clipboard is empty or invalid.", vbExclamation
        GoTo Cleanup
    End If
    On Error GoTo FailSafe
    
    endPos = sel.Range.End
    
    If endPos <= startPos Then GoTo Cleanup
    
    ' Process only the pasted content
    Set rng = doc.Range(Start:=startPos, End:=endPos)
    
    ' Normalize manual line breaks to paragraphs
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Refresh range after normalization
    Set rng = doc.Range(Start:=startPos, End:=sel.Range.End)
    
    ' Markdown processing order matters
    ProcessCodeBlocks rng
    ProcessMarkdownTables rng
    ProcessFormattingReverse rng
    
    If REMOVE_EMPTY_PARAGRAPHS_AFTER_PASTE Then
        RemoveEmptyParagraphs rng
    End If
    
    doc.UndoClear
    
Cleanup:
    On Error Resume Next
    Application.Options.CheckSpellingAsYouType = wasSpell
    Application.Options.CheckGrammarAsYouType = wasGrammar
    Application.ScreenUpdating = wasScreenUpdating
    On Error GoTo 0
    Exit Sub

FailSafe:
    MsgBox "Markdown paste failed: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

' --------------------------------------------------------------------------
' HELPER: Styles
' --------------------------------------------------------------------------
Private Sub EnsureCodeStyle(doc As Word.Document)
    Dim sty As Word.Style
    
    On Error Resume Next
    Set sty = doc.Styles(CODE_STYLE_NAME)
    On Error GoTo 0
    
    If sty Is Nothing Then
        Set sty = doc.Styles.Add(Name:=CODE_STYLE_NAME, Type:=wdStyleTypeParagraph)
    End If
    
    With sty
        .Font.Name = "Consolas"
        .Font.Color = wdColorBlack
        .Font.Size = 10
        .Shading.BackgroundPatternColor = wdColorGray15
        .NoSpaceBetweenParagraphsOfSameStyle = True
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
    End With
End Sub

' --------------------------------------------------------------------------
' LOGIC: Code Blocks
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
    reFence.Pattern = "^\s*```"
    reFence.Global = False
    reFence.IgnoreCase = True
    
    inCode = False
    
    For i = 1 To rng.Paragraphs.Count
        If i > rng.Paragraphs.Count Then Exit For
        
        Set para = rng.Paragraphs(i)
        txt = para.Range.Text
        
        If InStr(1, txt, "`", vbBinaryCompare) > 0 Then
            If reFence.Test(txt) Then
                para.Range.Delete
                inCode = Not inCode
                i = i - 1
                GoTo NextIter
            End If
        End If
        
        If inCode Then
            para.Style = doc.Styles(CODE_STYLE_NAME)
            para.Range.NoProofing = True
        End If
        
NextIter:
    Next i
End Sub

' --------------------------------------------------------------------------
' LOGIC: Markdown Tables
' --------------------------------------------------------------------------
Private Sub ProcessMarkdownTables(rng As Word.Range)
    Dim i As Long
    Dim total As Long
    
    total = rng.Paragraphs.Count
    
    For i = total - 1 To 1 Step -1
        
        If i + 1 <= rng.Paragraphs.Count Then
            
            Dim headerText As String
            Dim separatorText As String
            
            headerText = CleanParaText(rng.Paragraphs(i).Range.Text)
            separatorText = CleanParaText(rng.Paragraphs(i + 1).Range.Text)
            
            If Not IsParagraphCode(rng.Paragraphs(i)) _
               And Not IsParagraphCode(rng.Paragraphs(i + 1)) _
               And LooksLikeMarkdownTableRow(headerText) _
               And IsMarkdownTableSeparator(separatorText) Then
                
                Dim startPara As Long
                Dim endPara As Long
                Dim j As Long
                
                startPara = i
                endPara = i + 1
                j = i + 2
                
                Do While j <= rng.Paragraphs.Count
                    If IsParagraphCode(rng.Paragraphs(j)) Then Exit Do
                    
                    Dim rowText As String
                    rowText = CleanParaText(rng.Paragraphs(j).Range.Text)
                    
                    If LooksLikeMarkdownTableRow(rowText) Then
                        endPara = j
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                ConvertMarkdownTableBlock rng, startPara, endPara
            End If
        End If
    Next i
End Sub

Private Sub ConvertMarkdownTableBlock(rng As Word.Range, startPara As Long, endPara As Long)
    Dim doc As Word.Document
    Set doc = rng.Parent
    
    Dim rows As Collection
    Set rows = New Collection
    
    Dim separatorCells As Variant
    Dim alignments As Variant
    Dim colCount As Long
    Dim i As Long
    
    Dim headerCells As Variant
    headerCells = SplitMarkdownTableRow(CleanParaText(rng.Paragraphs(startPara).Range.Text))
    colCount = ArrayItemCount(headerCells)
    rows.Add headerCells
    
    separatorCells = SplitMarkdownTableRow(CleanParaText(rng.Paragraphs(startPara + 1).Range.Text))
    
    For i = startPara + 2 To endPara
        Dim dataCells As Variant
        dataCells = SplitMarkdownTableRow(CleanParaText(rng.Paragraphs(i).Range.Text))
        
        If ArrayItemCount(dataCells) > colCount Then
            colCount = ArrayItemCount(dataCells)
        End If
        
        rows.Add dataCells
    Next i
    
    alignments = GetMarkdownTableAlignments(separatorCells, colCount)
    
    Dim tableText As String
    tableText = BuildTabDelimitedTableText(rows, colCount)
    
    Dim blockStart As Long
    Dim blockEnd As Long
    blockStart = rng.Paragraphs(startPara).Range.Start
    blockEnd = rng.Paragraphs(endPara).Range.End
    
    Dim blockRng As Word.Range
    Set blockRng = doc.Range(blockStart, blockEnd)
    blockRng.Text = tableText
    
    Set blockRng = doc.Range(blockStart, blockStart + Len(tableText))
    
    Dim tbl As Word.Table
    Set tbl = blockRng.ConvertToTable( _
        Separator:=wdSeparateByTabs, _
        NumColumns:=colCount, _
        AutoFitBehavior:=wdAutoFitWindow)
    
    FormatMarkdownTable tbl, alignments
End Sub

Private Function BuildTabDelimitedTableText(rows As Collection, colCount As Long) As String
    Dim result As String
    Dim r As Long
    Dim c As Long
    
    For r = 1 To rows.Count
        Dim cells As Variant
        cells = rows(r)
        
        For c = 0 To colCount - 1
            If c <= UBound(cells) Then
                result = result & CleanMarkdownTableCell(CStr(cells(c)))
            End If
            
            If c < colCount - 1 Then
                result = result & vbTab
            End If
        Next c
        
        result = result & vbCr
    Next r
    
    BuildTabDelimitedTableText = result
End Function

Private Sub FormatMarkdownTable(tbl As Word.Table, alignments As Variant)
    On Error Resume Next
    
    tbl.Borders.Enable = True
    tbl.AutoFitBehavior wdAutoFitWindow
    
    With tbl.Range
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With
    
    With tbl.Rows(1).Range
        .Font.Bold = True
        .Shading.BackgroundPatternColor = wdColorGray15
    End With
    
    Dim c As Long
    Dim cell As Word.Cell
    
    For c = 1 To tbl.Columns.Count
        For Each cell In tbl.Columns(c).Cells
            cell.Range.ParagraphFormat.Alignment = CLng(alignments(c - 1))
        Next cell
    Next c
    
    tbl.TopPadding = 3
    tbl.BottomPadding = 3
    tbl.LeftPadding = 4
    tbl.RightPadding = 4
    
    On Error GoTo 0
End Sub

Private Function CleanParaText(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    CleanParaText = Trim(s)
End Function

Private Function IsParagraphCode(para As Word.Paragraph) As Boolean
    On Error Resume Next
    IsParagraphCode = (para.Style = CODE_STYLE_NAME)
    On Error GoTo 0
End Function

Private Function LooksLikeMarkdownTableRow(ByVal s As String) As Boolean
    s = Trim(s)
    
    If Len(s) = 0 Then
        LooksLikeMarkdownTableRow = False
        Exit Function
    End If
    
    If InStr(1, s, "|", vbBinaryCompare) = 0 Then
        LooksLikeMarkdownTableRow = False
        Exit Function
    End If
    
    If IsMarkdownTableSeparator(s) Then
        LooksLikeMarkdownTableRow = False
        Exit Function
    End If
    
    Dim cells As Variant
    cells = SplitMarkdownTableRow(s)
    
    LooksLikeMarkdownTableRow = (ArrayItemCount(cells) >= 2)
End Function

Private Function IsMarkdownTableSeparator(ByVal s As String) As Boolean
    s = Trim(s)
    
    If Len(s) = 0 Then
        IsMarkdownTableSeparator = False
        Exit Function
    End If
    
    If InStr(1, s, "|", vbBinaryCompare) = 0 Then
        IsMarkdownTableSeparator = False
        Exit Function
    End If
    
    Dim cells As Variant
    cells = SplitMarkdownTableRow(s)
    
    If ArrayItemCount(cells) < 2 Then
        IsMarkdownTableSeparator = False
        Exit Function
    End If
    
    Dim i As Long
    For i = LBound(cells) To UBound(cells)
        If Not IsMarkdownSeparatorCell(CStr(cells(i))) Then
            IsMarkdownTableSeparator = False
            Exit Function
        End If
    Next i
    
    IsMarkdownTableSeparator = True
End Function

Private Function IsMarkdownSeparatorCell(ByVal s As String) As Boolean
    Dim t As String
    t = Trim(s)
    t = Replace(t, " ", "")
    
    If Len(t) < 3 Then
        IsMarkdownSeparatorCell = False
        Exit Function
    End If
    
    If Left(t, 1) = ":" Then t = Mid(t, 2)
    If Right(t, 1) = ":" Then t = Left(t, Len(t) - 1)
    
    If Len(t) < 3 Then
        IsMarkdownSeparatorCell = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To Len(t)
        If Mid(t, i, 1) <> "-" Then
            IsMarkdownSeparatorCell = False
            Exit Function
        End If
    Next i
    
    IsMarkdownSeparatorCell = True
End Function

Private Function SplitMarkdownTableRow(ByVal s As String) As Variant
    s = CleanParaText(s)
    
    If Left(s, 1) = "|" Then s = Mid(s, 2)
    If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
    
    Dim arr() As String
    ReDim arr(0 To 0)
    
    Dim idx As Long
    Dim current As String
    Dim i As Long
    Dim ch As String
    
    idx = 0
    current = ""
    i = 1
    
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        
        If ch = "\" And i < Len(s) Then
            If Mid(s, i + 1, 1) = "|" Then
                current = current & "|"
                i = i + 2
            Else
                current = current & ch
                i = i + 1
            End If
            
        ElseIf ch = "|" Then
            arr(idx) = Trim(current)
            idx = idx + 1
            ReDim Preserve arr(0 To idx)
            current = ""
            i = i + 1
            
        Else
            current = current & ch
            i = i + 1
        End If
    Loop
    
    arr(idx) = Trim(current)
    
    SplitMarkdownTableRow = arr
End Function

Private Function CleanMarkdownTableCell(ByVal s As String) As String
    s = Trim(s)
    
    s = Replace(s, vbTab, " ")
    
    ' Convert common Markdown/HTML line breaks inside table cells
    s = Replace(s, "<br>", Chr(11))
    s = Replace(s, "<br/>", Chr(11))
    s = Replace(s, "<br />", Chr(11))
    
    CleanMarkdownTableCell = s
End Function

Private Function GetMarkdownTableAlignments(separatorCells As Variant, colCount As Long) As Variant
    Dim arr() As Long
    ReDim arr(0 To colCount - 1)
    
    Dim i As Long
    For i = 0 To colCount - 1
        If i <= UBound(separatorCells) Then
            arr(i) = MarkdownAlignmentFromSeparator(CStr(separatorCells(i)))
        Else
            arr(i) = wdAlignParagraphLeft
        End If
    Next i
    
    GetMarkdownTableAlignments = arr
End Function

Private Function MarkdownAlignmentFromSeparator(ByVal s As String) As Long
    Dim t As String
    t = Trim(s)
    t = Replace(t, " ", "")
    
    Dim leftColon As Boolean
    Dim rightColon As Boolean
    
    leftColon = (Left(t, 1) = ":")
    rightColon = (Right(t, 1) = ":")
    
    If leftColon And rightColon Then
        MarkdownAlignmentFromSeparator = wdAlignParagraphCenter
    ElseIf rightColon Then
        MarkdownAlignmentFromSeparator = wdAlignParagraphRight
    Else
        MarkdownAlignmentFromSeparator = wdAlignParagraphLeft
    End If
End Function

Private Function ArrayItemCount(arr As Variant) As Long
    ArrayItemCount = UBound(arr) - LBound(arr) + 1
End Function

' --------------------------------------------------------------------------
' LOGIC: Formatting
' --------------------------------------------------------------------------
Private Sub ProcessFormattingReverse(rng As Word.Range)
    Dim i As Long
    Dim para As Word.Paragraph
    Dim paraRng As Word.Range
    Dim txt As String
    Dim total As Long
    Dim doc As Word.Document
    
    Set doc = rng.Parent
    
    Dim reHead As Object
    Dim reQuote As Object
    Dim reListU As Object
    Dim reListO As Object
    Dim reTask As Object
    Dim reInline As Object
    Dim reLink As Object
    Dim reStrike As Object
    
    Set reHead = CreateObject("VBScript.RegExp")
    reHead.Pattern = "^(#{1,6})\s+"
    reHead.Global = False
    
    Set reQuote = CreateObject("VBScript.RegExp")
    reQuote.Pattern = "^\s*>\s*"
    reQuote.Global = False
    
    Set reTask = CreateObject("VBScript.RegExp")
    reTask.Pattern = "^(\s*)[-\*\+]\s+\[([ xX])\]\s+"
    reTask.Global = False
    
    Set reListU = CreateObject("VBScript.RegExp")
    reListU.Pattern = "^(\s*)([-\*\+])\s+"
    reListU.Global = False
    
    Set reListO = CreateObject("VBScript.RegExp")
    reListO.Pattern = "^(\s*)(\d+)\.\s+"
    reListO.Global = False
    
    Set reInline = CreateObject("VBScript.RegExp")
    reInline.Global = True
    
    Set reLink = CreateObject("VBScript.RegExp")
    reLink.Pattern = "\[(.+?)\]\((https?:\/\/[^\s\)]+)\)"
    reLink.Global = True
    
    Set reStrike = CreateObject("VBScript.RegExp")
    reStrike.Pattern = "~~(.+?)~~"
    reStrike.Global = True
    
    total = rng.Paragraphs.Count
    
    For i = total To 1 Step -1
        Set para = rng.Paragraphs(i)
        
        If Not IsParagraphCode(para) Then
            Set paraRng = para.Range
            txt = paraRng.Text
            
            If i Mod 50 = 0 Then doc.UndoClear
            
            If Len(txt) > 1 Then
                
                ' Block elements
                If InStr(1, txt, "#", vbBinaryCompare) > 0 Then
                    If Left(LTrim(txt), 1) = "#" Then ApplyHeading paraRng, reHead
                End If
                
                If InStr(1, txt, ">", vbBinaryCompare) > 0 Then
                    If Left(LTrim(txt), 1) = ">" Then ApplyQuote paraRng, reQuote
                End If
                
                txt = paraRng.Text
                
                If IsTaskListLine(txt, reTask) Then
                    ApplyTaskList paraRng, reTask
                ElseIf InStr(1, txt, "-", vbBinaryCompare) > 0 _
                    Or InStr(1, txt, "*", vbBinaryCompare) > 0 _
                    Or InStr(1, txt, "+", vbBinaryCompare) > 0 Then
                    ApplyListU paraRng, reListU
                ElseIf IsNumeric(Left(LTrim(txt), 1)) Then
                    ApplyListO paraRng, reListO
                End If
                
                ' Inline elements
                txt = paraRng.Text
                
                If InStr(1, txt, "`", vbBinaryCompare) > 0 Then RunRegexInlineCode paraRng, reInline
                If InStr(1, txt, "*", vbBinaryCompare) > 0 Or InStr(1, txt, "_", vbBinaryCompare) > 0 Then RunRegexInlineFormat paraRng, reInline
                If InStr(1, txt, "[", vbBinaryCompare) > 0 Then RunRegexLink paraRng, reLink
                If InStr(1, txt, "~", vbBinaryCompare) > 0 Then RunRegexStrike paraRng, reStrike
            End If
        End If
    Next i
End Sub

' --------------------------------------------------------------------------
' BLOCK FORMAT HELPERS
' --------------------------------------------------------------------------
Private Sub ApplyHeading(rng As Word.Range, re As Object)
    Dim m As Object
    Dim doc As Word.Document
    Dim lvl As Long
    
    Set doc = rng.Parent
    
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        lvl = Len(Trim(m.SubMatches(0)))
        If lvl > 6 Then lvl = 6
        
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.Style = doc.Styles("Heading " & lvl)
    End If
End Sub

Private Sub ApplyQuote(rng As Word.Range, re As Object)
    Dim m As Object
    Dim doc As Word.Document
    
    Set doc = rng.Parent
    
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.Style = doc.Styles("Quote")
    End If
End Sub

Private Function IsTaskListLine(ByVal txt As String, re As Object) As Boolean
    IsTaskListLine = re.Test(txt)
End Function

Private Sub ApplyTaskList(rng As Word.Range, re As Object)
    Dim m As Object
    Dim doc As Word.Document
    Dim ls As Long
    Dim checkedMark As String
    Dim k As Long
    
    Set doc = rng.Parent
    
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        
        ls = Len(m.SubMatches(0))
        
        If LCase(CStr(m.SubMatches(1))) = "x" Then
            checkedMark = ChrW(&H2611) & " "
        Else
            checkedMark = ChrW(&H2610) & " "
        End If
        
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.InsertBefore checkedMark
        
        rng.ListFormat.ApplyBulletDefault
        
        For k = 1 To (ls \ 2)
            rng.ListFormat.ListIndent
        Next k
    End If
End Sub

Private Sub ApplyListU(rng As Word.Range, re As Object)
    Dim m As Object
    Dim doc As Word.Document
    Dim ls As Long
    Dim k As Long
    
    Set doc = rng.Parent
    
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        ls = Len(m.SubMatches(0))
        
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.ListFormat.ApplyBulletDefault
        
        For k = 1 To (ls \ 2)
            rng.ListFormat.ListIndent
        Next k
    End If
End Sub

Private Sub ApplyListO(rng As Word.Range, re As Object)
    Dim m As Object
    Dim doc As Word.Document
    Dim ls As Long
    Dim k As Long
    
    Set doc = rng.Parent
    
    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        ls = Len(m.SubMatches(0))
        
        doc.Range(rng.Start, rng.Start + m.Length).Delete
        rng.ListFormat.ApplyNumberDefault
        
        For k = 1 To (ls \ 2)
            rng.ListFormat.ListIndent
        Next k
    End If
End Sub

' --------------------------------------------------------------------------
' INLINE FORMAT HELPERS
' --------------------------------------------------------------------------
Private Sub RunRegexInlineCode(rng As Word.Range, re As Object)
    Dim mCol As Object
    Dim m As Object
    Dim startP As Long
    Dim markLen As Long
    Dim rWhole As Word.Range
    Dim doc As Word.Document
    
    Set doc = rng.Parent
    
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
        rWhole.Font.Color = wdColorBlack
        rWhole.Shading.BackgroundPatternColor = wdColorGray15
        rWhole.NoProofing = True
    Loop
End Sub

Private Sub RunRegexInlineFormat(rng As Word.Range, re As Object)
    Dim patterns As Variant
    Dim types As Variant
    Dim i As Long
    
    patterns = Array("(\*\*\*|___)(.+?)\1", "(\*\*|__)(.+?)\1", "(\*|_)(.+?)\1")
    types = Array(0, 1, 2)
    
    For i = 0 To 2
        re.Pattern = patterns(i)
        ApplyFormatSimple rng, re, CLng(types(i))
    Next i
End Sub

Private Sub RunRegexLink(rng As Word.Range, re As Object)
    Dim mCol As Object
    Dim m As Object
    Dim startP As Long
    Dim rWhole As Word.Range
    Dim doc As Word.Document
    Dim txt As String
    Dim url As String
    
    Set doc = rng.Parent
    
    Do
        Set mCol = re.Execute(rng.Text)
        If mCol.Count = 0 Then Exit Do
        
        Set m = mCol(mCol.Count - 1)
        startP = rng.Start + m.FirstIndex
        
        Set rWhole = doc.Range(startP, startP + m.Length)
        
        txt = m.SubMatches(0)
        url = m.SubMatches(1)
        
        rWhole.Text = txt
        doc.Hyperlinks.Add Anchor:=rWhole, Address:=url, TextToDisplay:=txt
    Loop
End Sub

Private Sub RunRegexStrike(rng As Word.Range, re As Object)
    ApplyFormatSimple rng, re, 4
End Sub

Private Sub ApplyFormatSimple(rng As Word.Range, re As Object, fmtType As Long)
    Dim mCol As Object
    Dim m As Object
    Dim startP As Long
    Dim markLen As Long
    Dim rWhole As Word.Range
    Dim rMarker As Word.Range
    Dim doc As Word.Document
    
    Set doc = rng.Parent
    
    Do
        Set mCol = re.Execute(rng.Text)
        If mCol.Count = 0 Then Exit Do
        
        Set m = mCol(mCol.Count - 1)
        markLen = Len(m.SubMatches(0))
        startP = rng.Start + m.FirstIndex
        
        Set rWhole = doc.Range(startP, startP + m.Length)
        
        Set rMarker = doc.Range(rWhole.End - markLen, rWhole.End)
        rMarker.Delete
        
        Set rMarker = doc.Range(rWhole.Start, rWhole.Start + markLen)
        rMarker.Delete
        
        Select Case fmtType
            Case 0
                rWhole.Font.Bold = True
                rWhole.Font.Italic = True
            Case 1
                rWhole.Font.Bold = True
            Case 2
                rWhole.Font.Italic = True
            Case 4
                rWhole.Font.Strikethrough = True
        End Select
    Loop
End Sub

' --------------------------------------------------------------------------
' CLEANUP
' --------------------------------------------------------------------------
Private Sub RemoveEmptyParagraphs(rng As Word.Range)
    Dim i As Long
    Dim p As Word.Paragraph
    
    For i = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(i)
        
        If Len(p.Range.Text) <= 1 And Not IsParagraphCode(p) Then
            p.Range.Delete
        End If
    Next i
End Sub
