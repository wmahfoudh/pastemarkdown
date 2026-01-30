' --------------------------------------------------------------------------
' Markdown-to-Word VBA Macro
' --------------------------------------------------------------------------
' References required:
'   - Microsoft Forms 2.0 Object Library (FM20.DLL found in system32)
'   - Microsoft VBScript Regular Expressions 5.5 (vbscript.dll found in system32)
' --------------------------------------------------------------------------

Option Explicit

' Ensure a paragraph style named "Code" exists; create it if missing.
Private Sub EnsureCodeStyle()
    Dim sty As Style

    On Error Resume Next
    Set sty = ActiveDocument.Styles("Code")
    On Error GoTo 0

    If sty Is Nothing Then
        Set sty = ActiveDocument.Styles.Add(Name:="Code", Type:=wdStyleTypeParagraph)
        sty.Font.Name = "Consolas"
        sty.Shading.BackgroundPatternColor = wdColorGray15
    End If
End Sub

' Entry point: retrieves Markdown from clipboard and processes it.
Public Sub PasteMarkdown()
    Dim dataObj    As New MSForms.DataObject
    Dim clipText   As String
    Dim rng        As Range
    Dim startPos   As Long
    Dim endPos     As Long
    Dim para       As Paragraph

    ' Ensure the "Code" style exists before running.
    EnsureCodeStyle

    ' Retrieve text from clipboard.
    On Error Resume Next
    dataObj.GetFromClipboard
    clipText = dataObj.GetText(1)
    If Err.Number <> 0 Or Len(Trim$(clipText)) = 0 Then Exit Sub
    On Error GoTo 0

    ' Insert raw Markdown at current cursor position.
    startPos = Selection.Range.Start
    Selection.TypeText Text:=clipText
    endPos = Selection.Range.Start

    ' Define range covering the inserted Markdown.
    Set rng = ActiveDocument.Range(Start:=startPos, End:=endPos)

    ' Improve performance by disabling screen updating and pagination.
    With Application
        .ScreenUpdating = False
        .Options.Pagination = False
    End With

    ' Parse fenced code blocks before other block-level elements.
    ParseFencedCodeBlocks rng

    ' Parse block-level elements: blockquotes, headings, and lists.
    For Each para In rng.Paragraphs
        ParseBlockquotes para.Range
        ParseHeading     para.Range
        ParseList        para.Range
    Next para

    ' Parse inline-level elements.
    ParseInline        rng
    ParseInlineCode    rng
    ParseStrikethrough rng
    ParseLinks         rng

    ' Restore Word application settings.
    With Application
        .ScreenUpdating = True
        .Options.Pagination = True
    End With
End Sub

' Convert fenced code blocks (``` ... ```) to the "Code" style.
Private Sub ParseFencedCodeBlocks(rng As Range)
    Dim para         As Paragraph
    Dim inCode       As Boolean
    Dim codeStart    As Long
    Dim fencePattern As RegExp
    Dim txt          As String
    Dim codeRange    As Range

    Set fencePattern = New RegExp
    fencePattern.Pattern = "^```.*\r?$"
    fencePattern.Global  = False

    inCode = False
    For Each para In rng.Paragraphs
        txt = Trim(para.Range.Text)

        If Not inCode Then
            If fencePattern.Test(txt) Then
                inCode = True
                codeStart = para.Range.Start
                para.Range.Delete
            End If
        Else
            If fencePattern.Test(txt) Then
                Set codeRange = ActiveDocument.Range(Start:=codeStart, End:=para.Range.End)
                para.Range.Delete
                codeRange.Style = ActiveDocument.Styles("Code")
                inCode = False
            End If
        End If
    Next para
End Sub

' Convert Markdown blockquotes ("> text") to the built-in "Quote" style.
Private Sub ParseBlockquotes(rng As Range)
    Dim re As RegExp
    Dim m  As Match

    Set re = New RegExp
    re.Pattern = "^>\s*(.+)$"
    re.Global  = False

    If re.Test(rng.Text) Then
        Set m = re.Execute(rng.Text)(0)
        rng.Text = m.SubMatches(0)
        rng.Style = ActiveDocument.Styles("Quote")
    End If
End Sub

' Convert Markdown headings (# through ######) to Word Heading styles.
Private Sub ParseHeading(rng As Range)
    Dim re      As RegExp
    Dim mcol    As MatchCollection
    Dim m       As Match
    Dim lvl     As Long
    Dim txt     As String

    Set re = New RegExp
    re.Pattern = "^(#{1,6})\s*(.+)$"
    re.Global  = False

    If re.Test(rng.Text) Then
        Set mcol = re.Execute(rng.Text)
        Set m    = mcol(0)
        lvl      = Len(m.SubMatches(0))
        txt      = m.SubMatches(1)
        rng.Text = txt
        rng.Style = ActiveDocument.Styles("Heading " & lvl)
    End If
End Sub

' Convert Markdown unordered (*) and ordered (1.) lists to Word lists.
Private Sub ParseList(rng As Range)
    Dim reU     As RegExp
    Dim reO     As RegExp
    Dim m       As Match
    Dim spaces  As String
    Dim content As String
    Dim indent  As Long
    Dim i       As Long

    Set reU = New RegExp
    reU.Pattern = "^(\s*)([-\*\+])\s+(.+)$"
    reU.Global  = False

    If reU.Test(rng.Text) Then
        Set m       = reU.Execute(rng.Text)(0)
        spaces      = m.SubMatches(0)
        content     = m.SubMatches(2)
        rng.Text    = content
        rng.ListFormat.ApplyBulletDefault
        indent      = Len(spaces) \ 2 + 1
        For i = 2 To indent
            rng.ListFormat.ListIndent
        Next i
        Exit Sub
    End If

    Set reO = New RegExp
    reO.Pattern = "^(\s*)(\d+)\.\s+(.+)$"
    reO.Global  = False

    If reO.Test(rng.Text) Then
        Set m       = reO.Execute(rng.Text)(0)
        spaces      = m.SubMatches(0)
        content     = m.SubMatches(2)
        rng.Text    = content
        rng.ListFormat.ApplyNumberDefault
        indent      = Len(spaces) \ 2 + 1
        For i = 2 To indent
            rng.ListFormat.ListIndent
        Next i
    End If
End Sub

' Apply bold and italic formatting for ***, **, and * markers.
Private Sub ParseInline(rng As Range)
    Dim patterns As Variant
    Dim i        As Long

    ' List of regex patterns: bold+italic, bold, italic
    patterns = Array( _
        "(\*\*\*|___)(.+?)\1", _
        "(\*\*|__)(.+?)\1", _
        "(\*|_)(.+?)\1" _
    )

    For i = LBound(patterns) To UBound(patterns)
        ApplyInlineFormatting rng, CStr(patterns(i)), i
    Next i
End Sub

' Convert inline code spans (`code`) to Consolas with shading.
Private Sub ParseInlineCode(rng As Range)
    Dim re      As RegExp
    Dim mcol    As MatchCollection
    Dim m       As Match
    Dim marker  As String
    Dim markLen As Long
    Dim startP  As Long
    Dim endP    As Long
    Dim rWhole  As Range
    Dim rInner  As Range

    Set re = New RegExp
    re.Pattern = "(`+)(.+?)\1"
    re.Global  = True

    Do
        Set mcol = re.Execute(rng.Text)
        If mcol.Count = 0 Then Exit Do
        Set m       = mcol(mcol.Count - 1)
        marker      = m.SubMatches(0)
        markLen     = Len(marker)
        startP      = rng.Start + m.FirstIndex
        endP        = startP + m.Length
        Set rWhole  = ActiveDocument.Range(startP, endP)
        rWhole.Characters.Last.Delete markLen
        rWhole.Characters.First.Delete markLen
        Set rInner  = rWhole
        rInner.Font.Name = "Consolas"
        rInner.Shading.BackgroundPatternColor = wdColorGray15
    Loop
End Sub

' Convert ~~strikethrough~~ to struck text.
Private Sub ParseStrikethrough(rng As Range)
    Dim re      As RegExp
    Dim mcol    As MatchCollection
    Dim m       As Match
    Dim startP  As Long
    Dim endP    As Long
    Dim rWhole  As Range

    Set re = New RegExp
    re.Pattern = "~~(.+?)~~"
    re.Global  = True

    Do
        Set mcol = re.Execute(rng.Text)
        If mcol.Count = 0 Then Exit Do
        Set m      = mcol(mcol.Count - 1)
        startP     = rng.Start + m.FirstIndex
        endP       = startP + m.Length
        Set rWhole = ActiveDocument.Range(startP, endP)
        rWhole.Characters.Last.Delete 2
        rWhole.Characters.First.Delete 2
        rWhole.Font.StrikeThrough = True
    Loop
End Sub

' Convert [text](url) to active Word hyperlinks.
Private Sub ParseLinks(rng As Range)
    Dim re          As RegExp
    Dim mcol        As MatchCollection
    Dim m           As Match
    Dim startP      As Long
    Dim endP        As Long
    Dim rWhole      As Range
    Dim displayText As String
    Dim url         As String

    Set re = New RegExp
    re.Pattern = "\[(.+?)\]\((https?:\/\/[^\s\)]+)\)"
    re.Global  = True

    Do
        Set mcol = re.Execute(rng.Text)
        If mcol.Count = 0 Then Exit Do
        Set m      = mcol(mcol.Count - 1)
        startP     = rng.Start + m.FirstIndex
        endP       = startP + m.Length
        Set rWhole = ActiveDocument.Range(startP, endP)
        displayText = m.SubMatches(0)
        url         = m.SubMatches(1)
        rWhole.Text = displayText
        ActiveDocument.Hyperlinks.Add Anchor:=rWhole, Address:=url, TextToDisplay:=displayText
    Loop
End Sub

' Helper to apply bold and/or italic based on marker type.
Private Sub ApplyInlineFormatting( _
        rng As Range, _
        ByVal pattern As String, _
        ByVal formatType As Long)
    Dim re      As RegExp
    Dim mcol    As MatchCollection
    Dim m       As Match
    Dim marker  As String
    Dim markLen As Long
    Dim startP  As Long
    Dim endP    As Long
    Dim rWhole  As Range
    Dim rInner  As Range
    Dim rMarker As Range

    Set re = New RegExp
    re.Pattern = pattern
    re.Global  = True

    Do
        Set mcol = re.Execute(rng.Text)
        If mcol.Count = 0 Then Exit Do
        Set m       = mcol(mcol.Count - 1)
        marker      = m.SubMatches(0)
        markLen     = Len(marker)
        startP      = rng.Start + m.FirstIndex
        endP        = startP + m.Length
        Set rWhole  = ActiveDocument.Range(startP, endP)

        Set rMarker = ActiveDocument.Range(rWhole.End - markLen, rWhole.End)
        rMarker.Delete

        Set rMarker = ActiveDocument.Range(rWhole.Start, rWhole.Start + markLen)
        rMarker.Delete

        Set rInner = rWhole

        Select Case formatType
            Case 0
                rInner.Font.Bold   = True
                rInner.Font.Italic = True
            Case 1
                rInner.Font.Bold   = True
            Case 2
                rInner.Font.Italic = True
        End Select
    Loop
End Sub
