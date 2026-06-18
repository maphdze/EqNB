Attribute VB_Name = "EquationNumbering"
Option Explicit

Private Const EQUATION_SEQ_NAME As String = "Equation"

Public Sub InsertEquationLinePlain()
    InsertEquationLine "plain", "-"
End Sub

Public Sub InsertEquationLineChapterHyphen()
    InsertEquationLine "chapter", "-"
End Sub

Public Sub InsertEquationLineChapterDot()
    InsertEquationLine "chapter", "."
End Sub

Public Sub InsertEquationLineChapterColon()
    InsertEquationLine "chapter", ":"
End Sub

Public Sub RibbonInsertPlain(ByVal control As IRibbonControl)
    InsertEquationLinePlain
End Sub

Public Sub RibbonInsertChapterHyphen(ByVal control As IRibbonControl)
    InsertEquationLineChapterHyphen
End Sub

Public Sub RibbonInsertChapterDot(ByVal control As IRibbonControl)
    InsertEquationLineChapterDot
End Sub

Public Sub RibbonInsertReference(ByVal control As IRibbonControl)
    InsertEquationReference
End Sub

Public Sub RibbonRefreshFields(ByVal control As IRibbonControl)
    RefreshEquationFields
End Sub

Public Sub RibbonShowHelp(ByVal control As IRibbonControl)
    ShowEquationNumberingHelp
End Sub

Public Sub InsertEquationLine(Optional ByVal mode As String = "plain", Optional ByVal separator As String = "-")
    On Error GoTo Failed

    Dim doc As Document
    Set doc = ActiveDocument

    If mode = "chapter" Then
        If Not HasNumberedHeadingOne(doc) Then
            MsgBox "Chapter numbering requires at least one numbered Heading 1 paragraph.", vbExclamation, "Equation Numbering"
            Exit Sub
        End If
    End If

    Dim contentWidth As Single
    contentWidth = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
    If contentWidth <= 0 Then contentWidth = InchesToPoints(6)

    Dim centerTab As Single
    Dim rightTab As Single
    centerTab = contentWidth / 2
    rightTab = contentWidth

    Dim bookmarkName As String
    bookmarkName = CreateEquationBookmarkName()

    Selection.TypeParagraph
    Dim paragraphRange As Range
    Set paragraphRange = Selection.Paragraphs(1).Range
    paragraphRange.ParagraphFormat.LeftIndent = 0
    paragraphRange.ParagraphFormat.RightIndent = 0
    paragraphRange.ParagraphFormat.FirstLineIndent = 0
    paragraphRange.ParagraphFormat.TabStops.ClearAll
    paragraphRange.ParagraphFormat.TabStops.Add Position:=centerTab, Alignment:=wdAlignTabCenter
    paragraphRange.ParagraphFormat.TabStops.Add Position:=rightTab, Alignment:=wdAlignTabRight
    paragraphRange.ParagraphFormat.Alignment = wdAlignParagraphLeft
    paragraphRange.ParagraphFormat.SpaceBefore = 6
    paragraphRange.ParagraphFormat.SpaceAfter = 6

    Selection.TypeText vbTab

    Dim equationStart As Long
    Dim equationEnd As Long
    equationStart = Selection.Start
    Selection.TypeText ChrW(&H25A1)
    equationEnd = Selection.End

    Dim equationRange As Range
    Set equationRange = doc.Range(equationStart, equationEnd)

    doc.OMaths.Add equationRange
    Set equationRange = FindEquationRangeAt(doc, equationStart, equationEnd)

    Selection.TypeText vbTab & "("

    Dim captionStart As Long
    captionStart = Selection.Start

    If mode = "chapter" Then
        doc.Fields.Add Range:=Selection.Range, Type:=wdFieldStyleRef, Text:="1 \s", PreserveFormatting:=False
        Selection.Collapse wdCollapseEnd
        Selection.TypeText separator
    End If

    doc.Fields.Add Range:=Selection.Range, Type:=wdFieldSequence, Text:=EQUATION_SEQ_NAME & " \* ARABIC", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd

    Dim captionEnd As Long
    captionEnd = Selection.End
    Selection.TypeText ")"

    doc.Bookmarks.Add Name:=bookmarkName, Range:=doc.Range(captionStart, captionEnd)

    doc.Fields.Update
    equationRange.Select
    Exit Sub

Failed:
    MsgBox "Failed to insert equation line: " & Err.Description, vbCritical, "Equation Numbering"
End Sub

Private Function FindEquationRangeAt(ByVal doc As Document, ByVal rangeStart As Long, ByVal rangeEnd As Long) As Range
    Dim equation As OMath
    Dim bestRange As Range

    For Each equation In doc.OMaths
        If equation.Range.Start <= rangeStart And equation.Range.End >= rangeEnd Then
            Set bestRange = equation.Range
            Exit For
        End If
    Next equation

    If bestRange Is Nothing Then
        Set bestRange = doc.Range(rangeStart, rangeEnd)
    End If

    Set FindEquationRangeAt = bestRange
End Function

Public Sub InsertEquationReference()
    On Error GoTo Failed

    Dim refs As Collection
    Set refs = GetEquationReferences()

    If refs.Count = 0 Then
        MsgBox "No equations inserted by this template were found.", vbInformation, "Equation Numbering"
        Exit Sub
    End If

    Dim prompt As String
    prompt = "Enter the number of the equation to reference:" & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To refs.Count
        prompt = prompt & CStr(i) & ". " & refs(i)(1) & vbCrLf
    Next i

    Dim answer As String
    answer = InputBox(prompt, "Insert Equation Reference", "1")
    If Len(Trim$(answer)) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then
        MsgBox "Please enter a number from the list.", vbExclamation, "Equation Numbering"
        Exit Sub
    End If

    Dim index As Long
    index = CLng(answer)
    If index < 1 Or index > refs.Count Then
        MsgBox "The selected number is out of range.", vbExclamation, "Equation Numbering"
        Exit Sub
    End If

    Dim formatText As String
    formatText = InputBox("Enter reference format. Use {n} where the equation number should appear." & vbCrLf & _
        "Examples: ({n}), Equation ({n}), Eq.({n}), [{n}]", "Reference Format", "({n})")
    If Len(formatText) = 0 Then Exit Sub
    If InStr(formatText, "{n}") = 0 Then formatText = formatText & "{n}"

    Dim markerPosition As Long
    markerPosition = InStr(formatText, "{n}")

    Selection.TypeText Left$(formatText, markerPosition - 1)
    ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldRef, Text:=refs(index)(0) & " \h", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    Selection.TypeText Mid$(formatText, markerPosition + 3)
    ActiveDocument.Fields.Update
    Exit Sub

Failed:
    MsgBox "Failed to insert equation reference: " & Err.Description, vbCritical, "Equation Numbering"
End Sub

Public Sub RefreshEquationFields()
    On Error GoTo Failed
    ActiveDocument.Fields.Update
    MsgBox "Equation numbers and references were refreshed.", vbInformation, "Equation Numbering"
    Exit Sub

Failed:
    MsgBox "Failed to refresh fields: " & Err.Description, vbCritical, "Equation Numbering"
End Sub

Public Sub ShowEquationNumberingHelp()
    MsgBox "Equation Numbering macros:" & vbCrLf & _
        "1. Run InsertEquationLinePlain for plain equation numbering." & vbCrLf & _
        "2. Run InsertEquationLineChapterHyphen/Dot/Colon for chapter numbering." & vbCrLf & _
        "3. Run InsertEquationReference to insert a cross reference." & vbCrLf & _
        "4. Run RefreshEquationFields to update numbers and references.", vbInformation, "Equation Numbering"
End Sub

Public Function EquationNumberingSmokeTest() As String
    EquationNumberingSmokeTest = "OK"
End Function

Private Function GetEquationReferences() As Collection
    On Error GoTo RestoreHiddenBookmarks

    Dim refs As New Collection
    Dim bookmark As Bookmark
    Dim showHiddenBefore As Boolean

    showHiddenBefore = ActiveDocument.Bookmarks.ShowHidden
    ActiveDocument.Bookmarks.ShowHidden = True

    For Each bookmark In ActiveDocument.Bookmarks
        If Left$(bookmark.Name, 4) = "_Eqn" Then
            Dim item(2) As String
            item(0) = bookmark.Name
            item(1) = Trim$(Replace(bookmark.Range.Text, ChrW(13), ""))
            item(2) = CStr(bookmark.Range.Start)
            If Len(item(1)) = 0 Then item(1) = bookmark.Name
            AddReferenceInDocumentOrder refs, item
        End If
    Next bookmark

    ActiveDocument.Bookmarks.ShowHidden = showHiddenBefore

    Set GetEquationReferences = refs
    Exit Function

RestoreHiddenBookmarks:
    ActiveDocument.Bookmarks.ShowHidden = showHiddenBefore
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub AddReferenceInDocumentOrder(ByRef refs As Collection, ByRef item() As String)
    Dim i As Long

    For i = 1 To refs.Count
        If CLng(item(2)) < CLng(refs(i)(2)) Then
            refs.Add item, Before:=i
            Exit Sub
        End If
    Next i

    refs.Add item
End Sub

Private Function CreateEquationBookmarkName() As String
    Randomize
    CreateEquationBookmarkName = "_Eqn" & Format$(Now, "yymmddhhnnss") & CStr(Int(Rnd() * 1000))
End Function

Private Function HasNumberedHeadingOne(ByVal doc As Document) As Boolean
    Dim paragraph As Paragraph

    For Each paragraph In doc.Paragraphs
        If paragraph.Style = doc.Styles(wdStyleHeading1) Then
            If paragraph.Range.ListFormat.ListType <> wdListNoNumbering Then
                HasNumberedHeadingOne = True
                Exit Function
            End If
        End If
    Next paragraph

    HasNumberedHeadingOne = False
End Function
