Attribute VB_Name = "EquationNumbering"
Option Explicit

Private Const EQUATION_SEQ_NAME As String = "Equation"
Private Const REFERENCE_FORMAT_VARIABLE As String = "EquationReferenceFormat"
Private Const DEFAULT_REFERENCE_FORMAT As String = "({n})"
Private Const APP_TITLE As String = "EqNB"

Public Sub InsertEquationLinePlain()
    InsertEquationLine "plain", "-"
End Sub

Public Sub InsertInlineEquation()
    On Error GoTo Failed

    Dim doc As Document
    Set doc = ActiveDocument

    Dim equationStart As Long
    Dim equationEnd As Long
    equationStart = Selection.Start
    Selection.TypeText ChrW(&H25A1)
    equationEnd = Selection.End

    Dim equationRange As Range
    Set equationRange = doc.Range(equationStart, equationEnd)

    doc.OMaths.Add equationRange
    Set equationRange = doc.Range(equationStart, equationEnd)

    Dim equationMath As OMath
    Set equationMath = FindEquationAt(doc, equationStart, equationEnd)
    If Not equationMath Is Nothing Then
        On Error Resume Next
        equationMath.Type = wdOMathInline
        On Error GoTo Failed
        Set equationRange = equationMath.Range
    End If

    equationRange.Select
    Exit Sub

Failed:
    MsgBox "Failed to insert inline equation: " & Err.Description, vbCritical, APP_TITLE
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

Public Sub RibbonInsertInlineEquation(ByVal control As IRibbonControl)
    InsertInlineEquation
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

Public Sub RibbonSetReferenceFormat(ByVal control As IRibbonControl)
    SetEquationReferenceFormat
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
            MsgBox "Chapter numbering requires at least one numbered Heading 1 paragraph.", vbExclamation, APP_TITLE
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
    Dim equationMath As OMath
    Set equationMath = FindEquationAt(doc, equationStart, equationEnd)
    If Not equationMath Is Nothing Then
        On Error Resume Next
        equationMath.Type = wdOMathDisplay
        On Error GoTo Failed
        Set equationRange = equationMath.Range
    Else
        Set equationRange = doc.Range(equationStart, equationEnd)
    End If

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
    MsgBox "Failed to insert equation line: " & Err.Description, vbCritical, APP_TITLE
End Sub

Private Function FindEquationAt(ByVal doc As Document, ByVal rangeStart As Long, ByVal rangeEnd As Long) As OMath
    Dim equation As OMath

    For Each equation In doc.OMaths
        If equation.Range.Start <= rangeStart And equation.Range.End >= rangeEnd Then
            Set FindEquationAt = equation
            Exit For
        End If
    Next equation
End Function

Public Sub InsertEquationReference()
    On Error GoTo Failed

    Dim refs As Collection
    Set refs = GetEquationReferences()

    If refs.Count = 0 Then
        MsgBox "No equations inserted by EqNB were found.", vbInformation, APP_TITLE
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
        MsgBox "Please enter a number from the list.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    Dim index As Long
    index = CLng(answer)
    If index < 1 Or index > refs.Count Then
        MsgBox "The selected number is out of range.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    Dim formatText As String
    formatText = GetEquationReferenceFormat(ActiveDocument)

    Dim markerPosition As Long
    markerPosition = InStr(formatText, "{n}")

    Selection.TypeText Left$(formatText, markerPosition - 1)
    ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldRef, Text:=refs(index)(0) & " \h", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    Selection.TypeText Mid$(formatText, markerPosition + 3)
    ActiveDocument.Fields.Update
    Exit Sub

Failed:
    MsgBox "Failed to insert equation reference: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub SetEquationReferenceFormat()
    On Error GoTo Failed

    Dim formatText As String
    formatText = InputBox("Set the reference format for this document. Use {n} where the equation number should appear." & vbCrLf & "Examples: ({n}), Equation ({n}), Eq.({n}), [{n}]", "Reference Format", GetEquationReferenceFormat(ActiveDocument))
    If Len(formatText) = 0 Then Exit Sub
    If InStr(formatText, "{n}") = 0 Then
        MsgBox "The format must contain {n}.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    SetDocumentVariable ActiveDocument, REFERENCE_FORMAT_VARIABLE, formatText
    MsgBox "Reference format saved for this document: " & formatText, vbInformation, APP_TITLE
    Exit Sub

Failed:
    MsgBox "Failed to set reference format: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub RefreshEquationFields()
    On Error GoTo Failed
    ActiveDocument.Fields.Update
    MsgBox "Equation numbers and references were refreshed.", vbInformation, APP_TITLE
    Exit Sub

Failed:
    MsgBox "Failed to refresh fields: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub ShowEquationNumberingHelp()
    MsgBox "EqNB macros:" & vbCrLf & _
        "1. Insert numbered display equations." & vbCrLf & _
        "2. Insert inline equations without numbering." & vbCrLf & _
        "3. Set one reference format for the document." & vbCrLf & _
        "4. Insert cross-references and refresh fields.", vbInformation, APP_TITLE
End Sub

Public Function EquationNumberingSmokeTest() As String
    EquationNumberingSmokeTest = "OK"
End Function

Private Function GetEquationReferenceFormat(ByVal doc As Document) As String
    On Error GoTo UseDefault

    GetEquationReferenceFormat = doc.Variables(REFERENCE_FORMAT_VARIABLE).Value
    If Len(GetEquationReferenceFormat) = 0 Then GoTo UseDefault
    If InStr(GetEquationReferenceFormat, "{n}") = 0 Then GoTo UseDefault
    Exit Function

UseDefault:
    GetEquationReferenceFormat = DEFAULT_REFERENCE_FORMAT
End Function

Private Sub SetDocumentVariable(ByVal doc As Document, ByVal variableName As String, ByVal variableValue As String)
    On Error GoTo AddVariable

    doc.Variables(variableName).Value = variableValue
    Exit Sub

AddVariable:
    doc.Variables.Add Name:=variableName, Value:=variableValue
End Sub

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
