Attribute VB_Name = "EquationNumbering"
Option Explicit

Private Const EQUATION_TAG As String = "WordEquationNumbering.Equation"
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

    Selection.TypeParagraph
    Dim paragraphRange As Range
    Set paragraphRange = Selection.Paragraphs(1).Range
    paragraphRange.ParagraphFormat.TabStops.ClearAll
    paragraphRange.ParagraphFormat.TabStops.Add Position:=contentWidth / 2, Alignment:=wdAlignTabCenter
    paragraphRange.ParagraphFormat.TabStops.Add Position:=contentWidth, Alignment:=wdAlignTabRight
    paragraphRange.ParagraphFormat.SpaceBefore = 6
    paragraphRange.ParagraphFormat.SpaceAfter = 6

    Selection.TypeText vbTab

    Dim equationStart As Long
    equationStart = Selection.Start
    Selection.TypeText ChrW(&H25A1)

    Dim equationRange As Range
    Set equationRange = doc.Range(equationStart, Selection.End)
    doc.OMaths.Add equationRange
    equationRange.OMaths(1).BuildUp
    equationRange.Select
    Selection.Collapse wdCollapseEnd

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

    Dim bookmarkName As String
    bookmarkName = CreateEquationBookmarkName()
    doc.Bookmarks.Add Name:=bookmarkName, Range:=doc.Range(captionStart, captionEnd)

    Set paragraphRange = Selection.Paragraphs(1).Range
    Dim cc As ContentControl
    Set cc = doc.ContentControls.Add(wdContentControlRichText, paragraphRange)
    cc.Title = "Equation " & bookmarkName
    cc.Tag = EQUATION_TAG
    cc.LockContentControl = False

    doc.Fields.Update
    equationRange.Select
    Exit Sub

Failed:
    MsgBox "Failed to insert equation line: " & Err.Description, vbCritical, "Equation Numbering"
End Sub

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

    Selection.TypeText "Equation "
    ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldRef, Text:=refs(index)(0) & " \h", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
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
    Dim refs As New Collection
    Dim cc As ContentControl

    For Each cc In ActiveDocument.ContentControls
        If cc.Tag = EQUATION_TAG Then
            Dim item(1) As String
            item(0) = Replace(cc.Title, "Equation ", "")
            item(1) = Trim$(Replace(cc.Range.Text, ChrW(13), ""))
            If Len(item(1)) = 0 Then item(1) = cc.Title
            refs.Add item
        End If
    Next cc

    Set GetEquationReferences = refs
End Function

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
