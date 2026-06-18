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

    Dim insertStart As Long
    insertStart = Selection.Start
    Selection.Range.InsertXML BuildEquationLineXml(mode, separator, bookmarkName, PointsToTwips(centerTab), PointsToTwips(rightTab))

    doc.Fields.Update

    Dim insertedRange As Range
    Set insertedRange = doc.Range(insertStart, doc.Content.End)
    If insertedRange.OMaths.Count > 0 Then
        insertedRange.OMaths(1).Range.Select
    End If
    Exit Sub

Failed:
    MsgBox "Failed to insert equation line: " & Err.Description, vbCritical, "Equation Numbering"
End Sub

Private Function BuildEquationLineXml(ByVal mode As String, ByVal separator As String, ByVal bookmarkName As String, ByVal centerTab As Long, ByVal rightTab As Long) As String
    Dim bookmarkId As Long
    bookmarkId = CLng(Int(Rnd() * 2000000000))

    Dim captionXml As String
    captionXml = "<w:r><w:t>(</w:t></w:r>" & _
        "<w:bookmarkStart w:id=""" & CStr(bookmarkId) & """ w:name=""" & EscapeXml(bookmarkName) & """/>"

    If mode = "chapter" Then
        captionXml = captionXml & FieldXml("STYLEREF 1 \s") & _
            "<w:r><w:t>" & EscapeXml(separator) & "</w:t></w:r>"
    End If

    captionXml = captionXml & FieldXml("SEQ " & EQUATION_SEQ_NAME & " \* ARABIC") & _
        "<w:bookmarkEnd w:id=""" & CStr(bookmarkId) & """/>" & _
        "<w:r><w:t>)</w:t></w:r>"

    BuildEquationLineXml = "<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" " & _
        "xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"">" & _
        "<w:pPr><w:tabs>" & _
        "<w:tab w:val=""center"" w:pos=""" & CStr(centerTab) & """/>" & _
        "<w:tab w:val=""right"" w:pos=""" & CStr(rightTab) & """/>" & _
        "</w:tabs><w:spacing w:before=""120"" w:after=""120""/>" & _
        "<w:ind w:left=""0"" w:right=""0"" w:firstLine=""0""/></w:pPr>" & _
        "<w:r><w:tab/></w:r>" & _
        EmptyEquationXml() & _
        "<w:r><w:tab/></w:r>" & _
        captionXml & _
        "</w:p>"
End Function

Private Function EmptyEquationXml() As String
    EmptyEquationXml = "<m:oMath><m:r><w:rPr><w:rFonts w:ascii=""Cambria Math"" w:hAnsi=""Cambria Math""/></w:rPr><m:t>" & ChrW(&H25A1) & "</m:t></m:r></m:oMath>"
End Function

Private Function FieldXml(ByVal instruction As String) As String
    FieldXml = "<w:fldSimple w:instr=""" & EscapeXml(instruction) & """><w:r><w:t>?</w:t></w:r></w:fldSimple>"
End Function

Private Function EscapeXml(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, """", "&quot;")
    EscapeXml = value
End Function

Private Function PointsToTwips(ByVal value As Single) As Long
    PointsToTwips = CLng(value * 20)
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
    Dim bookmark As Bookmark

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

    Set GetEquationReferences = refs
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
