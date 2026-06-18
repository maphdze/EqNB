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

Public Sub InsertEquationLine(Optional ByVal mode As String = "plain", Optional ByVal separator As String = "-")
    On Error GoTo Failed

    Dim doc As Document
    Set doc = ActiveDocument

    If mode = "chapter" Then
        If Not HasNumberedHeadingOne(doc) Then
            MsgBox "章节编号需要文档中至少有一个使用“标题 1”且带自动编号的章节标题。", vbExclamation, "公式编号"
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
    cc.Title = "公式 " & bookmarkName
    cc.Tag = EQUATION_TAG
    cc.LockContentControl = False

    doc.Fields.Update
    equationRange.Select
    Exit Sub

Failed:
    MsgBox "插入公式行失败：" & Err.Description, vbCritical, "公式编号"
End Sub

Public Sub InsertEquationReference()
    On Error GoTo Failed

    Dim refs As Collection
    Set refs = GetEquationReferences()

    If refs.Count = 0 Then
        MsgBox "没有找到由本插件插入的公式。请先插入公式行。", vbInformation, "公式编号"
        Exit Sub
    End If

    Dim prompt As String
    prompt = "输入要引用的公式序号：" & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To refs.Count
        prompt = prompt & CStr(i) & ". " & refs(i)(1) & vbCrLf
    Next i

    Dim answer As String
    answer = InputBox(prompt, "插入公式引用", "1")
    If Len(Trim$(answer)) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then
        MsgBox "请输入列表中的数字序号。", vbExclamation, "公式编号"
        Exit Sub
    End If

    Dim index As Long
    index = CLng(answer)
    If index < 1 Or index > refs.Count Then
        MsgBox "序号超出范围。", vbExclamation, "公式编号"
        Exit Sub
    End If

    Selection.TypeText "公式 "
    ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldRef, Text:=refs(index)(0) & " \h", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    ActiveDocument.Fields.Update
    Exit Sub

Failed:
    MsgBox "插入公式引用失败：" & Err.Description, vbCritical, "公式编号"
End Sub

Public Sub RefreshEquationFields()
    On Error GoTo Failed
    ActiveDocument.Fields.Update
    MsgBox "已刷新文档中的编号和引用域。", vbInformation, "公式编号"
    Exit Sub

Failed:
    MsgBox "刷新失败：" & Err.Description, vbCritical, "公式编号"
End Sub

Public Sub ShowEquationNumberingHelp()
    MsgBox "公式编号宏：" & vbCrLf & _
        "1. 运行 InsertEquationLinePlain 插入纯流水号公式。" & vbCrLf & _
        "2. 运行 InsertEquationLineChapterHyphen/Dot/Colon 插入章节编号公式。" & vbCrLf & _
        "3. 运行 InsertEquationReference 插入正文引用。" & vbCrLf & _
        "4. 运行 RefreshEquationFields 刷新编号与引用。", vbInformation, "公式编号"
End Sub

Private Function GetEquationReferences() As Collection
    Dim refs As New Collection
    Dim cc As ContentControl

    For Each cc In ActiveDocument.ContentControls
        If cc.Tag = EQUATION_TAG Then
            Dim item(1) As String
            item(0) = Replace(cc.Title, "公式 ", "")
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
