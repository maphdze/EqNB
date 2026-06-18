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
    MsgBox UiText("error.inline") & Err.Description, vbCritical, APP_TITLE
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

Public Sub RibbonGetLabel(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = UiText(control.Id & ".label")
End Sub

Public Sub RibbonGetScreentip(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = UiText(control.Id & ".screentip")
End Sub

Public Sub RibbonGetSupertip(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = UiText(control.Id & ".supertip")
End Sub

Public Sub InsertEquationLine(Optional ByVal mode As String = "plain", Optional ByVal separator As String = "-")
    On Error GoTo Failed

    Dim doc As Document
    Set doc = ActiveDocument

    If mode = "chapter" Then
        If Not HasNumberedHeadingOne(doc) Then
            MsgBox UiText("error.chapterHeading"), vbExclamation, APP_TITLE
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
    MsgBox UiText("error.equationLine") & Err.Description, vbCritical, APP_TITLE
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
        MsgBox UiText("error.noEquations"), vbInformation, APP_TITLE
        Exit Sub
    End If

    Dim prompt As String
    prompt = UiText("reference.prompt") & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To refs.Count
        prompt = prompt & CStr(i) & ". " & refs(i)(1) & vbCrLf
    Next i

    Dim answer As String
    answer = InputBox(prompt, UiText("reference.title"), "1")
    If Len(Trim$(answer)) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then
        MsgBox UiText("error.referenceNumber"), vbExclamation, APP_TITLE
        Exit Sub
    End If

    Dim index As Long
    index = CLng(answer)
    If index < 1 Or index > refs.Count Then
        MsgBox UiText("error.referenceRange"), vbExclamation, APP_TITLE
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
    MsgBox UiText("error.reference") & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub SetEquationReferenceFormat()
    On Error GoTo Failed

    Dim formatText As String
    formatText = InputBox(UiText("format.prompt"), UiText("format.title"), GetEquationReferenceFormat(ActiveDocument))
    If Len(formatText) = 0 Then Exit Sub
    If InStr(formatText, "{n}") = 0 Then
        MsgBox UiText("error.formatMarker"), vbExclamation, APP_TITLE
        Exit Sub
    End If

    SetDocumentVariable ActiveDocument, REFERENCE_FORMAT_VARIABLE, formatText
    MsgBox UiText("format.saved") & formatText, vbInformation, APP_TITLE
    Exit Sub

Failed:
    MsgBox UiText("error.format") & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub RefreshEquationFields()
    On Error GoTo Failed
    ActiveDocument.Fields.Update
    MsgBox UiText("refresh.done"), vbInformation, APP_TITLE
    Exit Sub

Failed:
    MsgBox UiText("error.refresh") & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub ShowEquationNumberingHelp()
    MsgBox UiText("help.body"), vbInformation, APP_TITLE
End Sub

Public Function EquationNumberingSmokeTest() As String
    EquationNumberingSmokeTest = "OK"
End Function

Private Function UiText(ByVal key As String) As String
    If IsSimplifiedChineseUi() Then
        UiText = UiTextZhCn(key)
    Else
        UiText = UiTextEn(key)
    End If

    If Len(UiText) = 0 Then UiText = key
End Function

Private Function IsSimplifiedChineseUi() As Boolean
    On Error GoTo UseEnglish

    IsSimplifiedChineseUi = (Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 2052)
    Exit Function

UseEnglish:
    IsSimplifiedChineseUi = False
End Function

Private Function UiTextEn(ByVal key As String) As String
    Select Case key
        Case "EquationNumberingTab.label": UiTextEn = APP_TITLE
        Case "EquationInsertGroup.label": UiTextEn = "Insert"
        Case "EquationReferenceGroup.label": UiTextEn = "Reference"
        Case "EquationHelpGroup.label": UiTextEn = "Help"
        Case "InsertPlainEquation.label": UiTextEn = "Numbered"
        Case "InsertPlainEquation.screentip": UiTextEn = "Insert numbered display equation"
        Case "InsertPlainEquation.supertip": UiTextEn = "Insert a centered Word equation with a right-aligned sequential number."
        Case "InsertInlineEquation.label": UiTextEn = "Inline"
        Case "InsertInlineEquation.screentip": UiTextEn = "Insert inline equation"
        Case "InsertInlineEquation.supertip": UiTextEn = "Insert a Word inline equation without numbering or display style."
        Case "InsertChapterEquation.label": UiTextEn = "Chapter No."
        Case "InsertChapterEquation.screentip": UiTextEn = "Insert chapter-numbered equation"
        Case "InsertChapterEquation.supertip": UiTextEn = "Use Heading 1 chapter numbering plus equation sequence, such as 2-3."
        Case "InsertChapterDotEquation.label": UiTextEn = "Chapter.No."
        Case "InsertChapterDotEquation.screentip": UiTextEn = "Insert dot-separated chapter-numbered equation"
        Case "InsertChapterDotEquation.supertip": UiTextEn = "Use Heading 1 chapter numbering plus equation sequence, such as 2.3."
        Case "InsertEquationReference.label": UiTextEn = "Insert Ref"
        Case "InsertEquationReference.screentip": UiTextEn = "Insert equation cross-reference"
        Case "InsertEquationReference.supertip": UiTextEn = "Choose an equation inserted by EqNB and insert a refreshable REF field."
        Case "SetEquationReferenceFormat.label": UiTextEn = "Ref Format"
        Case "SetEquationReferenceFormat.screentip": UiTextEn = "Set equation reference format"
        Case "SetEquationReferenceFormat.supertip": UiTextEn = "Set the document-wide reference format, such as ({n}), Equation ({n}), Eq.({n}), or [{n}]."
        Case "RefreshEquationFields.label": UiTextEn = "Refresh"
        Case "RefreshEquationFields.screentip": UiTextEn = "Refresh numbers and references"
        Case "RefreshEquationFields.supertip": UiTextEn = "Refresh equation numbers and reference fields in the document."
        Case "EquationNumberingHelp.label": UiTextEn = "About"
        Case "reference.title": UiTextEn = "Insert Equation Reference"
        Case "reference.prompt": UiTextEn = "Enter the number of the equation to reference:"
        Case "format.title": UiTextEn = "Reference Format"
        Case "format.prompt": UiTextEn = "Set the reference format for this document. Use {n} where the equation number should appear." & vbCrLf & "Examples: ({n}), Equation ({n}), Eq.({n}), [{n}]"
        Case "format.saved": UiTextEn = "Reference format saved for this document: "
        Case "refresh.done": UiTextEn = "Equation numbers and references were refreshed."
        Case "help.body": UiTextEn = "EqNB macros:" & vbCrLf & _
            "1. Insert numbered display equations." & vbCrLf & _
            "2. Insert inline equations without numbering." & vbCrLf & _
            "3. Set one reference format for the document." & vbCrLf & _
            "4. Insert cross-references and refresh fields."
        Case "error.inline": UiTextEn = "Failed to insert inline equation: "
        Case "error.chapterHeading": UiTextEn = "Chapter numbering requires at least one numbered Heading 1 paragraph."
        Case "error.equationLine": UiTextEn = "Failed to insert equation line: "
        Case "error.noEquations": UiTextEn = "No equations inserted by EqNB were found."
        Case "error.referenceNumber": UiTextEn = "Please enter a number from the list."
        Case "error.referenceRange": UiTextEn = "The selected number is out of range."
        Case "error.reference": UiTextEn = "Failed to insert equation reference: "
        Case "error.formatMarker": UiTextEn = "The format must contain {n}."
        Case "error.format": UiTextEn = "Failed to set reference format: "
        Case "error.refresh": UiTextEn = "Failed to refresh fields: "
    End Select
End Function

Private Function UiTextZhCn(ByVal key As String) As String
    Select Case key
        Case "EquationNumberingTab.label": UiTextZhCn = APP_TITLE
        Case "EquationInsertGroup.label": UiTextZhCn = "插入"
        Case "EquationReferenceGroup.label": UiTextZhCn = "引用"
        Case "EquationHelpGroup.label": UiTextZhCn = "帮助"
        Case "InsertPlainEquation.label": UiTextZhCn = "编号公式"
        Case "InsertPlainEquation.screentip": UiTextZhCn = "插入编号行间公式"
        Case "InsertPlainEquation.supertip": UiTextZhCn = "插入一个居中的 Word 原生公式，并在右侧添加顺序编号。"
        Case "InsertInlineEquation.label": UiTextZhCn = "行内公式"
        Case "InsertInlineEquation.screentip": UiTextZhCn = "插入行内公式"
        Case "InsertInlineEquation.supertip": UiTextZhCn = "插入一个 Word 原生行内公式，不添加编号，不使用 display style。"
        Case "InsertChapterEquation.label": UiTextZhCn = "章节编号"
        Case "InsertChapterEquation.screentip": UiTextZhCn = "插入章节编号公式"
        Case "InsertChapterEquation.supertip": UiTextZhCn = "使用标题 1 章节号加公式流水号，例如 2-3。"
        Case "InsertChapterDotEquation.label": UiTextZhCn = "章节.编号"
        Case "InsertChapterDotEquation.screentip": UiTextZhCn = "插入句点分隔的章节编号公式"
        Case "InsertChapterDotEquation.supertip": UiTextZhCn = "使用标题 1 章节号加公式流水号，例如 2.3。"
        Case "InsertEquationReference.label": UiTextZhCn = "插入引用"
        Case "InsertEquationReference.screentip": UiTextZhCn = "插入公式交叉引用"
        Case "InsertEquationReference.supertip": UiTextZhCn = "从 EqNB 插入的公式中选择一个，在正文插入可刷新的 REF 引用。"
        Case "SetEquationReferenceFormat.label": UiTextZhCn = "引用格式"
        Case "SetEquationReferenceFormat.screentip": UiTextZhCn = "设置公式引用格式"
        Case "SetEquationReferenceFormat.supertip": UiTextZhCn = "为当前文档设置统一引用格式，例如 ({n})、式({n})、Eq.({n}) 或 [{n}]。"
        Case "RefreshEquationFields.label": UiTextZhCn = "刷新"
        Case "RefreshEquationFields.screentip": UiTextZhCn = "刷新编号和引用"
        Case "RefreshEquationFields.supertip": UiTextZhCn = "刷新文档中的公式编号和正文引用域。"
        Case "EquationNumberingHelp.label": UiTextZhCn = "说明"
        Case "reference.title": UiTextZhCn = "插入公式引用"
        Case "reference.prompt": UiTextZhCn = "请输入要引用的公式序号："
        Case "format.title": UiTextZhCn = "引用格式"
        Case "format.prompt": UiTextZhCn = "设置当前文档的统一引用格式。用 {n} 表示公式编号位置。" & vbCrLf & "示例：({n})、式({n})、Eq.({n})、[{n}]"
        Case "format.saved": UiTextZhCn = "已保存当前文档的引用格式："
        Case "refresh.done": UiTextZhCn = "公式编号和引用已刷新。"
        Case "help.body": UiTextZhCn = "EqNB 宏：" & vbCrLf & _
            "1. 插入带右侧编号的行间公式。" & vbCrLf & _
            "2. 插入不编号的行内公式。" & vbCrLf & _
            "3. 为当前文档设置统一引用格式。" & vbCrLf & _
            "4. 插入交叉引用并刷新域。"
        Case "error.inline": UiTextZhCn = "插入行内公式失败："
        Case "error.chapterHeading": UiTextZhCn = "章节编号需要至少一个已编号的“标题 1”段落。"
        Case "error.equationLine": UiTextZhCn = "插入编号公式失败："
        Case "error.noEquations": UiTextZhCn = "没有找到由 EqNB 插入的公式。"
        Case "error.referenceNumber": UiTextZhCn = "请输入列表中的数字。"
        Case "error.referenceRange": UiTextZhCn = "选择的序号超出范围。"
        Case "error.reference": UiTextZhCn = "插入公式引用失败："
        Case "error.formatMarker": UiTextZhCn = "格式中必须包含 {n}。"
        Case "error.format": UiTextZhCn = "设置引用格式失败："
        Case "error.refresh": UiTextZhCn = "刷新域失败："
    End Select
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
