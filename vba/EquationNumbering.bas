Attribute VB_Name = "EquationNumbering"
Option Explicit

Private Const EQUATION_SEQ_NAME As String = "Equation"
Private Const REFERENCE_FORMAT_VARIABLE As String = "EquationReferenceFormat"
Private Const NUMBER_MODE_VARIABLE As String = "EqNBNumberMode"
Private Const CHAPTER_NUMBER_FORMAT_VARIABLE As String = "EqNBChapterNumberFormat"
Private Const EQUATION_NUMBER_FORMAT_VARIABLE As String = "EqNBEquationNumberFormat"
Private Const EQUATION_DELIMITER_VARIABLE As String = "EqNBEquationDelimiter"
Private Const EQUATION_SEPARATOR_VARIABLE As String = "EqNBEquationSeparator"
Private Const DEFAULT_REFERENCE_FORMAT As String = "({n})"
Private Const DEFAULT_NUMBER_MODE As String = "plain"
Private Const DEFAULT_NUMBER_FORMAT As String = "ARABIC"
Private Const DEFAULT_DELIMITER As String = "paren"
Private Const DEFAULT_SEPARATOR As String = "-"
Private Const APP_TITLE As String = "EqNB"
Private Const EQUATION_BOOKMARK_PREFIX As String = "_Eqn"
Private Const EQUATION_REF_BOOKMARK_PREFIX As String = "_EqnRef"

Public Sub InsertEquationLinePlain()
    InsertEquationLine "plain", "-"
End Sub

Public Sub InsertHashEquationLinePlain()
    InsertHashEquationLine GetNumberMode(ActiveDocument), GetEquationSeparator(ActiveDocument)
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
    Set equationRange = FindEquationRangeAt(doc, equationStart, equationEnd)
    equationRange.Select
    Exit Sub

Failed:
    MsgBox "Failed to insert inline equation: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub InsertEquationLineChapterHyphen()
    InsertEquationLine "chapter", "-"
End Sub

Public Sub InsertHashEquationLineChapterHyphen()
    InsertHashEquationLine "chapter", GetEquationSeparator(ActiveDocument)
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

Public Sub RibbonInsertHashPlain(ByVal control As IRibbonControl)
    InsertHashEquationLinePlain
End Sub

Public Sub RibbonInsertInlineEquation(ByVal control As IRibbonControl)
    InsertInlineEquation
End Sub

Public Sub RibbonInsertChapterHyphen(ByVal control As IRibbonControl)
    InsertEquationLineChapterHyphen
End Sub

Public Sub RibbonInsertHashChapterHyphen(ByVal control As IRibbonControl)
    InsertHashEquationLineChapterHyphen
End Sub

Public Sub RibbonInsertChapterDot(ByVal control As IRibbonControl)
    InsertEquationLineChapterDot
End Sub

Public Sub RibbonMarkChapterStart(ByVal control As IRibbonControl)
    MarkEquationChapterStart
End Sub

Public Sub RibbonInsertReference(ByVal control As IRibbonControl)
    InsertEquationReference
End Sub

Public Sub RibbonSetReferenceFormat(ByVal control As IRibbonControl)
    SetEquationReferenceFormat
End Sub

Public Sub RibbonSetNumberFormat(ByVal control As IRibbonControl)
    SetEquationNumberFormat
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
    MsgBox "Failed to insert equation line: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub InsertHashEquationLine(Optional ByVal mode As String = "plain", Optional ByVal separator As String = "-")
    On Error GoTo Failed

    Dim doc As Document
    Set doc = ActiveDocument

    Dim bookmarkName As String
    bookmarkName = CreateEquationBookmarkName()
    Dim referenceBookmarkName As String
    referenceBookmarkName = CreateEquationReferenceBookmarkName(bookmarkName)

    Selection.TypeParagraph

    Dim paragraphRange As Range
    Set paragraphRange = Selection.Paragraphs(1).Range
    paragraphRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    paragraphRange.ParagraphFormat.LeftIndent = 0
    paragraphRange.ParagraphFormat.RightIndent = 0
    paragraphRange.ParagraphFormat.FirstLineIndent = 0
    paragraphRange.ParagraphFormat.SpaceBefore = 6
    paragraphRange.ParagraphFormat.SpaceAfter = 6

    Dim hashRangeStart As Long
    Dim placeholderStart As Long
    Dim placeholderEnd As Long
    Dim captionStart As Long
    Dim captionEnd As Long

    hashRangeStart = Selection.Start
    placeholderStart = Selection.Start
    Selection.TypeText ChrW(&H25A1)
    placeholderEnd = Selection.End
    Dim delimiterLeft As String
    Dim delimiterRight As String
    delimiterLeft = GetDelimiterLeft(GetEquationDelimiter(doc))
    delimiterRight = GetDelimiterRight(GetEquationDelimiter(doc))

    Selection.TypeText "#" & delimiterLeft
    captionStart = Selection.Start

    Dim chapterFormat As String
    Dim equationFormat As String
    chapterFormat = GetChapterNumberFormat(doc)
    equationFormat = GetEquationNumberFormat(doc)

    If mode = "chapter" And chapterFormat <> "NONE" Then
        doc.Fields.Add Range:=Selection.Range, Type:=wdFieldSequence, Text:="Chapter \c " & GetFieldFormatSwitch(chapterFormat), PreserveFormatting:=False
        Selection.Collapse wdCollapseEnd
        If equationFormat <> "NONE" And separator <> "NONE" Then Selection.TypeText separator
    End If

    If equationFormat <> "NONE" Then
        doc.Fields.Add Range:=Selection.Range, Type:=wdFieldSequence, Text:=EQUATION_SEQ_NAME & " " & GetFieldFormatSwitch(equationFormat), PreserveFormatting:=False
        Selection.Collapse wdCollapseEnd
    End If
    captionEnd = Selection.End
    Selection.TypeText delimiterRight

    doc.Bookmarks.Add Name:=bookmarkName, Range:=doc.Range(captionStart, captionEnd)

    Dim hashRangeEnd As Long
    hashRangeEnd = Selection.End

    Dim equationRange As Range
    Set equationRange = doc.Range(hashRangeStart, hashRangeEnd)
    doc.OMaths.Add equationRange

    doc.Fields.Update

    TryFinalizeHashEquation doc, hashRangeStart, hashRangeEnd
    EnsureEquationBookmark doc, bookmarkName, captionStart, captionEnd
    doc.Fields.Update
    EnsureEquationReferenceBookmark doc, bookmarkName, referenceBookmarkName
    doc.Range(placeholderStart, placeholderEnd).Select
    SendKeys "{LEFT}{RIGHT}{RIGHT}", True
    DoEvents
    WaitSeconds 0.1
    Exit Sub

Failed:
    MsgBox "Failed to insert hash equation line: " & Err.Description, vbCritical, APP_TITLE
End Sub

Private Function TryFinalizeHashEquation(ByVal doc As Document, ByVal rangeStart As Long, ByVal rangeEnd As Long) As Boolean
    On Error GoTo GiveUp

    Dim equationRange As Range
    Set equationRange = FindEquationRangeAt(doc, rangeStart, rangeEnd)
    equationRange.Select
    Selection.Collapse wdCollapseEnd

    ' Word's equation-internal # numbering is finalized by an interactive Enter
    ' at the end of the equation. This mirrors the manual tutorial workflow.
    Application.Activate
    DoEvents
    SendKeys "{ENTER}", True
    DoEvents
    WaitSeconds 0.2
    DeleteEmptyParagraphAtSelection
    DoEvents
    TryFinalizeHashEquation = True
    Exit Function

GiveUp:
    TryFinalizeHashEquation = False
End Function

Private Sub DeleteEmptyParagraphAtSelection()
    On Error GoTo GiveUp

    Dim paragraphRange As Range
    Set paragraphRange = Selection.Paragraphs(1).Range

    Dim paragraphText As String
    paragraphText = paragraphRange.Text
    paragraphText = Replace(paragraphText, ChrW(13), "")
    paragraphText = Replace(paragraphText, ChrW(7), "")
    paragraphText = Trim$(paragraphText)

    If Len(paragraphText) = 0 Then
        paragraphRange.Delete
    End If

GiveUp:
End Sub

Private Sub EnsureEquationBookmark(ByVal doc As Document, ByVal bookmarkName As String, ByVal rangeStart As Long, ByVal rangeEnd As Long)
    On Error GoTo GiveUp

    If doc.Bookmarks.Exists(bookmarkName) Then
        NormalizeHashEquationBookmark doc, bookmarkName
        If Len(GetBookmarkDisplayText(doc.Bookmarks(bookmarkName))) > 0 Then Exit Sub
        doc.Bookmarks(bookmarkName).Delete
    End If

    If rangeStart < rangeEnd Then
        doc.Bookmarks.Add Name:=bookmarkName, Range:=doc.Range(rangeStart, rangeEnd)
    End If

GiveUp:
End Sub

Private Sub EnsureEquationReferenceBookmark(ByVal doc As Document, ByVal displayBookmarkName As String, ByVal referenceBookmarkName As String)
    On Error GoTo GiveUp

    If Not doc.Bookmarks.Exists(displayBookmarkName) Then Exit Sub

    Dim numberText As String
    numberText = GetBookmarkDisplayText(doc.Bookmarks(displayBookmarkName))
    If Len(numberText) = 0 Then Exit Sub

    If doc.Bookmarks.Exists(referenceBookmarkName) Then
        Dim oldReferenceRange As Range
        Set oldReferenceRange = doc.Bookmarks(referenceBookmarkName).Range
        doc.Bookmarks(referenceBookmarkName).Delete
        oldReferenceRange.Delete
    End If

    Dim insertionRange As Range
    Set insertionRange = doc.Bookmarks(displayBookmarkName).Range.Paragraphs(1).Range
    insertionRange.End = insertionRange.End - 1
    insertionRange.Collapse wdCollapseEnd
    insertionRange.InsertAfter numberText
    insertionRange.Font.Hidden = True

    doc.Bookmarks.Add Name:=referenceBookmarkName, Range:=doc.Range(insertionRange.Start, insertionRange.Start + Len(numberText))
    insertionRange.Collapse wdCollapseEnd
    insertionRange.Font.Hidden = False

GiveUp:
End Sub

Private Sub WaitSeconds(ByVal seconds As Double)
    Dim finishTime As Single
    finishTime = Timer + CSng(seconds)

    Do While Timer < finishTime
        DoEvents
    Loop
End Sub

Public Sub MarkEquationChapterStart()
    On Error GoTo Failed

    Dim doc As Document
    Set doc = ActiveDocument

    Selection.Collapse wdCollapseEnd
    doc.Fields.Add Range:=Selection.Range, Type:=wdFieldSequence, Text:="Chapter \h \* ARABIC", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    doc.Fields.Add Range:=Selection.Range, Type:=wdFieldSequence, Text:=EQUATION_SEQ_NAME & " \r 0 \h", PreserveFormatting:=False
    Selection.Collapse wdCollapseEnd
    doc.Fields.Update

    MsgBox "Chapter marker inserted. Chapter-style equation numbers will use this chapter value.", vbInformation, APP_TITLE
    Exit Sub

Failed:
    MsgBox "Failed to mark chapter start: " & Err.Description, vbCritical, APP_TITLE
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

    Dim referenceStart As Long
    referenceStart = Selection.Start

    Selection.TypeText Left$(formatText, markerPosition - 1)
    Dim referenceField As Field
    Set referenceField = ActiveDocument.Fields.Add(Range:=Selection.Range, Type:=wdFieldRef, Text:=refs(index)(3) & " \h \* CHARFORMAT", PreserveFormatting:=False)
    Selection.Collapse wdCollapseEnd
    Selection.TypeText Mid$(formatText, markerPosition + 3)
    ActiveDocument.Fields.Update
    ActiveDocument.Range(referenceStart, Selection.End).Font.Hidden = False
    referenceField.Update
    referenceField.Result.Font.Hidden = False
    Exit Sub

Failed:
    MsgBox "Failed to insert equation reference: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub SetEquationReferenceFormat()
    On Error GoTo Failed

    Dim formatText As String
    formatText = InputBox("Set the reference format for this document. Use {n} where the equation number should appear." & vbCrLf & _
        "Examples: ({n}), Equation ({n}), Eq.({n}), [{n}]", "Reference Format", GetEquationReferenceFormat(ActiveDocument))
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
    RefreshEquationReferenceBookmarks ActiveDocument
    ActiveDocument.Fields.Update
    MsgBox "Equation numbers and references were refreshed.", vbInformation, APP_TITLE
    Exit Sub

Failed:
    MsgBox "Failed to refresh fields: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub ShowEquationNumberingHelp()
    MsgBox "EqNB macros:" & vbCrLf & _
        "1. Insert numbered equations." & vbCrLf & _
        "2. Insert inline equations without numbering." & vbCrLf & _
        "3. Set one reference format for the document." & vbCrLf & _
        "4. Insert cross-references and refresh fields." & vbCrLf & _
        "5. Experimental # equations follow Word's equation-internal numbering method.", vbInformation, APP_TITLE
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

    RefreshEquationReferenceBookmarks ActiveDocument

    For Each bookmark In ActiveDocument.Bookmarks
        If IsEquationDisplayBookmark(bookmark.Name) Then
            Dim item(3) As String
            item(0) = bookmark.Name
            item(1) = GetBookmarkDisplayText(bookmark)
            item(2) = CStr(bookmark.Range.Start)
            item(3) = GetReferenceBookmarkNameForList(ActiveDocument, bookmark.Name)
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

Public Sub SetEquationNumberFormat()
    On Error GoTo Failed

    Dim modeChoice As String
    modeChoice = InputBox("Number mode:" & vbCrLf & _
        "1 = equation number only, e.g. (1)" & vbCrLf & _
        "2 = chapter-equation number, e.g. (1-1)", _
        "Equation Number Format", IIf(GetNumberMode(ActiveDocument) = "chapter", "2", "1"))
    If Len(Trim$(modeChoice)) = 0 Then Exit Sub

    Dim numberMode As String
    If Trim$(modeChoice) = "2" Then
        numberMode = "chapter"
    ElseIf Trim$(modeChoice) = "1" Then
        numberMode = "plain"
    Else
        MsgBox "Please enter 1 or 2.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    Dim chapterFormat As String
    chapterFormat = PromptNumberFormat("Chapter number format", GetChapterNumberFormat(ActiveDocument))
    If Len(chapterFormat) = 0 Then Exit Sub

    Dim equationFormat As String
    equationFormat = PromptNumberFormat("Equation number format", GetEquationNumberFormat(ActiveDocument))
    If Len(equationFormat) = 0 Then Exit Sub

    Dim delimiter As String
    delimiter = PromptDelimiter(GetEquationDelimiter(ActiveDocument))
    If Len(delimiter) = 0 Then Exit Sub

    Dim separator As String
    separator = PromptSeparator(GetEquationSeparator(ActiveDocument))
    If Len(separator) = 0 Then Exit Sub

    SetDocumentVariable ActiveDocument, NUMBER_MODE_VARIABLE, numberMode
    SetDocumentVariable ActiveDocument, CHAPTER_NUMBER_FORMAT_VARIABLE, chapterFormat
    SetDocumentVariable ActiveDocument, EQUATION_NUMBER_FORMAT_VARIABLE, equationFormat
    SetDocumentVariable ActiveDocument, EQUATION_DELIMITER_VARIABLE, delimiter
    SetDocumentVariable ActiveDocument, EQUATION_SEPARATOR_VARIABLE, separator

    MsgBox "Equation number format saved for this document." & vbCrLf & _
        "Preview: " & BuildNumberPreview(numberMode, chapterFormat, equationFormat, delimiter, separator), _
        vbInformation, APP_TITLE
    Exit Sub

Failed:
    MsgBox "Failed to set equation number format: " & Err.Description, vbCritical, APP_TITLE
End Sub

Private Function PromptNumberFormat(ByVal title As String, ByVal currentFormat As String) As String
    Dim answer As String
    answer = InputBox(title & ":" & vbCrLf & _
        "0 = None" & vbCrLf & _
        "1 = 123" & vbCrLf & _
        "2 = I II III" & vbCrLf & _
        "3 = i ii iii" & vbCrLf & _
        "4 = ABC" & vbCrLf & _
        "5 = abc", title, NumberFormatToChoice(currentFormat))

    If Len(Trim$(answer)) = 0 Then
        PromptNumberFormat = ""
    Else
        PromptNumberFormat = ChoiceToNumberFormat(Trim$(answer))
        If Len(PromptNumberFormat) = 0 Then MsgBox "Please enter 0, 1, 2, 3, 4, or 5.", vbExclamation, APP_TITLE
    End If
End Function

Private Function PromptDelimiter(ByVal currentDelimiter As String) As String
    Dim answer As String
    answer = InputBox("Delimiter:" & vbCrLf & _
        "0 = None" & vbCrLf & _
        "1 = ( )" & vbCrLf & _
        "2 = [ ]" & vbCrLf & _
        "3 = { }" & vbCrLf & _
        "4 = full-width ( )" & vbCrLf & _
        "5 = full-width [ ]", "Equation Number Format", DelimiterToChoice(currentDelimiter))

    If Len(Trim$(answer)) = 0 Then
        PromptDelimiter = ""
    Else
        PromptDelimiter = ChoiceToDelimiter(Trim$(answer))
        If Len(PromptDelimiter) = 0 And Trim$(answer) <> "0" Then MsgBox "Please enter 0, 1, 2, 3, 4, or 5.", vbExclamation, APP_TITLE
    End If
End Function

Private Function PromptSeparator(ByVal currentSeparator As String) As String
    Dim answer As String
    answer = InputBox("Separator between chapter and equation:" & vbCrLf & _
        "0 = None" & vbCrLf & _
        "1 = ." & vbCrLf & _
        "2 = -" & vbCrLf & _
        "3 = en dash" & vbCrLf & _
        "4 = em dash" & vbCrLf & _
        "5 = /", "Equation Number Format", SeparatorToChoice(currentSeparator))

    If Len(Trim$(answer)) = 0 Then
        PromptSeparator = ""
    Else
        PromptSeparator = ChoiceToSeparator(Trim$(answer))
        If Len(PromptSeparator) = 0 And Trim$(answer) <> "0" Then MsgBox "Please enter 0, 1, 2, 3, 4, or 5.", vbExclamation, APP_TITLE
    End If
End Function

Private Function GetNumberMode(ByVal doc As Document) As String
    GetNumberMode = GetDocumentVariable(doc, NUMBER_MODE_VARIABLE, DEFAULT_NUMBER_MODE)
    If GetNumberMode <> "plain" And GetNumberMode <> "chapter" Then GetNumberMode = DEFAULT_NUMBER_MODE
End Function

Private Function GetChapterNumberFormat(ByVal doc As Document) As String
    GetChapterNumberFormat = GetDocumentVariable(doc, CHAPTER_NUMBER_FORMAT_VARIABLE, DEFAULT_NUMBER_FORMAT)
    If Not IsValidNumberFormat(GetChapterNumberFormat) Then GetChapterNumberFormat = DEFAULT_NUMBER_FORMAT
End Function

Private Function GetEquationNumberFormat(ByVal doc As Document) As String
    GetEquationNumberFormat = GetDocumentVariable(doc, EQUATION_NUMBER_FORMAT_VARIABLE, DEFAULT_NUMBER_FORMAT)
    If Not IsValidNumberFormat(GetEquationNumberFormat) Then GetEquationNumberFormat = DEFAULT_NUMBER_FORMAT
End Function

Private Function GetEquationDelimiter(ByVal doc As Document) As String
    GetEquationDelimiter = GetDocumentVariable(doc, EQUATION_DELIMITER_VARIABLE, DEFAULT_DELIMITER)
    If Not IsValidDelimiter(GetEquationDelimiter) Then GetEquationDelimiter = DEFAULT_DELIMITER
End Function

Private Function GetEquationSeparator(ByVal doc As Document) As String
    GetEquationSeparator = ChoiceToSeparator(SeparatorToChoice(GetDocumentVariable(doc, EQUATION_SEPARATOR_VARIABLE, DEFAULT_SEPARATOR)))
End Function

Private Function GetDocumentVariable(ByVal doc As Document, ByVal variableName As String, ByVal defaultValue As String) As String
    On Error GoTo UseDefault

    GetDocumentVariable = doc.Variables(variableName).Value
    If Len(GetDocumentVariable) = 0 Then GoTo UseDefault
    Exit Function

UseDefault:
    GetDocumentVariable = defaultValue
End Function

Private Function ChoiceToNumberFormat(ByVal choice As String) As String
    Select Case choice
        Case "0"
            ChoiceToNumberFormat = "NONE"
        Case "1"
            ChoiceToNumberFormat = "ARABIC"
        Case "2"
            ChoiceToNumberFormat = "ROMAN"
        Case "3"
            ChoiceToNumberFormat = "roman"
        Case "4"
            ChoiceToNumberFormat = "ALPHABETIC"
        Case "5"
            ChoiceToNumberFormat = "alphabetic"
        Case Else
            ChoiceToNumberFormat = ""
    End Select
End Function

Private Function NumberFormatToChoice(ByVal numberFormat As String) As String
    Select Case numberFormat
        Case "NONE"
            NumberFormatToChoice = "0"
        Case "ARABIC"
            NumberFormatToChoice = "1"
        Case "ROMAN"
            NumberFormatToChoice = "2"
        Case "roman"
            NumberFormatToChoice = "3"
        Case "ALPHABETIC"
            NumberFormatToChoice = "4"
        Case "alphabetic"
            NumberFormatToChoice = "5"
        Case Else
            NumberFormatToChoice = "1"
    End Select
End Function

Private Function IsValidNumberFormat(ByVal numberFormat As String) As Boolean
    Select Case numberFormat
        Case "NONE", "ARABIC", "ROMAN", "roman", "ALPHABETIC", "alphabetic"
            IsValidNumberFormat = True
        Case Else
            IsValidNumberFormat = False
    End Select
End Function

Private Function GetFieldFormatSwitch(ByVal numberFormat As String) As String
    If numberFormat = "NONE" Then
        GetFieldFormatSwitch = ""
    Else
        GetFieldFormatSwitch = "\* " & numberFormat
    End If
End Function

Private Function ChoiceToDelimiter(ByVal choice As String) As String
    Select Case choice
        Case "0"
            ChoiceToDelimiter = "NONE"
        Case "1"
            ChoiceToDelimiter = "paren"
        Case "2"
            ChoiceToDelimiter = "bracket"
        Case "3"
            ChoiceToDelimiter = "brace"
        Case "4"
            ChoiceToDelimiter = "fwparen"
        Case "5"
            ChoiceToDelimiter = "fwbracket"
        Case Else
            ChoiceToDelimiter = ""
    End Select
End Function

Private Function DelimiterToChoice(ByVal delimiter As String) As String
    Select Case delimiter
        Case "NONE"
            DelimiterToChoice = "0"
        Case "paren"
            DelimiterToChoice = "1"
        Case "bracket"
            DelimiterToChoice = "2"
        Case "brace"
            DelimiterToChoice = "3"
        Case "fwparen"
            DelimiterToChoice = "4"
        Case "fwbracket"
            DelimiterToChoice = "5"
        Case Else
            DelimiterToChoice = "1"
    End Select
End Function

Private Function IsValidDelimiter(ByVal delimiter As String) As Boolean
    Select Case delimiter
        Case "NONE", "paren", "bracket", "brace", "fwparen", "fwbracket"
            IsValidDelimiter = True
        Case Else
            IsValidDelimiter = False
    End Select
End Function

Private Function GetDelimiterLeft(ByVal delimiter As String) As String
    Select Case delimiter
        Case "NONE"
            GetDelimiterLeft = ""
        Case "bracket"
            GetDelimiterLeft = "["
        Case "brace"
            GetDelimiterLeft = "{"
        Case "fwparen"
            GetDelimiterLeft = ChrW(&HFF08)
        Case "fwbracket"
            GetDelimiterLeft = ChrW(&H3010)
        Case Else
            GetDelimiterLeft = "("
    End Select
End Function

Private Function GetDelimiterRight(ByVal delimiter As String) As String
    Select Case delimiter
        Case "NONE"
            GetDelimiterRight = ""
        Case "bracket"
            GetDelimiterRight = "]"
        Case "brace"
            GetDelimiterRight = "}"
        Case "fwparen"
            GetDelimiterRight = ChrW(&HFF09)
        Case "fwbracket"
            GetDelimiterRight = ChrW(&H3011)
        Case Else
            GetDelimiterRight = ")"
    End Select
End Function

Private Function ChoiceToSeparator(ByVal choice As String) As String
    Select Case choice
        Case "0"
            ChoiceToSeparator = "NONE"
        Case "1"
            ChoiceToSeparator = "."
        Case "2"
            ChoiceToSeparator = "-"
        Case "3"
            ChoiceToSeparator = ChrW(8211)
        Case "4"
            ChoiceToSeparator = ChrW(8212)
        Case "5"
            ChoiceToSeparator = "/"
        Case Else
            ChoiceToSeparator = ""
    End Select
End Function

Private Function SeparatorToChoice(ByVal separator As String) As String
    Select Case separator
        Case "NONE"
            SeparatorToChoice = "0"
        Case "."
            SeparatorToChoice = "1"
        Case "-"
            SeparatorToChoice = "2"
        Case ChrW(8211)
            SeparatorToChoice = "3"
        Case ChrW(8212)
            SeparatorToChoice = "4"
        Case "/"
            SeparatorToChoice = "5"
        Case Else
            SeparatorToChoice = "2"
    End Select
End Function

Private Function BuildNumberPreview(ByVal numberMode As String, ByVal chapterFormat As String, ByVal equationFormat As String, ByVal delimiter As String, ByVal separator As String) As String
    Dim core As String
    core = ""

    If numberMode = "chapter" And chapterFormat <> "NONE" Then
        core = SampleNumber(chapterFormat)
    End If

    If equationFormat <> "NONE" Then
        If Len(core) > 0 And separator <> "NONE" Then core = core & separator
        core = core & SampleNumber(equationFormat)
    End If

    BuildNumberPreview = GetDelimiterLeft(delimiter) & core & GetDelimiterRight(delimiter)
End Function

Private Function SampleNumber(ByVal numberFormat As String) As String
    Select Case numberFormat
        Case "ROMAN"
            SampleNumber = "II"
        Case "roman"
            SampleNumber = "ii"
        Case "ALPHABETIC"
            SampleNumber = "B"
        Case "alphabetic"
            SampleNumber = "b"
        Case Else
            SampleNumber = "2"
    End Select
End Function

Private Sub NormalizeEquationBookmarks(ByVal doc As Document)
    On Error GoTo GiveUp

    Dim bookmarkNames As New Collection
    Dim bookmark As Bookmark
    Dim name As Variant

    For Each bookmark In doc.Bookmarks
        If IsEquationDisplayBookmark(bookmark.Name) Then
            bookmarkNames.Add bookmark.Name
        End If
    Next bookmark

    For Each name In bookmarkNames
        NormalizeHashEquationBookmark doc, CStr(name)
    Next name

GiveUp:
End Sub

Private Sub RefreshEquationReferenceBookmarks(ByVal doc As Document)
    On Error GoTo GiveUp

    Dim showHiddenBefore As Boolean
    showHiddenBefore = ActiveDocument.Bookmarks.ShowHidden
    ActiveDocument.Bookmarks.ShowHidden = True

    NormalizeEquationBookmarks doc

    Dim bookmarkNames As New Collection
    Dim bookmark As Bookmark
    Dim name As Variant

    For Each bookmark In doc.Bookmarks
        If IsEquationDisplayBookmark(bookmark.Name) Then
            bookmarkNames.Add bookmark.Name
        End If
    Next bookmark

    For Each name In bookmarkNames
        EnsureEquationReferenceBookmark doc, CStr(name), CreateEquationReferenceBookmarkName(CStr(name))
    Next name

    ActiveDocument.Bookmarks.ShowHidden = showHiddenBefore
    Exit Sub

GiveUp:
    ActiveDocument.Bookmarks.ShowHidden = showHiddenBefore
End Sub

Private Function GetReferenceBookmarkNameForList(ByVal doc As Document, ByVal displayBookmarkName As String) As String
    Dim referenceBookmarkName As String
    referenceBookmarkName = CreateEquationReferenceBookmarkName(displayBookmarkName)

    If doc.Bookmarks.Exists(referenceBookmarkName) Then
        GetReferenceBookmarkNameForList = referenceBookmarkName
    Else
        GetReferenceBookmarkNameForList = displayBookmarkName
    End If
End Function

Private Function IsEquationDisplayBookmark(ByVal bookmarkName As String) As Boolean
    If Left$(bookmarkName, Len(EQUATION_BOOKMARK_PREFIX)) <> EQUATION_BOOKMARK_PREFIX Then
        IsEquationDisplayBookmark = False
        Exit Function
    End If

    IsEquationDisplayBookmark = Left$(bookmarkName, Len(EQUATION_REF_BOOKMARK_PREFIX)) <> EQUATION_REF_BOOKMARK_PREFIX
End Function

Private Sub NormalizeHashEquationBookmark(ByVal doc As Document, ByVal bookmarkName As String)
    On Error GoTo GiveUp

    If Not doc.Bookmarks.Exists(bookmarkName) Then Exit Sub

    Dim bookmarkRange As Range
    Set bookmarkRange = doc.Bookmarks(bookmarkName).Range

    Dim text As String
    text = CleanEquationReferenceText(bookmarkRange.Text, False)

    Dim hashPosition As Long
    Dim openPosition As Long
    Dim closePosition As Long

    hashPosition = InStrRev(text, "#(")
    If hashPosition = 0 Then Exit Sub

    openPosition = hashPosition + 1
    closePosition = InStr(openPosition + 1, text, ")")
    If closePosition <= openPosition Then Exit Sub

    Dim newStart As Long
    Dim newEnd As Long
    newStart = bookmarkRange.Start + openPosition
    newEnd = bookmarkRange.Start + closePosition - 1

    If newStart >= newEnd Then Exit Sub

    doc.Bookmarks(bookmarkName).Delete
    doc.Bookmarks.Add Name:=bookmarkName, Range:=doc.Range(newStart, newEnd)

GiveUp:
End Sub

Private Function GetBookmarkDisplayText(ByVal bookmark As Bookmark) As String
    On Error GoTo UseRangeText

    Dim text As String
    text = TrimEquationReferenceText(bookmark.Range.Text)

    If bookmark.Range.Fields.Count > 1 Then
        Dim fieldText As String
        Dim fieldIndex As Long
        fieldText = text

        For fieldIndex = 1 To bookmark.Range.Fields.Count
            fieldText = Replace(fieldText, TrimEquationReferenceText(bookmark.Range.Fields(fieldIndex).Code.Text), "")
            fieldText = Replace(fieldText, ChrW(19), "")
            fieldText = Replace(fieldText, ChrW(20), "")
            fieldText = Replace(fieldText, ChrW(21), "")
        Next fieldIndex

        If Len(TrimEquationReferenceText(fieldText)) > 0 Then
            text = TrimEquationReferenceText(fieldText)
        End If
    End If

    GetBookmarkDisplayText = text
    Exit Function

UseRangeText:
    GetBookmarkDisplayText = TrimEquationReferenceText(bookmark.Range.Text)
End Function

Private Function TrimEquationReferenceText(ByVal text As String) As String
    TrimEquationReferenceText = CleanEquationReferenceText(text, True)
End Function

Private Function CleanEquationReferenceText(ByVal text As String, ByVal stripHashWrapper As Boolean) As String
    text = Replace(text, ChrW(13), "")
    text = Replace(text, ChrW(7), "")
    text = Replace(text, ChrW(19), "")
    text = Replace(text, ChrW(20), "")
    text = Replace(text, ChrW(21), "")
    If stripHashWrapper Then text = StripHashEquationWrapper(text)
    If stripHashWrapper Then text = StripLeadingHash(text)
    CleanEquationReferenceText = Trim$(text)
End Function

Private Function StripHashEquationWrapper(ByVal text As String) As String
    Dim hashPosition As Long
    Dim openPosition As Long
    Dim closePosition As Long
    Dim closeDelimiter As String

    hashPosition = InStrRev(text, "#")
    If hashPosition = 0 Then
        StripHashEquationWrapper = text
        Exit Function
    End If

    openPosition = hashPosition + 1
    If openPosition > Len(text) Then
        StripHashEquationWrapper = text
        Exit Function
    End If

    closeDelimiter = MatchingDelimiter(Mid$(text, openPosition, 1))
    If Len(closeDelimiter) = 0 Then
        StripHashEquationWrapper = text
        Exit Function
    End If

    closePosition = InStr(openPosition + 1, text, closeDelimiter)

    If closePosition > openPosition Then
        StripHashEquationWrapper = Mid$(text, openPosition + 1, closePosition - openPosition - 1)
    Else
        StripHashEquationWrapper = Mid$(text, openPosition + 1)
    End If
End Function

Private Function MatchingDelimiter(ByVal openDelimiter As String) As String
    Select Case openDelimiter
        Case "("
            MatchingDelimiter = ")"
        Case "["
            MatchingDelimiter = "]"
        Case "{"
            MatchingDelimiter = "}"
        Case ChrW(&HFF08)
            MatchingDelimiter = ChrW(&HFF09)
        Case ChrW(&H3010)
            MatchingDelimiter = ChrW(&H3011)
        Case Else
            MatchingDelimiter = ""
    End Select
End Function

Private Function StripLeadingHash(ByVal text As String) As String
    text = Trim$(text)
    Do While Left$(text, 1) = "#"
        text = Trim$(Mid$(text, 2))
    Loop
    StripLeadingHash = text
End Function

Private Function CreateEquationBookmarkName() As String
    Randomize
    CreateEquationBookmarkName = EQUATION_BOOKMARK_PREFIX & Format$(Now, "yymmddhhnnss") & CStr(Int(Rnd() * 1000))
End Function

Private Function CreateEquationReferenceBookmarkName(ByVal displayBookmarkName As String) As String
    If Left$(displayBookmarkName, Len(EQUATION_BOOKMARK_PREFIX)) = EQUATION_BOOKMARK_PREFIX Then
        CreateEquationReferenceBookmarkName = EQUATION_REF_BOOKMARK_PREFIX & Mid$(displayBookmarkName, Len(EQUATION_BOOKMARK_PREFIX) + 1)
    Else
        CreateEquationReferenceBookmarkName = EQUATION_REF_BOOKMARK_PREFIX & displayBookmarkName
    End If
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
