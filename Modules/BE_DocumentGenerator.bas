Option Explicit

' 자동 보고서 생성을 담당하는 모듈.

Private Const wdCollapseEnd As Long = 0
Private Const wdAutoFitContent As Long = 1

Public Sub GenerateStructuredDataReport(ByVal SheetName As String, _
                                       Optional ByVal TableName As String = "", _
                                       Optional ByVal OutputPath As String = "", _
                                       Optional ByVal TemplatePath As String = "", _
                                       Optional ByVal ShowWord As Boolean = False)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dataBody As Variant
    Dim headers As Variant
    Dim rowCount As Long, colCount As Long
    Dim outputFile As String
    Dim wordApp As Object, wordDoc As Object

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets(SheetName)
    Set lo = ResolveListObject(ws, TableName)

    If Not lo Is Nothing Then
        colCount = lo.ListColumns.Count
        headers = HeaderArrayFromListObject(lo)
        If lo.DataBodyRange Is Nothing Then
            rowCount = 0
        Else
            dataBody = lo.DataBodyRange.Value
            rowCount = lo.DataBodyRange.Rows.Count
        End If
    Else
        Dim rng As Range
        Set rng = GetSourceRange(ws)
        If rng Is Nothing Then Err.Raise vbObjectError + 513, "GenerateStructuredDataReport", "데이터 범위를 찾을 수 없습니다."

        headers = HeaderArrayFromFirstRow(rng.Value)
        colCount = GetColumnCountFromHeaders(headers)

        If rng.Rows.Count > 1 Then
            dataBody = DataBodyFromValues(rng.Value)
            rowCount = rng.Rows.Count - 1
        Else
            rowCount = 0
        End If
    End If

    outputFile = ResolveOutputPath(OutputPath, ws.Name)
    EnsureFolderExists GetFolderFromPath(outputFile)

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = ShowWord
    wordApp.DisplayAlerts = False

    If Len(TemplatePath) > 0 Then
        TemplatePath = NormalizePath(TemplatePath)
        If Dir$(TemplatePath) = "" Then Err.Raise vbObjectError + 514, "GenerateStructuredDataReport", _
                                                    "템플릿 파일을 찾을 수 없습니다: " & TemplatePath
        Set wordDoc = wordApp.Documents.Add(TemplatePath)
    Else
        Set wordDoc = wordApp.Documents.Add
    End If

    BuildDocument wordDoc, headers, dataBody, rowCount, colCount, ws.Name

    wordDoc.SaveAs2 outputFile

    If ShowWord Then
        wordApp.Visible = True
    Else
        wordDoc.Close SaveChanges:=False
        wordApp.Quit SaveChanges:=False
    End If

    MsgBox "보고서가 생성되었습니다." & vbCrLf & outputFile, vbInformation

Cleanup:
    On Error Resume Next
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Sub

ErrHandler:
    Dim errMsg As String
    errMsg = "GenerateStructuredDataReport 실패: " & Err.Description
    MsgBox errMsg, vbCritical
    On Error Resume Next
    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit SaveChanges:=False
    End If
    Resume Cleanup
End Sub

Private Function ResolveListObject(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error GoTo ErrHandler

    If Len(tableName) > 0 Then
        Set ResolveListObject = ws.ListObjects(tableName)
    ElseIf ws.ListObjects.Count = 1 Then
        Set ResolveListObject = ws.ListObjects(1)
    ElseIf ws.ListObjects.Count > 1 Then
        Err.Raise vbObjectError + 515, "ResolveListObject", "여러 개의 표가 있습니다. TableName을 지정해 주세요."
    End If

    Exit Function
ErrHandler:
    Err.Raise Err.Number, "ResolveListObject", Err.Description
End Function

Private Function GetSourceRange(ByVal ws As Worksheet) As Range
    Dim rng As Range

    On Error Resume Next
    Set rng = ws.UsedRange
    On Error GoTo 0

    If rng Is Nothing Then Exit Function
    If rng.Cells.Count = 1 Then
        If Len(Trim$(CStr(rng.Value))) = 0 Then
            Set rng = Nothing
        End If
    End If

    Set GetSourceRange = rng
End Function

Private Function HeaderArrayFromListObject(ByVal lo As ListObject) As Variant
    Dim headers As Variant
    Dim arr() As String
    Dim c As Long

    headers = lo.HeaderRowRange.Value
    ReDim arr(1 To lo.ListColumns.Count)
    For c = 1 To lo.ListColumns.Count
        arr(c) = Trim$(CStr(headers(1, c)))
    Next c
    HeaderArrayFromListObject = arr
End Function

Private Function HeaderArrayFromFirstRow(ByVal values As Variant) As Variant
    Dim arr() As String
    Dim firstRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim c As Long

    If Not IsArray(values) Then
        ReDim arr(1 To 1)
        arr(1) = Trim$(CStr(values))
        HeaderArrayFromFirstRow = arr
        Exit Function
    End If

    firstRow = LBound(values, 1)
    firstCol = LBound(values, 2)
    lastCol = UBound(values, 2)

    ReDim arr(1 To lastCol - firstCol + 1)
    For c = firstCol To lastCol
        arr(c - firstCol + 1) = Trim$(CStr(values(firstRow, c)))
    Next c

    HeaderArrayFromFirstRow = arr
End Function

Private Function GetColumnCountFromHeaders(ByVal headers As Variant) As Long
    If IsEmpty(headers) Then
        GetColumnCountFromHeaders = 0
    Else
        GetColumnCountFromHeaders = UBound(headers) - LBound(headers) + 1
    End If
End Function

Private Function DataBodyFromValues(ByVal values As Variant) As Variant
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim r As Long, c As Long
    Dim arr() As Variant

    If Not IsArray(values) Then Exit Function

    firstRow = LBound(values, 1)
    lastRow = UBound(values, 1)
    If lastRow <= firstRow Then Exit Function

    firstCol = LBound(values, 2)
    lastCol = UBound(values, 2)

    ReDim arr(1 To lastRow - firstRow, 1 To lastCol - firstCol + 1)
    For r = firstRow + 1 To lastRow
        For c = firstCol To lastCol
            arr(r - firstRow, c - firstCol + 1) = values(r, c)
        Next c
    Next r

    DataBodyFromValues = arr
End Function

Private Function ResolveOutputPath(ByVal requestedPath As String, ByVal sheetName As String) As String
    Dim basePath As String
    Dim normalized As String

    If Len(Trim$(requestedPath)) > 0 Then
        normalized = NormalizePath(requestedPath)
        Dim dotPos As Long
        Dim slashPos As Long
        dotPos = InStrRev(normalized, ".")
        slashPos = InStrRev(normalized, "\")
        If Right$(normalized, 1) = "\" Then
            normalized = normalized & sheetName & "_Report_" & Format$(Now, "yyyymmdd_hhnnss") & ".docx"
        ElseIf dotPos = 0 Or (slashPos > 0 And dotPos < slashPos) Then
            normalized = normalized & ".docx"
        End If
        ResolveOutputPath = normalized
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then basePath = CurDir$
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"

    ResolveOutputPath = basePath & sheetName & "_Report_" & Format$(Now, "yyyymmdd_hhnnss") & ".docx"
End Function

Private Function NormalizePath(ByVal rawPath As String) As String
    Dim trimmed As String
    Dim basePath As String

    trimmed = Trim$(rawPath)
    If Len(trimmed) = 0 Then
        NormalizePath = trimmed
        Exit Function
    End If

    trimmed = Replace(trimmed, "/", "\")
    If InStr(trimmed, ":") = 0 And Left$(trimmed, 2) <> "\\" Then
        basePath = ThisWorkbook.Path
        If Len(basePath) = 0 Then basePath = CurDir$
        If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
        NormalizePath = basePath & trimmed
    Else
        NormalizePath = trimmed
    End If
End Function

Private Function GetFolderFromPath(ByVal filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    If pos = 0 Then
        GetFolderFromPath = ""
    Else
        GetFolderFromPath = Left$(filePath, pos - 1)
    End If
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String

    If Len(folderPath) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 And parentPath <> folderPath Then EnsureFolderExists parentPath

    fso.CreateFolder folderPath
End Sub

Private Sub BuildDocument(ByVal wordDoc As Object, ByVal headers As Variant, ByVal dataBody As Variant, _
                          ByVal rowCount As Long, ByVal colCount As Long, ByVal sheetLabel As String)
    Dim docRange As Object

    Set docRange = wordDoc.Content
    docRange.Text = sheetLabel & " 데이터 보고서" & vbCrLf
    docRange.Style = wordDoc.Styles("Title")
    docRange.InsertParagraphAfter
    docRange.Collapse wdCollapseEnd

    docRange.Text = "생성 시각: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
    docRange.InsertParagraphAfter
    docRange.Collapse wdCollapseEnd

    If rowCount = 0 Or colCount = 0 Then
        docRange.Text = "사용 가능한 데이터가 없습니다." & vbCrLf
        docRange.InsertParagraphAfter
        Exit Sub
    End If

    docRange.Text = "총 행 수: " & rowCount & vbCrLf & _
                    "총 열 수: " & colCount & vbCrLf & vbCrLf
    docRange.InsertParagraphAfter
    docRange.Collapse wdCollapseEnd

    InsertSectionHeading wordDoc, "요약 정보"
    InsertSummaryTable wordDoc, headers, dataBody, rowCount, colCount
    InsertSectionHeading wordDoc, "열별 요약"
    InsertBulletList wordDoc, BuildColumnSummaryLines(headers, dataBody, rowCount, colCount)
    InsertSectionHeading wordDoc, "행별 요약"
    InsertBulletList wordDoc, BuildRowSummaryLines(headers, dataBody, rowCount, colCount)
End Sub

Private Sub InsertSectionHeading(ByVal wordDoc As Object, ByVal headingText As String)
    Dim rng As Object
    Set rng = wordDoc.Content
    rng.Collapse wdCollapseEnd
    rng.Text = headingText & vbCrLf
    rng.Style = wordDoc.Styles("Heading 1")
    rng.InsertParagraphAfter
End Sub

Private Sub InsertSummaryTable(ByVal wordDoc As Object, ByVal headers As Variant, ByVal dataBody As Variant, _
                               ByVal rowCount As Long, ByVal colCount As Long)
    Dim rng As Object
    Dim tbl As Object
    Dim r As Long, c As Long

    Set rng = wordDoc.Content
    rng.Collapse wdCollapseEnd

    Set tbl = wordDoc.Tables.Add(rng, rowCount + 1, colCount)
    tbl.AutoFitBehavior wdAutoFitContent
    tbl.Rows(1).HeadingFormat = True
    tbl.Rows(1).Range.Bold = True
    tbl.Rows(1).Shading.BackgroundPatternColor = RGB(236, 236, 236)

    For c = LBound(headers) To UBound(headers)
        tbl.Cell(1, c - LBound(headers) + 1).Range.Text = headers(c)
    Next c

    For r = 1 To rowCount
        For c = 1 To colCount
            tbl.Cell(r + 1, c).Range.Text = ToSafeText(dataBody, r, c)
        Next c
    Next r

    Set rng = tbl.Range
    rng.Collapse wdCollapseEnd
    rng.InsertParagraphAfter
End Sub

Private Sub InsertBulletList(ByVal wordDoc As Object, ByVal lines As Variant)
    Dim textBlock As String
    Dim i As Long
    Dim insertRange As Object
    Dim listRange As Object
    Dim endPosition As Long

    If IsEmpty(lines) Then Exit Sub

    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            textBlock = textBlock & Trim$(lines(i)) & vbCrLf
        End If
    Next i

    If Len(textBlock) = 0 Then Exit Sub
    If Right$(textBlock, 2) = vbCrLf Then textBlock = Left$(textBlock, Len(textBlock) - 2)

    Set insertRange = wordDoc.Content
    insertRange.Collapse wdCollapseEnd
    insertRange.Text = textBlock
    endPosition = insertRange.End

    Set listRange = wordDoc.Range(insertRange.Start, endPosition)
    listRange.ListFormat.ApplyBulletDefault

    Dim tailRange As Object
    Set tailRange = wordDoc.Range(endPosition, endPosition)
    tailRange.InsertParagraphAfter
    tailRange.ListFormat.RemoveNumbers
End Sub

Private Function BuildColumnSummaryLines(ByVal headers As Variant, ByVal dataBody As Variant, _
                                         ByVal rowCount As Long, ByVal colCount As Long) As Variant
    Dim lines() As String
    Dim c As Long

    If rowCount = 0 Or colCount = 0 Then Exit Function

    ReDim lines(1 To colCount)
    For c = 1 To colCount
        lines(c) = BuildSingleColumnSummary(headers(c), dataBody, rowCount, c)
    Next c

    BuildColumnSummaryLines = lines
End Function

Private Function BuildSingleColumnSummary(ByVal headerText As String, ByVal dataBody As Variant, _
                                          ByVal rowCount As Long, ByVal colIndex As Long) As String
    Dim numericCount As Long
    Dim textCount As Long
    Dim blankCount As Long
    Dim sumValues As Double
    Dim minValue As Double
    Dim maxValue As Double
    Dim firstNumeric As Boolean
    Dim dict As Object
    Dim r As Long
    Dim v As Variant
    Dim parts As Collection
    Dim sampleText As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set parts = New Collection
    firstNumeric = True

    For r = 1 To rowCount
        v = dataBody(r, colIndex)
        If IsError(v) Or IsNull(v) Then
            blankCount = blankCount + 1
        ElseIf IsNumeric(v) Then
            numericCount = numericCount + 1
            sumValues = sumValues + CDbl(v)
            If firstNumeric Then
                minValue = CDbl(v)
                maxValue = CDbl(v)
                firstNumeric = False
            Else
                If CDbl(v) < minValue Then minValue = CDbl(v)
                If CDbl(v) > maxValue Then maxValue = CDbl(v)
            End If
        Else
            Dim textValue As String
            textValue = Trim$(CStr(v))
            If Len(textValue) = 0 Then
                blankCount = blankCount + 1
            Else
                textCount = textCount + 1
                If Not dict.Exists(textValue) Then
                    dict.Add textValue, 1
                Else
                    dict(textValue) = dict(textValue) + 1
                End If
            End If
        End If
    Next r

    If numericCount > 0 Then
        parts.Add "숫자 " & numericCount & "건 (평균 " & FormatNumber(sumValues / numericCount, 2, vbTrue, vbFalse, vbFalse) & _
                   ", 최소 " & FormatNumber(minValue, 2, vbTrue, vbFalse, vbFalse) & _
                   ", 최대 " & FormatNumber(maxValue, 2, vbTrue, vbFalse, vbFalse) & ")"
    End If

    If dict.Count > 0 Then
        sampleText = BuildSampleValues(dict, 3)
        If Len(sampleText) > 0 Then
            parts.Add "텍스트 " & textCount & "건, 고유값 " & dict.Count & "건 (주요 값: " & sampleText & ")"
        Else
            parts.Add "텍스트 " & textCount & "건, 고유값 " & dict.Count & "건"
        End If
    End If

    If blankCount > 0 Then parts.Add "공백 " & blankCount & "건"
    If parts.Count = 0 Then parts.Add "데이터 없음"

    BuildSingleColumnSummary = headerText & " - " & JoinCollection(parts, ", ")
End Function

Private Function BuildSampleValues(ByVal dict As Object, ByVal topN As Long) As String
    Dim keys As Variant
    Dim counts As Variant
    Dim i As Long, j As Long
    Dim limit As Long
    Dim tempValue As Variant
    Dim result() As String

    If dict Is Nothing Then Exit Function
    If dict.Count = 0 Then Exit Function

    keys = dict.Keys
    counts = dict.Items

    For i = 0 To dict.Count - 2
        For j = i + 1 To dict.Count - 1
            If counts(j) > counts(i) Then
                tempValue = counts(i)
                counts(i) = counts(j)
                counts(j) = tempValue

                tempValue = keys(i)
                keys(i) = keys(j)
                keys(j) = tempValue
            End If
        Next j
    Next i

    limit = MinLong(topN, dict.Count)
    ReDim result(0 To limit - 1)
    For i = 0 To limit - 1
        result(i) = CStr(keys(i)) & " (" & counts(i) & ")"
    Next i

    BuildSampleValues = Join(result, ", ")
End Function

Private Function BuildRowSummaryLines(ByVal headers As Variant, ByVal dataBody As Variant, _
                                      ByVal rowCount As Long, ByVal colCount As Long) As Variant
    Dim lines() As String
    Dim r As Long

    If rowCount = 0 Or colCount = 0 Then Exit Function

    ReDim lines(1 To rowCount)
    For r = 1 To rowCount
        lines(r) = BuildSingleRowSummary(headers, dataBody, r, colCount)
    Next r

    BuildRowSummaryLines = lines
End Function

Private Function BuildSingleRowSummary(ByVal headers As Variant, ByVal dataBody As Variant, _
                                       ByVal rowIndex As Long, ByVal colCount As Long) As String
    Dim parts As Collection
    Dim c As Long
    Dim label As String

    Set parts = New Collection

    label = headers(LBound(headers)) & "=" & ToSafeText(dataBody, rowIndex, 1)

    If colCount > 1 Then
        For c = 2 To colCount
            Dim valueText As String
            valueText = ToSafeText(dataBody, rowIndex, c)
            If Len(valueText) > 0 Then
                parts.Add headers(c) & "=" & valueText
            End If
        Next c
    End If

    If parts.Count > 0 Then
        BuildSingleRowSummary = label & " | " & JoinCollection(parts, ", ")
    Else
        BuildSingleRowSummary = label
    End If
End Function

Private Function JoinCollection(ByVal col As Collection, ByVal delimiter As String) As String
    Dim arr() As String
    Dim i As Long

    If col Is Nothing Then Exit Function
    If col.Count = 0 Then Exit Function

    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        arr(i - 1) = CStr(col(i))
    Next i

    JoinCollection = Join(arr, delimiter)
End Function

Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then
        MinLong = a
    Else
        MinLong = b
    End If
End Function

Private Function ToSafeText(ByVal dataBody As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    Dim v As Variant

    If IsEmpty(dataBody) Then Exit Function
    If Not IsArray(dataBody) Then Exit Function

    v = dataBody(rowIndex, colIndex)

    If IsError(v) Then
        ToSafeText = "#오류"
    ElseIf IsNull(v) Then
        ToSafeText = ""
    ElseIf IsDate(v) Then
        ToSafeText = Format$(CDate(v), "yyyy-mm-dd hh:nn:ss")
    Else
        ToSafeText = Trim$(CStr(v))
    End If
End Function
