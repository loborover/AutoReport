Option Explicit

#If Win64 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As LongPtr, ByVal wMsg As Long, _
        ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Public UI As AutoReportHandler

'---------------------------
' 1) 파싱 결과를 담는 UDT
'---------------------------
Public Type MDToken
    DocType As DocumentTypes   ' dc_DailyPlan / dc_PartList
    Month As Integer
    Day As Integer
    LineAddr As String         ' 예: "C11"
    fullPath As String         ' 원본 경로
    FileName As String         ' 파일명만
    DateValue As Date          ' BaseYear 적용된 실제 Date
    WeekdayVb As VbDayOfWeek   ' vbMonday 등
    WeekdayK As String         ' "월","화","수" ...
End Type

Public Enum DocumentTypes
    dc_BOM = -11
    dc_DailyPlan = -12
    dc_PartList = -13
End Enum
Public Enum MorS
    MainG = -100
    SubG = -200
End Enum
Public Enum RorC
    Row = -144
    Column = -211
End Enum
Public Enum arLabelShape
    Arrow = 58
    Arrow_Done = 36
    Box = 1
    Box_Dash = 2
    Box_Rounded = 5
    Box_Card = 75
    Box_Hexagon = 10
    Box_Octagon = 6
    Box_Pentagon = 51
    Box_Plque = 28
    Cross = 11
    Round = 69
    Vrtcl_Connecter = 74
    Vrtcl_ArrowCallout = 56
    Vrtcl_Document = 67
    Vrtcl_Wave = 103
    Vrtcl_Rounded = 86
End Enum
Public Enum ObjDirectionVertical
    dvBothSide = -48
    dvUP = -88
    dvMid = 48
    dvDown = 88
End Enum
Public Enum ObjDirectionSide
    dsLeft = -44
    dsRight = 44
End Enum
Public Enum ObjDirection4Way
    d4UP = -88
    d4DOWN = 88
    d4LEFT = -44
    d4RIGHT = 44
End Enum

Public Function SaveFilesWithCustomDirectory(directoryPath As String, _
                ByRef wb As Workbook, _
                ByRef PDFpagesetup As PrintSetting, _
                Optional ByRef vTitle As String = "UndefinedFile", _
                Optional SaveToXlsx As Boolean = False, _
                Optional SaveToPDF As Boolean = True, _
                Optional OriginalKiller As Boolean = True) As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Dim ExcelPath As String, savePath As String, ToDeleteDir As String
    ExcelPath = ThisWorkbook.Path: ToDeleteDir = wb.FullName
'주소가 없으면 생성
    If Dir(ExcelPath & "\" & directoryPath, vbDirectory) = "" Then MkDir ExcelPath & "\" & directoryPath
'파일 저장용 주소 생성
    savePath = ExcelPath & "\" & directoryPath & "\" & vTitle
'이미 저장된 파일이 있다면 삭제
    If Dir(savePath & ".xlsx") <> "" Then Kill savePath & ".xlsx"
    If Dir(savePath & ".pdf") <> "" Then Kill savePath & ".pdf"
'PDF 셋업 후 PDF출력
    AutoPageSetup ws, PDFpagesetup
    If SaveToPDF Then ws.PrintOut ActivePrinter:="Microsoft Print to PDF", PrintToFile:=True, prtofilename:=savePath & ".pdf"
'엑셀로 저장할지 결정
    If SaveToXlsx Then wb.Close SaveChanges:=True, FileName:=savePath Else wb.Close SaveChanges:=False
    If OriginalKiller Then Kill ToDeleteDir
    SaveFilesWithCustomDirectory = savePath
    On Error GoTo 0
End Function

Function FindFilesWithTextInName(directoryPath As String, searchText As String, _
                                        Optional FileExtForSort As String) As Collection
    Dim FileName As String, filePath As String, FEFS As Long
    Dim resultPaths As New Collection
    
    FileName = Dir(directoryPath & "\*.*") ' 지정된 디렉토리에서 파일 목록 얻기
    ' 파일 목록을 확인하면서 조건에 맞는 파일 찾기
    Do While FileName <> ""
        ' 파일 이름에 특정 텍스트가 포함되어 있는지 확인
        FEFS = IIf(FileExtForSort = "", 1, InStr(1, FileName, FileExtForSort, vbBinaryCompare))
        If InStr(1, FileName, searchText, vbTextCompare) > 0 And FEFS > 0 Then
            ' 조건에 맞는 파일의 경로를 생성
            filePath = directoryPath & "\" & FileName
            ' 조건에 맞는 파일의 경로를 리스트에 추가
            resultPaths.Add filePath
        End If
        FileName = Dir ' 다음 파일 검색
    Loop
    
    ' 조건에 맞는 파일이 하나 이상인 경우 리스트 반환
    If resultPaths.Count > 0 Then
        Set FindFilesWithTextInName = resultPaths
    Else
        ' 조건에 맞는 파일을 찾지 못한 경우 빈 Collection 반환
        Set FindFilesWithTextInName = New Collection
    End If
End Function

Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim element As Variant
    On Error Resume Next
    IsInArray = (UBound(Filter(arr, valToBeFound)) > -1)
    On Error GoTo 0
End Function

Public Function IsInCollection(valToBeFound As Variant, col As Collection) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If valToBeFound = col(i) Then
            IsInCollection = True
            Exit Function
        Else
            IsInCollection = False
        End If
    Next i
End Function

Function ColumnLetter(ColumnNumber As Long) As String
    Dim d As Long
    Dim m As Long
    Dim Name As String
    
    d = ColumnNumber
    Do
        m = (d - 1) Mod 26
        Name = Chr(65 + m) & Name
        d = (d - m) \ 26
    Loop While d > 0
    
    ColumnLetter = Name
End Function

Public Function GetRangeBoundary(rng As Range, _
                                         Optional ByRef First_Row As Long = -1, Optional ByRef Last_Row As Long = -1, _
                                        Optional ByRef First_Column As Long = -1, Optional ByRef Last_Column As Long = -1, _
                                        Optional isLeftToRight As Boolean = True) As Long
    Dim FOS As Boolean ' True = Function, False = Sub
    If First_Row = -1 Or _
        First_Column = -1 Or _
        Last_Row = -1 Or _
        Last_Column = -1 Then FOS = True
    
    First_Row = rng.Row
    Last_Row = rng.Rows(rng.Rows.Count).Row
    
    If isLeftToRight Then
        First_Column = rng.Column
        Last_Column = rng.Columns(rng.Columns.Count).Column
    Else
        First_Column = rng.Columns(rng.Columns.Count).Column
        Last_Column = rng.Column
    End If
    
    If Not FOS Then Exit Function
    
    GetRangeBoundary = First_Row
    
End Function

' CountCountinuousNonEmptyCells / 비어있지 않은 셀의 개수를 반환하는 함수 / CountNonEmptyCells
Public Function fCCNEC(ByVal TargetRange As Range) As Long
    Dim cell As Range
    Dim Count As Long
    Dim foundValue As Boolean

    Count = 0
    foundValue = False
    
    For Each cell In TargetRange
        If Not IsEmpty(cell.value) Then
            If Not foundValue Then
                foundValue = True ' 최초의 값 있는 셀을 찾음
            End If
            Count = Count + 1 ' 연속된 값 카운트
        ElseIf foundValue Then
            Exit For ' 첫 값 이후 공백을 만나면 종료
        End If
    Next cell
    
    fCCNEC = Count
End Function

' 셀 기준으로  줄 긋는 서브루틴
Public Sub CellLiner(ByRef Target As Range, _
                                Optional vEdge As XlBordersIndex = xlEdgeTop, _
                                Optional vLineStyle As XlLineStyle = xlContinuous, _
                                Optional vWeight As XlBorderWeight = xlThin)
    Dim ws As Worksheet: Set ws = Target.Worksheet
    Dim PrcssR As Range, vRorC As String
    
    If vEdge = xlEdgeTop Or xlEdgeBottom Then
        vRorC = CStr(Target.Row)
    ElseIf vEdge = xlEdgeLeft Or xlEdgeRight Then
        vRorC = CStr(Target.Column)
    Else: Exit Sub
    End If
    Set PrcssR = ws.Range(vRorC & ":" & vRorC)
    With PrcssR.Borders(vEdge)
        .LineStyle = vLineStyle
        .Weight = vWeight
        .Color = RGB(0, 0, 0)
    End With
End Sub

Public Function ForLining(ByRef Target As Range, Optional Division As RorC = Row) As Range
    Dim ws As Worksheet: Set ws = Target.Parent
    
    Select Case Division
    Case Row
        Set ForLining = ws.Range(Target.Row & ":" & Target.Row)
    Case Column
        Set ForLining = ws.Range(Target.Column & ":" & Target.Column)
    End Select
    
End Function

' Utillity CFAW_PDF
Public Function CheckFileAlreadyWritten_PDF(ByRef Document_Name As String, dt As DocumentTypes) As String
    Dim Document_Path As String, DTs As String
    
    Select Case dt
        Case -11 ' BOM
            DTs = "BOM"
            Document_Name = Replace(Document_Name, ".", "_") & ".pdf"
        Case -12 ' DailyPlan
            DTs = "DailyPlan"
            Document_Name = Document_Name & ".pdf"
        Case -13 ' PartList
            DTs = "PartList"
            Document_Name = Document_Name & ".pdf"
    End Select
    
    Document_Path = ThisWorkbook.Path & "\" & DTs
    If Not Dir(Document_Path & "\" & Document_Name, vbDirectory) <> "" Then
        CheckFileAlreadyWritten_PDF = "Ready"
        Exit Function
    Else
        CheckFileAlreadyWritten_PDF = "Written"
        Exit Function
    End If
End Function
Public Sub SelfMerge(ByRef MergeTarget As Range)
    Dim r As Long, c As Long
    Dim cell As Range
    Dim ValueList As String
    'Dim ws As Worksheet: Set ws = MergeTarget.Parent
    
    If MergeTarget Is Nothing Then Exit Sub
    For r = 1 To MergeTarget.Rows.Count
        For c = 1 To MergeTarget.Columns.Count
            Set cell = MergeTarget.Cells(r, c)
            If Trim(cell.value) <> "" Then ValueList = ValueList & cell.value & vbLf
        Next c
    Next r
    
    ' 병합 및 텍스트 삽입
    With MergeTarget
        .Merge
        .value = ValueList
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Function ExtractBracketValue(ByVal Txt As String, Optional ByRef Searching As Long = 1) As String
    Dim sPos As Long, ePos As Long
    sPos = InStr(Searching, Txt, "["): ePos = InStr(Searching + 1, Txt, "]")
    
    If sPos > 0 And ePos > sPos Then
        ExtractBracketValue = Mid(Txt, sPos + 1, ePos - sPos - 1)
    Else
        ExtractBracketValue = ""
    End If
    Searching = ePos
End Function

Public Sub DeleteDuplicateRowsInColumn(ByVal targetCol As Long, ByRef StartRow As Long, ByRef EndRow As Long, _
        Optional ByRef tgtWs As Worksheet)

    Dim colValues As New Collection   ' 중복 체크용 컬렉션
    Dim i As Long, DeleteRowCount As Long
    Dim cellVal As String

    If tgtWs Is Nothing Then Set tgtWs = ActiveSheet ' 범용성 확보

    ' 아래에서 위로 순회하면서 중복 검사 및 삭제
    For i = EndRow To StartRow Step -1
        ' 지정된 컬럼의 값을 가져와 공백 제거
        cellVal = Trim$(tgtWs.Cells(i, targetCol).value)

        ' 빈 문자열이 아닐 때만 검사
        If Len(cellVal) > 0 Then
            On Error Resume Next
            ' 키로 cellVal을 지정하여 컬렉션에 추가 시도
            colValues.Add Item:=cellVal, key:=cellVal

            ' 오류 번호 457: 이미 동일한 Key가 존재함을 의미
            If Err.Number = 457 Then
                ' 중복으로 판단된 행을 삭제
                tgtWs.Rows(i).Delete
                DeleteRowCount = DeleteRowCount + 1
            End If

            ' 오류 상태 초기화
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    
    EndRow = EndRow - DeleteRowCount
End Sub

'---------------------------
' 2) 정규식 헬퍼(Late Binding)
'---------------------------
Private Function RxFirst(ByVal pattern As String, ByVal text As String) As String
    Dim rx As Object, m As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.pattern = pattern
    rx.Global = False
    rx.IgnoreCase = True
    If rx.test(text) Then
        Set m = rx.Execute(text)(0)
        RxFirst = m.SubMatches(0) ' 반드시 () 캡처 1개짜리 패턴 전제
    Else
        RxFirst = vbNullString
    End If
End Function

'---------------------------
' 3) 한국어 요일 반환
'---------------------------
Private Function WeekdayKorean(d As Date) As String
    Select Case Weekday(d, vbSunday)
        Case vbSunday:    WeekdayKorean = "일"
        Case vbMonday:    WeekdayKorean = "월"
        Case vbTuesday:   WeekdayKorean = "화"
        Case vbWednesday: WeekdayKorean = "수"
        Case vbThursday:  WeekdayKorean = "목"
        Case vbFriday:    WeekdayKorean = "금"
        Case vbSaturday:  WeekdayKorean = "토"
    End Select
End Function

'---------------------------
' 4) 파일명 파서
'   예) "DailyPlan 5월-28일_C11.xlsx"
'---------------------------
Private Function ParseMDToken(ByVal fullPath As String, Optional ByVal BaseYear As Long = 0) As MDToken
    Dim t As MDToken, nm As String
    Dim ms As String, ds As String, ln As String, dt As Date, y As Long
   
    nm = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    nm = Replace$(nm, ".xlsx", "", , , vbTextCompare)
    t.fullPath = fullPath
    t.FileName = nm
   
    ' 문서타입
    If InStr(1, nm, "DailyPlan", vbTextCompare) > 0 Then
        t.DocType = dc_DailyPlan
    ElseIf InStr(1, nm, "PartList", vbTextCompare) > 0 Then
        t.DocType = dc_PartList
    Else
        t.DocType = 0 ' 알 수 없음
    End If
   
    ' 월/일   (예: "5월-28일" / "09월-05일")
    ms = RxFirst("([0-9]{1,2})(?=월)", nm)
    ds = RxFirst("([0-9]{1,2})(?=일)", nm)
   
    If Len(ms) > 0 Then t.Month = CInt(ms)
    If Len(ds) > 0 Then t.Day = CInt(ds)
   
    ' 라인   (예: "_C11" , "C11")
    ln = RxFirst("C([0-9]{1,3})", nm)
    If Len(ln) > 0 Then t.LineAddr = "C" & ln
   
    ' 연도
    If BaseYear = 0 Then
        y = Year(Date) ' 기본 현재 연도
    Else
        y = BaseYear
    End If
   
    If t.Month >= 1 And t.Day >= 1 Then
        On Error Resume Next
        dt = DateSerial(y, t.Month, t.Day)
        On Error GoTo 0
        If dt > 0 Then
            t.DateValue = dt
            t.WeekdayVb = Weekday(dt, vbSunday)
            t.WeekdayK = WeekdayKorean(dt)
        End If
    End If
   
    ParseMDToken = t
End Function

'---------------------------------------------
' 5) ListView 선별 추가기 (요일/라인 필터)
'    wantDocType  : 0 이면 타입 무시
'    wantLine     : "" 이면 라인 무시 (예: "C11")
'    wantWeekday  : 0 이면 요일 무시 (vbMonday 등)
'---------------------------------------------
Public Sub FillListView_ByFilter(ByRef files As Collection, ByRef lv As ListView, _
        Optional ByVal wantDocType As DocumentTypes = 0, _
        Optional ByVal wantLine As String = "", _
        Optional ByVal wantWeekday As VbDayOfWeek = 0, _
        Optional ByVal BaseYear As Long = 0)
   
    Dim i As Long
    Dim t As MDToken
    Dim it As listItem
   
    With lv
        .ListItems.Clear
        ' 컬럼 헤더 구성 예시 (필요 시 한 번만 구성)
        If .ColumnHeaders.Count = 0 Then
            .ColumnHeaders.Add , , "날짜"
            .ColumnHeaders.Add , , "요일"
            .ColumnHeaders.Add , , "라인"
            .ColumnHeaders.Add , , "문서"
            .ColumnHeaders.Add , , "경로"
        End If
    End With
   
    For i = 1 To files.Count
        t = ParseMDToken(CStr(files(i)), BaseYear)
        If wantDocType <> 0 Then If t.DocType <> wantDocType Then GoTo CONTINUE_NEXT ' 타입 필터
        If Len(wantLine) > 0 Then If StrComp(t.LineAddr, wantLine, vbTextCompare) <> 0 Then GoTo CONTINUE_NEXT ' 라인 필터
        If wantWeekday <> 0 Then If t.WeekdayVb <> wantWeekday Then GoTo CONTINUE_NEXT ' 요일 필터
       
        ' ListView 입력
        If t.DateValue > 0 Then
            Set it = lv.ListItems.Add(, , Format$(t.DateValue, "m월-d일"))
        Else
            Set it = lv.ListItems.Add(, , "미상")
        End If
       
        it.SubItems(1) = t.WeekdayK
        it.SubItems(2) = IIf(Len(t.LineAddr) > 0, t.LineAddr, "-")
        it.SubItems(3) = IIf(t.DocType = dc_DailyPlan, "DailyPlan", IIf(t.DocType = dc_PartList, "PartList", "-"))
        it.SubItems(4) = t.fullPath
        it.Checked = True
       
CONTINUE_NEXT:
    Next i
End Sub

'---------------------------------------------
' 6) 사용 중인 GetFoundSentences 교체판
'    - 패턴 문자열 대신 용도 구분: "date" 또는 "line"
'    - 기존 코드 호환 목적: "*월-*일" -> "date", "*-Line" -> "line"
'---------------------------------------------
Public Function GetFoundSentences(ByVal Search As String, ByVal Target As String) As String
    Dim nm As String, ms As String, ds As String, ln As String
    nm = Mid$(Target, InStrRev(Target, "\") + 1)
    nm = Replace$(nm, ".xlsx", "", , , vbTextCompare)
   
    If InStr(1, Search, "월", vbTextCompare) > 0 Then
        ms = RxFirst("([0-9]{1,2})(?=월)", nm)
        ds = RxFirst("([0-9]{1,2})(?=일)", nm)
        If Len(ms) > 0 And Len(ds) > 0 Then
            GetFoundSentences = CStr(CLng(ms)) & "월-" & CStr(CLng(ds)) & "일"
        Else
            GetFoundSentences = ""
        End If
        Exit Function
    End If
   
    If InStr(1, Search, "Line", vbTextCompare) > 0 Or InStr(1, Search, "C", vbTextCompare) > 0 Then
        ln = RxFirst("C([0-9]{1,3})", nm)
        If Len(ln) > 0 Then GetFoundSentences = "C" & ln Else GetFoundSentences = ""
        Exit Function
    End If
   
    ' 기타: 기본은 공백
    GetFoundSentences = ""
End Function
'--- 날짜/라인 키 빌드: 파일명 예) "DailyPlan 5월-28일_C11.xlsx"
Private Function BuildKeyFromPath(ByVal fullPath As String, Optional ByVal BaseYear As Long = 0) As String
    Dim nm As String, m As String, d As String, ln As String
    Dim y As Long, dt As Date
   
    nm = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    nm = Replace$(nm, ".xlsx", "", , , vbTextCompare)
   
    m = RxFirst("([0-9]{1,2})(?=월)", nm)
    d = RxFirst("([0-9]{1,2})(?=일)", nm)
    ln = RxFirst("C([0-9]{1,3})", nm)
   
    If Len(m) = 0 Or Len(d) = 0 Or Len(ln) = 0 Then
        BuildKeyFromPath = vbNullString
        Exit Function
    End If
   
    If BaseYear = 0 Then y = Year(Date) Else y = BaseYear
    On Error Resume Next
    dt = DateSerial(y, CLng(m), CLng(d))
    On Error GoTo 0
    If dt = 0 Then
        BuildKeyFromPath = vbNullString
        Exit Function
    End If
   
    ' 키 정규화: yyyy-mm-dd|C##
    BuildKeyFromPath = Format$(dt, "yyyy-mm-dd") & "|" & "C" & CStr(CLng(ln))
End Function

'--- 교집합을 outLV에 채우기 (입력: 파일 경로 컬렉션 2개)
Public Sub FillListView_Intersection(ByRef filesA As Collection, ByRef filesB As Collection, ByRef outLV As ListView, _
                                            Optional ByVal BaseYear As Long = 0, _
                                            Optional ByVal A_Discription As String, Optional ByVal B_Discription As String, Optional ByVal C_Discription As String, Optional ByVal D_Discription As String)
    Dim i As Long
    Dim keyMap As New Collection         ' Key 전용 Map (Collection을 Map처럼 사용)
    Dim itemA As String, itemB As String, key As String
    Dim it As listItem
    If A_Discription = "" Then A_Discription = "A경로": If B_Discription = "" Then B_Discription = "B경로"
    If C_Discription = "" Then C_Discription = "C경로": If D_Discription = "" Then D_Discription = "D경로"
    ' 컬럼 구성(최초 1회)
    With outLV
        .ListItems.Clear
        If .ColumnHeaders.Count = 0 Then
            .ColumnHeaders.Add , , A_Discription, LenA(A_Discription)
            .ColumnHeaders.Add , , B_Discription, LenA(B_Discription)
            .ColumnHeaders.Add , , C_Discription, LenA(C_Discription)
            .ColumnHeaders.Add , , D_Discription, LenA(D_Discription)
        End If
    End With
   
    ' 1) A집합 Key 적재 (Key 충돌은 무시)
    For i = 1 To filesA.Count
        itemA = CStr(filesA(i))
        key = BuildKeyFromPath(itemA, BaseYear)
        If Len(key) > 0 Then
            On Error Resume Next
                keyMap.Add itemA, key     ' Item=원본경로, Key=정규화키
                ' 이미 존재하면 Err=457 -> 최초 한 개만 보관(존재성 체크가 목적)
                Err.Clear
            On Error GoTo 0
        End If
    Next i
   
    ' 2) B를 순회하며 교집합만 출력
    For i = 1 To filesB.Count
        itemB = CStr(filesB(i))
        key = BuildKeyFromPath(itemB, BaseYear)
        If Len(key) = 0 Then GoTo CONT_NEXT
       
        ' 존재성 검사: col.Item(key) → 에러 없으면 존재
        Dim aPath As String, dtText As String, lnText As String
        On Error Resume Next
            aPath = CStr(keyMap.Item(key))   ' 없으면 에러
        If Err.Number = 0 Then
            ' 키에서 표시용 날짜/라인 분리
            dtText = Split(key, "|")(0)      ' yyyy-mm-dd
            lnText = Split(key, "|")(1)      ' C##
            With outLV
                Set it = .ListItems.Add(, , Format$(CDate(dtText), "m월-d일"))
                it.SubItems(1) = lnText
                it.SubItems(2) = aPath
                it.SubItems(3) = itemB
                it.Checked = True
            End With
        End If
        Err.Clear
        On Error GoTo 0
CONT_NEXT:
    Next i
    'LvAutoFit outLV
End Sub

' 문자열의 예상 폭을 pt로 근사 계산 (가볍고 빠른 추정치)
Public Function LenA(ByVal Expression As String, _
                     Optional ByVal Achr As Single = 14.9, _
                     Optional ByVal LatinScale As Single = 2 / 5) As Single
    Dim w As Single, i As Long, code As Long, n As Long: n = Len(Expression)
    If n = 0 Then LenA = 0: Exit Function
    For i = 1 To n
        code = AscW(Mid$(Expression, i, 1)) ' Mid$ 사용: Variant 방지 + 약간 더 빠름
        If code >= &HAC00 And code <= &HD7A3 Then w = w + Achr Else w = w + Achr * LatinScale ' 가(AC00=44032) ~ 힣(D7A3=55203)
    Next i
    LenA = w  ' Single 그대로 반환 (소수점 유지)
End Function

Public Sub LvAutoFit(ByRef lvw As MSComctlLib.ListView, Optional ByVal UseHeader As Boolean = True)
    Const LVM_FIRST& = &H1000
    Const LVM_SETCOLUMNWIDTH& = (LVM_FIRST + 30)
    Const LVSCW_AUTOSIZE& = -1
    Const LVSCW_AUTOSIZE_USEHEADER& = -2
    Dim i As Long, mode As Long
    mode = IIf(UseHeader, LVSCW_AUTOSIZE_USEHEADER, LVSCW_AUTOSIZE)
    For i = 0 To lvw.ColumnHeaders.Count - 1
        Call SendMessage(lvw.hWnd, LVM_SETCOLUMNWIDTH, i, mode)
    Next
End Sub

Public Sub Diagnose_MSCOMCTL()
    Debug.Print String(60, "-")
    Debug.Print "[Excel/Office Bitness]"
#If Win64 Then
    Debug.Print "Office: 64-bit"
#Else
    Debug.Print "Office: 32-bit"
#End If

    Debug.Print String(60, "-")
    Debug.Print "[Common Controls OCX Files]"
    Debug.Print "SysWOW64\\MSCOMCTL.OCX : "; IIf(FileExists("C:\Windows\SysWOW64\MSCOMCTL.OCX"), "Yes", "No")
    Debug.Print "System32\\MSCOMCTL.OCX : "; IIf(FileExists("C:\Windows\System32\MSCOMCTL.OCX"), "Yes", "No")
    Debug.Print "Office VFS\\MSCOMCTL   : "; IIf(FileExists("C:\Program Files\Microsoft Office\root\VFS\System\MSCOMCTL.OCX"), "Yes", "No")

    Debug.Print String(60, "-")
    Debug.Print "[References Status]"
    On Error Resume Next
    Dim r As Reference
    For Each r In ThisWorkbook.VBProject.References
        Debug.Print IIf(r.IsBroken, "MISSING: ", "OK      : "); r.Description
    Next r
    On Error GoTo 0

    Debug.Print String(60, "-")
    Debug.Print "※ MISSING이면 Tools>References에서 Browse로 MSCOMCTL.OCX 재지정 후 체크."
End Sub

Private Function FileExists(ByVal f As String) As Boolean
    FileExists = (Len(Dir$(f, vbNormal)) > 0)
End Function
