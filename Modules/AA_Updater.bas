' 폴더구조로 선별 후 출력
Sub ExportAllVbaComponents()
    Dim vbComp As Object
    Dim FSO As Object
    Dim basePath As String
    Dim folderModules As String, folderClasses As String, folderForms As String
    Dim FileName As String

    ' 기본 경로 설정
    basePath = ThisWorkbook.Path & "\ExcelExportedCodes\"
    folderModules = basePath & "Modules\"
    folderClasses = basePath & "Classes\"
    folderForms = basePath & "Forms\"

    ' 폴더 생성
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(basePath) Then FSO.CreateFolder basePath
    If Not FSO.FolderExists(folderModules) Then FSO.CreateFolder folderModules
    If Not FSO.FolderExists(folderClasses) Then FSO.CreateFolder folderClasses
    If Not FSO.FolderExists(folderForms) Then FSO.CreateFolder folderForms

    ' 구성 요소 반복하며 내보내기
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: FileName = folderModules & vbComp.Name & ".bas"   ' 표준 모듈
            Case 2: FileName = folderClasses & vbComp.Name & ".cls"   ' 클래스 모듈
            Case 3: FileName = folderForms & vbComp.Name & ".frm"     ' 사용자 폼
            Case Else: FileName = vbNullString
        End Select

        If FileName <> vbNullString Then
            vbComp.Export FileName
        End If
    Next vbComp

    MsgBox "구성 요소가 폴더 구조로 내보내졌습니다." & vbLf & basePath, vbInformation
End Sub
' .Txt .Md 출력
Sub ExportAllModulesDirectlyToTextAndMarkdown()
    Dim vbComp As Object
    Dim FSO As Object
    Dim exportPath As String
    Dim ext As String, FileName As String
    Dim codeLine As Variant
    Dim codeLines() As String
    Dim txtStream As Object, mdStream As Object
    Dim baseName As String, timeStamp As String
    Dim TxtFile As String, mdFile As String
    Dim totalLines As Long

    ' 파일명 구성
    baseName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    timeStamp = Format(Now, "yymmddhhmm")
    exportPath = ThisWorkbook.Path & "\ExcelExportedCodes\"
    TxtFile = exportPath & baseName & "_SourceCode_" & timeStamp & ".txt"
    mdFile = exportPath & baseName & "_SourceCode_" & timeStamp & ".md"

    ' 폴더 생성
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(exportPath) Then FSO.CreateFolder exportPath

    ' 스트림 생성 (UTF-8)
    Set txtStream = CreateObject("ADODB.Stream")
    With txtStream
        .Charset = "utf-8"
        .Type = 2
        .open
    End With

    Set mdStream = CreateObject("ADODB.Stream")
    With mdStream
        .Charset = "utf-8"
        .Type = 2
        .open
    End With

    ' 구성 요소 반복
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: ext = ".bas"
            Case 2: ext = ".cls"
            Case 3: ext = ".frm"
            Case Else: ext = ""
        End Select

        If ext <> "" Then
            FileName = vbComp.Name & ext
            totalLines = vbComp.CodeModule.CountOfLines

            ' 코드 읽기
            If totalLines > 0 Then
                codeLines = Split(vbComp.CodeModule.Lines(1, totalLines), vbLf)
            Else
                codeLines = Split("", vbLf)
            End If

            ' TXT 파일 작성
            txtStream.WriteText String(60, "'") & vbLf
            txtStream.WriteText FileName & " Start" & vbLf
            txtStream.WriteText String(60, "'") & vbLf

            ' MD 파일 작성
            mdStream.WriteText "### " & FileName & vbLf
            mdStream.WriteText "````vba" & vbLf

            For Each codeLine In codeLines
                txtStream.WriteText codeLine & vbLf
                mdStream.WriteText codeLine & vbLf
            Next codeLine

            txtStream.WriteText String(60, "'") & vbLf
            txtStream.WriteText FileName & " End" & vbLf
            txtStream.WriteText String(60, "'") & vbLf & vbLf

            mdStream.WriteText "````" & vbLf & vbLf
        End If
    Next vbComp

    ' 저장 및 닫기
    txtStream.SaveToFile TxtFile, 2
    txtStream.Close
    mdStream.SaveToFile mdFile, 2
    mdStream.Close

    MsgBox "모든 코드가 병합되어 저장되었습니다!" & vbLf & _
           TxtFile & vbLf & mdFile, vbInformation
End Sub

Sub ForceUpdateMacro()
    Dim latestVersion As String
    Dim localVersion As String
    Dim versionUrl As String
    Dim fileUrl As String
    Dim savePath As String
    Dim ws As Worksheet
    Dim VersionCell As Range
    
    ' Setting 확인
    Set ws = ThisWorkbook.Worksheets("Setting")
    If ("Dev" = ws.Cells.Find(What:="Develop", lookAt:=xlWhole, MatchCase:=True).Offset(0, 1).value) Then
        MsgBox "개발 모드이므로 업데이트 진행 제한", vbInformation, "개발여부 확인"
        Exit Sub
    End If
    Set VersionCell = ws.Cells.Find(What:="Version", lookAt:=xlWhole, MatchCase:=True)
    'Debug.Print VersionCell.Address
    
    ' GitLab Raw URL 설정
    versionUrl = "http://mod.lge.com/hub/seongsu1.lee/excelmacroupdater/-/raw/main/Version.txt"
    fileUrl = "http://mod.lge.com/hub/seongsu1.lee/excelmacroupdater/-/raw/main/AutoReport.xlsb"
    
    ' 현재 사용 중인 버전 (Setting Worksheet의 Version 행을 찾아 값 열의 값을 참조함)
    localVersion = VersionCell.Offset(0, 1).value
    
    ' 최신 버전 확인
    latestVersion = GetWebText(versionUrl)
    
    ' 버전 비교 및 업데이트 수행
    If Trim(localVersion) < Trim(latestVersion) Then
        MsgBox "새 버전(" & latestVersion & ")이 감지되었습니다. 업데이트를 진행합니다.", vbInformation
        
        ' 다운로드 경로 설정
        savePath = Environ("TEMP") & "\NewMacro.xlsb"
        
        ' 최신 매크로 파일 다운로드
        If DownloadFile(fileUrl, savePath) Then
            ' 기존 파일 닫기 및 새 파일 실행
            ThisWorkbook.Close False
            Workbooks.open savePath
        Else
            MsgBox "업데이트 다운로드에 실패했습니다.", vbExclamation
        End If
    Else
        MsgBox "현재 최신 버전을 사용 중입니다.", vbInformation
    End If
End Sub

Function GetWebText(url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' 요청 전송
    http.open "GET", url, False
    http.Send
    
    ' 응답 확인
    If http.Status = 200 Then
        GetWebText = http.responseText
    Else
        GetWebText = "Error"
    End If
End Function

Function DownloadFile(url As String, savePath As String) As Boolean
    Dim http As Object
    Dim stream As Object
    
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' 파일 다운로드 요청
    http.open "GET", url, False
    http.Send
    
    ' 다운로드 확인
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.open
        stream.Write http.responseBody
        stream.SaveToFile savePath, 2
        stream.Close
        
        ' 다운로드 성공
        DownloadFile = True
    Else
        ' 다운로드 실패
        DownloadFile = False
    End If
End Function