Option Explicit

Private isParsed As Boolean
Private twb As Workbook, tws As Worksheet
Private rwb As Workbook, rws As Worksheet
Public Sub Read_Documents(Optional Handle As Boolean = False)
    Dim DPCount As Long, PLCount As Long, MDCount As Long, i As Long, c As Long, Cycle As Long
    Dim vDate(1 To 2) As String, vLine(1 To 2) As String
    Dim Dir_Main As String: Dir_Main = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "")
    Dim Dir_DP As String: Dir_DP = Dir_Main & "DailyPlan"
    Dim Dir_PLi As String: Dir_PLi = Dir_Main & "PartList"
    Dim Clt_DP As New Collection: Set Clt_DP = FindFilesWithTextInName(Dir_DP, "DailyPlan", ".xlsx")
    Dim Clt_PLi As New Collection: Set Clt_PLi = FindFilesWithTextInName(Dir_PLi, "PartList", ".xlsx")
    Dim LV_MD As ListView: Set LV_MD = AutoReportHandler.ListView_MD_Own: LV_MD.ListItems.Clear
    
    FillListView_Intersection Clt_DP, Clt_PLi, LV_MD, 2025, "날짜", "라인", "DailyPlan", "PartList"

    DPCount = Clt_DP.Count: PLCount = Clt_PLi.Count: MDCount = LV_MD.ListItems.Count
    If Handle Then MsgBox "DailyPlan : " & DPCount & "장 연결됨" & vbLf & _
                                "PartList : " & PLCount & "장 연결됨" & vbLf & _
                                "Multi Documents : " & MDCount & "장 연결됨" & vbLf & _
                                Cycle
End Sub

Private Sub SetUp_Targets(ByRef Target_WorkBook As Workbook, ByRef Target_WorkSheet As Worksheet, _
                            ByRef Reference_WorkBook As Workbook, ByRef Reference_WorkSheet As Worksheet)
    Set twb = Target_WorkBook: Set tws = Target_WorkSheet: Set rwb = Reference_WorkBook: Set rws = Reference_WorkSheet
End Sub
                            
Private Sub Parse_wbwsPointer()
    Dim Linked(1 To 4) As Boolean
    Linked(1) = Not twb Is Nothing: Linked(2) = Not tws Is Nothing: Linked(3) = Not rwb Is Nothing: Linked(4) = Not rws Is Nothing
    If Linked(1) And Linked(2) And Linked(3) And Linked(4) Then Exit Sub
    Set twb = Nothing: Set tws = Nothing: Set rwb = Nothing: Set rws = Nothing
    
    isParsed = True ' Parsing Boolean
End Sub

Public Sub MixMatching(ByVal Target_item As String)
    
End Sub