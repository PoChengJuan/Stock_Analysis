Imports eCA_HostObject
''' <summary>
''' 20181117
''' V1.0.0
''' Mark
''' HostHandling的主要object
''' </summary>
Public Class clsHandlingObject
  '----DBTool----
  Friend Shared gDBTool As Dictionary(Of String, eCA_DBTool.clsDBTool)

  '----LogTool----
  Friend Shared gLogTool As eCALogTool._ILogTool

  Private ShareName As String = "Handling"
  Private ShareKey As String = "Handling"

  '班別資訊
  Public gdicClass As New Dictionary(Of String, clsClass)
  '系統狀態資訊(一些暫存的系統參數)
  Public gdicSystemStatus As New Dictionary(Of String, clsSystemStatus)

  '客製項目
  Public gdicLine_Area As New Dictionary(Of String, clsLine_Area) '生產線區域
  Public gdicLine As New Dictionary(Of String, clsLine_Status) '生產線機器狀態
  Public gdicLine_Hist As New Dictionary(Of String, clsLine_Status_Hist) '生產線機器狀態記錄
  Public gdicLineInfo As New Dictionary(Of String, clsLineInfo) '生產線機器維謢資訊
  Public gdicLineProduction_Info As New Dictionary(Of String, clsLineProduction_Info) '生產線機器生產資訊
  Public gdicProduce_Info As New Dictionary(Of String, clsProduce_Info) '生產資訊
  Public gdicReport_Set As New Dictionary(Of String, clsDATA_REPORT_SET) '設定給BigData設式的異常和警告的設定值
  Public gdicMaintenance As New Dictionary(Of String, clsMAINTENANCE) '-機台保養參數設定
  Public gdicMaintenance_DTL As New Dictionary(Of String, clsMAINTENANCE_DTL) '-保養設定明細
  Public gdicMaintenance_Status As New Dictionary(Of String, clsMAINTENANCE_STATUS) '-保養狀態
  Public gdicClassAssignation As New Dictionary(Of String, clsCLASS_ASSIGNATION) '-每個班別人員分配
  Public gdicClassAttendance As New Dictionary(Of String, clsCLASS_ATTENDANCE) '-保養狀態
  '14-1.Alarm
  Public gdicAlarm As New Dictionary(Of String, clsALARM)
  Public gdicOwner As New Dictionary(Of String, clsOwner)
  Public gdicCT_PO_DTL As New Dictionary(Of String, clsHOST_CT_TMP_PO_DTL)

  '各個需要進行Lock的項目
  Public objLineProduction_InfoLock As New Object
  Public objProduction_InfoLock As New Object
  Public objCT_PO_DTLLock As New Object

  'Business_Rule
  Public gdicBusiness_Rule As New Dictionary(Of String, clsBusiness_Rule)

  '事件交握
  Public gdicCommand_Report As New Dictionary(Of String, clsCommandReport)
  '實作WAIT_UUID
  Public gdicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)

  '物件建立時執行的事件
  Public Sub New(ByRef New_DBTool As Dictionary(Of String, eCA_DBTool.clsDBTool), ByRef New_LogTool As eCALogTool._ILogTool, ByRef RetMsg As String)
    MyBase.New()
    Try
      '傳入主程式的DBTool和LogTool
      gDBTool = New_DBTool
      gLogTool = New_LogTool
      'LoadDB並建構所有物件的關係
      Init_HandlingObject(RetMsg)

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '物件結束時觸發的事件，用來清除物件的內容
  Protected Overrides Sub Finalize()
    Class_Terminate_Renamed()
    MyBase.Finalize()
  End Sub
  Private Sub Class_Terminate_Renamed()
    '目的:結束物件
    gdicLine_Area = Nothing
    gdicLine = Nothing
    gdicLineInfo = Nothing
    gdicLineProduction_Info = Nothing
    gdicProduce_Info = Nothing
  End Sub

  '=================Public Function=======================
  '資料加入Dictionary

  ''把PO加入gcolPO
  'Public Function O_Add_PO(ByRef obj As clsPO) As Boolean
  '  Try
  '    Dim key As String = obj.gid()
  '    If Not gdicPO.ContainsKey(key) Then
  '      gdicPO.TryAdd(key, obj)
  '    Else
  '      SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''把PO_DTL加入gcolPO_DTL
  'Public Function O_Add_PO_DTL(ByRef obj As clsPO_DTL) As Boolean
  '  Try
  '    Dim key As String = obj.gid()
  '    If Not gdicPO_DTL.ContainsKey(key) Then
  '      gdicPO_DTL.TryAdd(key, obj)
  '    Else
  '      SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''把PO_Line加入gcolPO_Line
  'Public Function O_Add_PO_Line(ByRef obj As clsPO_LINE) As Boolean
  '  Try
  '    Dim key As String = obj.gid()
  '    If Not gdicPO_Line.ContainsKey(key) Then
  '      gdicPO_Line.TryAdd(key, obj)
  '    Else
  '      SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function

  '從gcolBusiness_Rule取得指定的objBusiness_Rule
  Public Function O_Get_Business_Rule(ByVal Rule_No As enuBusinessRuleNO, Optional ByRef RetObj As clsBusiness_Rule = Nothing) As Boolean
    Try
      Dim key As String = clsBusiness_Rule.Get_Combination_Key(Rule_No)
      Dim obj As clsBusiness_Rule
      If gdicBusiness_Rule.ContainsKey(key) Then
        obj = gdicBusiness_Rule.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '把Business_Rule加入gdicBusiness_Rule
  Public Function O_Add_Business_Rule(ByRef obj As clsBusiness_Rule) As Boolean
    Try
      Dim key As String = obj.gid
      If Not gdicBusiness_Rule.ContainsKey(key) Then
        gdicBusiness_Rule.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從gcolBusiness_Rule刪除objBusiness_Rule
  Public Function O_Remove_Business_Rule(ByRef obj As clsBusiness_Rule) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicBusiness_Rule.ContainsKey(key) Then
        gdicBusiness_Rule.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '把CLine_Area加入gcolCLine_Area
  Public Function O_Add_Line_Area(ByRef obj As clsLine_Area) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicLine_Area.ContainsKey(key) Then
        gdicLine_Area.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '把CLine加入gcolCLine
  Public Function O_Add_Line(ByRef obj As clsLine_Status) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicLine.ContainsKey(key) Then
        gdicLine.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '把CLineInfo加入gcolCLineInfo
  Public Function O_Add_LineInfo(ByRef obj As clsLineInfo) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicLineInfo.ContainsKey(key) Then
        gdicLineInfo.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '把CLineProduction_Info加入gcolCLineProduction_Info
  Public Function O_Add_LineProduction_Info(ByRef obj As clsLineProduction_Info) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicLineProduction_Info.ContainsKey(key) Then
        gdicLineProduction_Info.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '把CProduce_Info加入gcolCProduce_Info
  Public Function O_Add_Produce_Info(ByRef obj As clsProduce_Info) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicProduce_Info.ContainsKey(key) Then
        gdicProduce_Info.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Add_Class(ByRef obj As clsClass) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicClass.ContainsKey(key) Then
        gdicClass.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Add_ClassAssignation(ByRef obj As clsCLASS_ASSIGNATION) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicClassAssignation.ContainsKey(key) Then
        gdicClassAssignation.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Add_ClassAttendance(ByRef obj As clsCLASS_ATTENDANCE) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicClassAttendance.ContainsKey(key) Then
        gdicClassAttendance.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Add_SystemStatus(ByRef obj As clsSystemStatus) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicSystemStatus.ContainsKey(key) Then
        gdicSystemStatus.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Add_Maintenance(ByRef obj As clsMAINTENANCE) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicMaintenance.ContainsKey(key) Then
        gdicMaintenance.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Add_MaintenanceDTL(ByRef obj As clsMAINTENANCE_DTL) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicMaintenance_DTL.ContainsKey(key) Then
        gdicMaintenance_DTL.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Add_MaintenanceStatus(ByRef obj As clsMAINTENANCE_STATUS) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicMaintenance_Status.ContainsKey(key) Then
        gdicMaintenance_Status.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '加入dicAlarm
  Public Function O_Add_Alarm(ByRef obj As clsALARM) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicAlarm.ContainsKey(key) Then
        gdicAlarm.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function  '加入dicAlarm
  Public Function O_Add_CT_PO_DTL(ByRef obj As clsHOST_CT_TMP_PO_DTL) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicCT_PO_DTL.ContainsKey(key) Then
        gdicCT_PO_DTL.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '加入dicOwner
  Public Function O_Add_Owner(ByRef obj As clsOwner) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicOwner.ContainsKey(key) Then
        gdicOwner.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '資料從Dictionary刪除

  ''從gcolPO刪除objPO
  'Public Function O_Remove_PO(ByRef obj As clsPO) As Boolean
  '  Try
  '    Dim key As String = obj.gid()
  '    If gdicPO.ContainsKey(key) Then
  '      gdicPO.TryRemove(key, obj)
  '    Else
  '      SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''從gcolPO_DTL刪除objPO_DTL
  'Public Function O_Remove_PO_DTL(ByRef obj As clsPO_DTL) As Boolean
  '  Try
  '    Dim key As String = obj.gid()
  '    If gdicPO_DTL.ContainsKey(key) Then
  '      gdicPO_DTL.TryRemove(key, obj)
  '    Else
  '      SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''從gcolPO_Line刪除objPO_Line
  'Public Function O_Remove_PO_Line(ByRef obj As clsPO_LINE) As Boolean
  '  Try
  '    Dim key As String = obj.gid()
  '    If gdicPO_Line.ContainsKey(key) Then
  '      gdicPO_Line.TryRemove(key, obj)
  '    Else
  '      SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  '從gcolCLine_Area刪除objCLine_Area
  Public Function O_Remove_Line_Area(ByRef obj As clsLine_Area) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicLine_Area.ContainsKey(key) Then
        gdicLine_Area.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gcolCLine刪除objCLine
  Public Function O_Remove_Line(ByRef obj As clsLine_Status) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicLine.ContainsKey(key) Then
        gdicLine.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gcolCLineInfo刪除objCLineInfo
  Public Function O_Remove_LineInfo(ByRef obj As clsLineInfo) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicLineInfo.ContainsKey(key) Then
        gdicLineInfo.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gcolCLineProduction刪除objCLineProduction
  Public Function O_Remove_CLineProduction(ByRef obj As clsLineProduction_Info) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicLineProduction_Info.ContainsKey(key) Then
        gdicLineProduction_Info.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gcolCProduce_Info刪除objCProduce_Info
  Public Function O_Remove_Produce_Info(ByRef obj As clsProduce_Info) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicProduce_Info.ContainsKey(key) Then
        gdicProduce_Info.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Remove_Class(ByRef obj As clsClass) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicClass.ContainsKey(key) Then
        gdicClass.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Remove_ClassAssignation(ByRef obj As clsCLASS_ASSIGNATION) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicClassAssignation.ContainsKey(key) Then
        gdicClassAssignation.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Remove_ClassAttendance(ByRef obj As clsCLASS_ATTENDANCE) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicClassAttendance.ContainsKey(key) Then
        gdicClassAttendance.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Remove_SystemStatus(ByRef obj As clsSystemStatus) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicSystemStatus.ContainsKey(key) Then
        gdicSystemStatus.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Remove_Maintenance(ByRef obj As clsMAINTENANCE) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicMaintenance.ContainsKey(key) Then
        gdicMaintenance.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Remove_MaintenanceDTL(ByRef obj As clsMAINTENANCE_DTL) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicMaintenance_DTL.ContainsKey(key) Then
        gdicMaintenance_DTL.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Remove_MaintenanceStatus(ByRef obj As clsMAINTENANCE_STATUS) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicMaintenance_Status.ContainsKey(key) Then
        gdicMaintenance_Status.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從gdicAlarm刪除Alarm
  Public Function O_Remove_Alarm(ByRef obj As clsALARM) As Boolean
    Try
      Dim key As String = obj.gid()
      If gdicAlarm.ContainsKey(key) Then
        gdicAlarm.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gdicCT_PO_DTL刪除CT_PO_DTL
  Public Function O_Remove_CT_PO_DTL(ByRef obj As clsHOST_CT_TMP_PO_DTL) As Boolean
    Try
      Dim key As String = obj.gid()
      If gdicCT_PO_DTL.ContainsKey(key) Then
        gdicCT_PO_DTL.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從gdicAlarm刪除Owner
  Public Function O_Remove_Owner(ByRef obj As clsOwner) As Boolean
    Try
      Dim key As String = obj.gid()
      If gdicOwner.ContainsKey(key) Then
        gdicOwner.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '取得Dictionary內的資料

  '從gcolCLine_Area取得指定的objCLine_Area
  Public Function O_Get_Line_Area(ByVal Factory_No As String, ByVal Area_No As String,
                                 Optional ByRef RetObj As clsLine_Area = Nothing) As Boolean
    Try
      Dim key As String = clsLine_Area.Get_Combination_Key(Factory_No, Area_No)
      Dim obj As clsLine_Area
      If gdicLine_Area.ContainsKey(key) Then
        obj = gdicLine_Area.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從gcolCLine_Area取得指定的objCLine_Area
  Public Function O_Get_Line_Area_By_AREA_TYPE1_AREA_INDEX(ByVal AREA_TYPE1 As String, ByVal AREA_INDEX As String,
                                 Optional ByRef RetObj As clsLine_Area = Nothing) As Boolean
    Try
      Dim dicLine_Area = gdicLine_Area.Where(Function(_obj)
                                               If _obj.Value.AREA_TYPE1 = AREA_TYPE1 AndAlso _obj.Value.AREA_INDEX = AREA_INDEX Then Return True
                                               Return False
                                             End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)

      If dicLine_Area IsNot Nothing AndAlso dicLine_Area.Any = True Then
        RetObj = dicLine_Area.First.Value
        Return True
      End If

      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gcolCLine取得指定的objCLine
  Public Function O_Get_CLine(ByVal Factory_No As String, ByVal Area_No As String,
                                      ByVal Device_No As String, ByVal Unt_ID As String,
                                 Optional ByRef RetObj As clsLine_Status = Nothing) As Boolean
    Try
      Dim key As String = clsLine_Status.Get_Combination_Key(Factory_No, Area_No, Device_No, Unt_ID)
      Dim obj As clsLine_Status
      If gdicLine.ContainsKey(key) Then
        obj = gdicLine.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gcolCLineInfo取得指定的objCLineInfo
  Public Function O_Get_CLineInfo(ByVal Factory_No As String,
                                  ByVal Area_No As String,
                                  ByVal Device_No As String,
                                  ByVal Unt_ID As String,
                                  ByVal MAINTENANCE_ID As String,
                                  ByVal FUCTION_ID As String,
                                  Optional ByRef RetObj As clsLineInfo = Nothing) As Boolean
    Try
      Dim key As String = clsLineInfo.Get_Combination_Key(Factory_No, Area_No, Device_No, Unt_ID, MAINTENANCE_ID, FUCTION_ID)
      Dim obj As clsLineInfo
      If gdicLineInfo.ContainsKey(key) Then
        obj = gdicLineInfo.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gcolCLineProduction_Info取得指定的objCLineProduction_Info
  Public Function O_Get_CLineProduction_Info(ByVal Factory_No As String, ByVal Area_No As String,
                                             ByVal Device_No As String, ByVal Unt_ID As String,
                                             Optional ByRef RetObj As clsLineProduction_Info = Nothing) As Boolean
    Try
      Dim key As String = clsLineProduction_Info.Get_Combination_Key(Factory_No, Area_No, Device_No, Unt_ID)
      Dim obj As clsLineProduction_Info
      If gdicLineProduction_Info.ContainsKey(key) Then
        obj = gdicLineProduction_Info.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從gcolCProduce_Info取得指定的objCProduce_Info
  Public Function O_Get_CProduce_Info(ByVal Factory_No As String, ByVal Area_No As String,
                                      ByVal PO_ID As String, ByVal SKU_NO As String,
                                 Optional ByRef RetObj As clsProduce_Info = Nothing) As Boolean
    Try
      Dim key As String = clsProduce_Info.Get_Combination_Key(Factory_No, Area_No, PO_ID, SKU_NO)
      Dim obj As clsProduce_Info
      If gdicProduce_Info.ContainsKey(key) Then
        obj = gdicProduce_Info.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_Class(ByVal Class_No As String,
                              Optional ByRef RetObj As clsClass = Nothing) As Boolean
    Try
      Dim key As String = clsClass.Get_Combination_Key(Class_No)
      Dim obj As clsClass
      If gdicClass.ContainsKey(key) Then
        obj = gdicClass.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_ClassAssignation(ByVal Factory_No As String,
                                         ByVal Area_No As String,
                                         ByVal Class_No As String,
                                         Optional ByRef RetObj As clsCLASS_ASSIGNATION = Nothing) As Boolean
    Try
      Dim key As String = clsCLASS_ASSIGNATION.Get_Combination_Key(Factory_No, Area_No, Class_No)
      Dim obj As clsCLASS_ASSIGNATION
      If gdicClassAssignation.ContainsKey(key) Then
        obj = gdicClassAssignation.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_ClassAttendance(ByVal Factory_No As String,
                                         ByVal Area_No As String,
                                         ByVal Class_No As String,
                                         Optional ByRef RetObj As clsCLASS_ATTENDANCE = Nothing) As Boolean
    Try
      Dim key As String = clsCLASS_ATTENDANCE.Get_Combination_Key(Class_No)
      Dim obj As clsCLASS_ATTENDANCE
      If gdicClassAttendance.ContainsKey(key) Then
        obj = gdicClassAttendance.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_SystemStatus(ByVal Status_No As String,
                                     Optional ByRef RetObj As clsSystemStatus = Nothing) As Boolean
    Try
      Dim key As String = clsSystemStatus.Get_Combination_Key(Status_No)
      Dim obj As clsSystemStatus
      If gdicSystemStatus.ContainsKey(key) Then
        obj = gdicSystemStatus.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_Maintenance(ByVal Factory_No As String,
                                    ByVal Device_No As String,
                                    ByVal Area_No As String,
                                    ByVal Unit_ID As String,
                                    ByVal Maintenance_ID As String,
                                    Optional ByRef RetObj As clsMAINTENANCE = Nothing) As Boolean
    Try
      Dim key As String = clsMAINTENANCE.Get_Combination_Key(Factory_No, Device_No, Area_No, Unit_ID, Maintenance_ID)
      Dim obj As clsMAINTENANCE
      If gdicMaintenance.ContainsKey(key) Then
        obj = gdicMaintenance.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_MaintenanceDTL(ByVal Factory_No As String,
                                       ByVal Device_No As String,
                                       ByVal Area_No As String,
                                       ByVal Unit_ID As String,
                                       ByVal Maintenance_ID As String,
                                       ByVal Function_ID As String,
                                       Optional ByRef RetObj As clsMAINTENANCE_DTL = Nothing) As Boolean
    Try
      Dim key As String = clsMAINTENANCE_DTL.Get_Combination_Key(Factory_No, Device_No, Area_No, Unit_ID, Maintenance_ID, Function_ID)
      Dim obj As clsMAINTENANCE_DTL
      If gdicMaintenance_DTL.ContainsKey(key) Then
        obj = gdicMaintenance_DTL.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_MaintenanceStatus(ByVal Factory_No As String,
                                          ByVal Device_No As String,
                                          ByVal Area_No As String,
                                          ByVal Unit_ID As String,
                                          ByVal Maintenance_ID As String,
                                          ByVal Function_ID As String,
                                          Optional ByRef RetObj As clsMAINTENANCE_STATUS = Nothing) As Boolean
    Try
      Dim key As String = clsMAINTENANCE_STATUS.Get_Combination_Key(Factory_No, Device_No, Area_No, Unit_ID, Maintenance_ID, Function_ID)
      Dim obj As clsMAINTENANCE_STATUS
      If gdicMaintenance_Status.ContainsKey(key) Then
        obj = gdicMaintenance_Status.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_Alarm(ByVal Factory_No As String, ByVal Area_No As String, ByVal Device_No As String, ByVal Unit_ID As String, ByVal Alarm_Code As String, Optional ByRef RetObj As clsALARM = Nothing) As Boolean
    Try
      Dim key As String = clsALARM.Get_Combination_Key(Factory_No, Area_No, Device_No, Unit_ID, Alarm_Code)

      If gdicAlarm.ContainsKey(key) Then
        RetObj = gdicAlarm.Item(key)
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_CT_PO_DTL(ByVal PO_ID As String, ByVal PO_SERIAL_NO As String, Optional ByRef RetObj As clsHOST_CT_TMP_PO_DTL = Nothing) As Boolean
    Try
      Dim key As String = clsHOST_CT_TMP_PO_DTL.Get_Combination_Key(PO_ID, PO_SERIAL_NO)

      If gdicCT_PO_DTL.ContainsKey(key) Then
        RetObj = gdicCT_PO_DTL.Item(key)
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '=================Private Function=======================
  Private Function Init_HandlingObject(ByRef RetMsg As String) As Boolean
    Try
      If I_BuildToolsConnection(RetMsg) = False Then
        Return False
      End If
      If I_Load_DB(RetMsg) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function

  Private Function I_BuildToolsConnection(ByRef RetMsg As String) As Boolean
    Try
      SendMessageToLog("BuildToolsConnection...", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      '檢查是否有設定
      Dim WMS_DBTool As eCA_DBTool.clsDBTool = Nothing
      Dim MCS_DBTool As eCA_DBTool.clsDBTool = Nothing
      Dim ERP_DBTool As eCA_DBTool.clsDBTool = Nothing
      If gDBTool.ContainsKey(enuDBEnum.WMS) = False Then
        MsgBox("尚未設定 " & enuDBEnum.WMS.ToString & " 的DB資訊")
      Else
        WMS_DBTool = gDBTool.Item(enuDBEnum.WMS)
      End If
      If gDBTool.ContainsKey(enuDBEnum.MCS) = False Then
        MsgBox("尚未設定 " & enuDBEnum.MCS.ToString & " 的DB資訊")
      Else
        MCS_DBTool = gDBTool.Item(enuDBEnum.MCS)
      End If
      If gDBTool.ContainsKey(enuDBEnum.ERP) = False Then
        MsgBox("尚未設定 " & enuDBEnum.ERP.ToString & " 的DB資訊")
      Else
        ERP_DBTool = gDBTool.Item(enuDBEnum.ERP)
      End If
      If gDBTool.Item(1).OpenConnection(RetMsg, True) Then

      End If
      If gDBTool.Item(2).OpenConnection(RetMsg, True) Then

      End If
      If gDBTool.Item(3).OpenConnection(RetMsg, True) Then

      End If
      Common_DBManagement.DBTool = WMS_DBTool
      WMS_CM_Line_AreaManagement.DBTool = WMS_DBTool
      WMS_CT_LINE_STATUSManagement.DBTool = WMS_DBTool
      WMS_CH_LINE_STATUS_HISTManagement.DBTool = WMS_DBTool
      WMS_CT_LINE_INFOManagement.DBTool = WMS_DBTool
      WMS_CH_LINE_HISTManagement.DBTool = WMS_DBTool
      WMS_CH_LINE_PRODUCTION_HISTManagement.DBTool = WMS_DBTool
      WMS_CT_LINE_PRODUCTION_INFOManagement.DBTool = WMS_DBTool
      WMS_CH_LINE_PRODUCTION_HISTManagement.DBTool = WMS_DBTool
      WMS_CT_PRODUCE_INFOManagement.DBTool = WMS_DBTool
      WMS_CT_PRODUCTION_REPORTManagement.DBTool = WMS_DBTool
      WMS_CH_PRODUCE_HISTManagement.DBTool = WMS_DBTool
      WMS_CH_PRODUCE_RESUME_HISTManagement.DBTool = WMS_DBTool
      WMS_M_DATA_REPORT_SETManagement.DBTool = WMS_DBTool
      WMS_CH_COUNT_MODIFY_HISTManagement.DBTool = WMS_DBTool
      WMS_M_MAINTENANCEManagement.DBTool = WMS_DBTool
      WMS_M_MAINTENANCE_DTLManagement.DBTool = WMS_DBTool
      WMS_T_MAINTENANCE_STATUSManagement.DBTool = WMS_DBTool

      WMS_CH_INVENTORY_COMPARISONManagement.DBTool = WMS_DBTool
      WMS_CT_INVENTORY_COMPARISONManagement.DBTool = WMS_DBTool
      'DB = WMS_DBTool

      WMS_T_PO_DTLManagement.DBTool = WMS_DBTool
      WMS_T_PO_LINEManagement.DBTool = WMS_DBTool
      WMS_T_PO_POSTINGManagement.DBTool = WMS_DBTool
      WMS_H_PO_POSTING_HISTManagement.DBTool = WMS_DBTool
      WMS_T_POManagement.DBTool = WMS_DBTool
      WMS_T_WO_DTLManagement.DBTool = WMS_DBTool

      WMS_T_ALARMManagement.DBTool = WMS_DBTool
      WMS_M_ClassManagement.DBTool = WMS_DBTool
      WMS_T_SystemStatusManagement.DBTool = WMS_DBTool
      WMS_T_STOCKTAKINGManagement.DBTool = WMS_DBTool
      WMS_T_PO_DTL_TRANSACTIONManagement.DBTool = WMS_DBTool

      WMS_CM_CLASS_ASSIGNATIONManagement.DBTool = WMS_DBTool
      WMS_CM_CLASS_ATTENDANCEManagement.DBTool = WMS_DBTool
      WMS_CM_Split_LabelManagement.DBTool = WMS_DBTool

      WMS_CH_CLASS_ASSIGNATION_HISTManagement.DBTool = WMS_DBTool
      WMS_CH_CLASS_ATTENDANCE_HISTManagement.DBTool = WMS_DBTool

      WMS_M_UUIDManagement.DBTool = WMS_DBTool
      WMS_M_OwnerManagement.DBTool = WMS_DBTool

      WMS_CT_VCManagement.DBTool = WMS_DBTool
      WMS_CT_VCMappingManagement.DBTool = WMS_DBTool

      WMS_CT_ACCOUNT_REPORTManagement.DBTool = WMS_DBTool
      WMS_CH_ACCOUNT_REPORTManagement.DBTool = WMS_DBTool
      WMS_T_Carrier_ItemManagement.DBTool = WMS_DBTool

      WMS_M_Packe_UnitManagement.DBTool = WMS_DBTool
      WMS_M_SKU_Packe_StructureManagement.DBTool = WMS_DBTool
      WMS_M_RETURN_SUPPLIER_SETTINGManagement.DBTool = WMS_DBTool
      'Interface
      GUI_H_Command_HistManagement.DBTool = WMS_DBTool
      GUI_T_CommandManagement.DBTool = WMS_DBTool
      HOST_H_Command_HistManagement.DBTool = WMS_DBTool
      HOST_T_CommandManagement.DBTool = WMS_DBTool
      MCS_H_Command_HistManagement.DBTool = WMS_DBTool
      MCS_T_CommandManagement.DBTool = MCS_DBTool
      WMS_H_GUI_Command_HistManagement.DBTool = WMS_DBTool
      WMS_T_GUI_CommandManagement.DBTool = WMS_DBTool
      WMS_H_HOST_Command_HistManagement.DBTool = WMS_DBTool
      WMS_T_HOST_CommandManagement.DBTool = WMS_DBTool
      WMS_H_MCS_Command_HistManagement.DBTool = WMS_DBTool
      WMS_M_SKUManagement.DBTool = WMS_DBTool
      WMS_T_MCS_CommandManagement.DBTool = MCS_DBTool
      WMS_M_Business_RuleManagement.DBTool = WMS_DBTool
      HOST_CT_TMP_PO_DTLManagement.DBTool = WMS_DBTool
      WMS_M_CarrierManagement.DBTool = WMS_DBTool
      WMS_T_Carrier_StatusManagement.DBTool = WMS_DBTool
      WMS_M_Item_LabelManagement.DBTool = WMS_DBTool
      WMS_CT_GUID_LabelManagement.DBTool = WMS_DBTool

      GUI_M_Message_Send_DTLManagement.DBTool = WMS_DBTool
      GUI_M_Message_SendManagement.DBTool = WMS_DBTool
      GUI_M_Message_TypeManagement.DBTool = WMS_DBTool
      GUI_M_UserManagement.DBTool = WMS_DBTool

      WMS_H_STOCKTAKING_CARRIERManagement.DBTool = WMS_DBTool
      WMS_H_STOCKTAKING_DTLManagement.DBTool = WMS_DBTool
      WMS_H_STOCKTAKINGManagement.DBTool = WMS_DBTool
      WMS_T_OUTBOUND_DTLManagement.DBTool = WMS_DBTool

      WMS_T_PO_MERGEManagement.DBTool = WMS_DBTool
      WMS_T_INBOUND_DTLManagement.DBTool = WMS_DBTool
      WMS_M_SLManagement.DBTool = WMS_DBTool

      'HOST_T_COMMAND_REPORTManagement.DBTool = WMS_DBTool
      WMS_T_COMMAND_REPORT.DBTool = WMS_DBTool

      ERP_DBManagement.DBTool = ERP_DBTool
      PURTCManagement.DBTool = ERP_DBTool
      PURTDManagement.DBTool = ERP_DBTool
      PURTEManagement.DBTool = ERP_DBTool
      PURTFManagement.DBTool = ERP_DBTool
      MOCTAManagement.DBTool = ERP_DBTool
      MOCTBManagement.DBTool = ERP_DBTool
      MOCTOManagement.DBTool = ERP_DBTool
      MOCTPManagement.DBTool = ERP_DBTool
      '中介檔
      EPSXBManagement.DBTool = ERP_DBTool
      INVMBManagement.DBTool = ERP_DBTool
      INVXDManagement.DBTool = ERP_DBTool
      INVXFManagement.DBTool = ERP_DBTool
      MOCXBManagement.DBTool = ERP_DBTool
      MOCXDManagement.DBTool = ERP_DBTool
      PURXCManagement.DBTool = ERP_DBTool
      INVXBManagement.DBTool = ERP_DBTool

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  Private Function I_Load_DB(ByRef RetMsg As String) As Boolean
    Try

      SendMessageToLog("Load DB...", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If I_Load_DB_CT_PO_DTL(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load CT_PO_DTL TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If
      ''單據結構
      'If I_Load_DB_PO(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load WMS_T_PO TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If
      'If I_Load_DB_PO_DTL(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load WMS_T_PO_DTL TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If
      'If I_Load_DB_PO_Line(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load WMS_T_PO_Line TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If

      'If I_Load_DB_Business_Rule(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load Business_Rule TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If
      'If I_Load_DB_Alarm(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load WMS_T_Alarm TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If
      'If I_Load_DB_Owner(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load Owner TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If


      'If I_Load_DB_Class(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load WMS_M_Class TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If

      'If I_Load_DB_Class_Assignation(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load WMS_CM_CLASS_ASSIGNATION TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If

      'If I_Load_DB_Class_Attendance(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load WMS_CM_CLASS_ATTENDANCE TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If

      'If I_Load_DB_System_Status(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load WMS_T_System_Status TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If
      'If I_Load_DB_Maintenance(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load WMS_M_Maintenance TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If
      'If I_Load_DB_MaintenanceDTL(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load WMS_M_Maintenance_DTL TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If
      'If I_Load_DB_MaintenanceStatus(RetMsg) = False Then
      '  SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  SendMessageToLog("Load WMS_T_Maintenance_Status TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'End If
      If I_Load_DB_System_Status(RetMsg) = False Then
        SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        SendMessageToLog("Load System_Status TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      If I_Load_DB_Business_Rule(RetMsg) = False Then
        SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        SendMessageToLog("Load Business_Rule TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

      If I_Load_DB_COMMAND_REPORT(RetMsg) = False Then
        SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        SendMessageToLog("Load WMS_T_COMMAND_REPORT TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

      If I_Load_DB_HOST_T_COMMAND_REPORT(RetMsg) = False Then
        SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        SendMessageToLog("Load HOST_T_COMMAND_REPORT TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

      'If I_Load_CT_Line_PRODUCTION_INFO(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load WMS_CT_Line_Production_Info TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If
      'If I_Load_CT_Line_Status(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load WMS_CT_Line_Status TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If
      'If I_Load_CT_Produce_Info(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load WMS_CT_Produce_Info TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If
      'If I_Load_CT_Produce_Info(RetMsg) = False Then
      '    SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    SendMessageToLog("Load WMS_CT_Produce_Info TABLE ERROR", eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    Return False
      'End If



      SendMessageToLog("Load DB Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function

  ''把Table WMS_T_PO的資料載進記憶體

  '把Table WMS_M_Business_Rule 的資料載進記憶體
  Private Function I_Load_DB_Business_Rule(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_M_Business_RuleManagement.GetWMS_M_Business_RuleDataListByALL()
      For Each objDB As clsBusiness_Rule In lst_Data
        objDB.Add_Relationship(Me)
      Next

      Return True
    Catch ex As Exception
      RetMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '把Table WMS_T_COMMAND 的資料載進記憶體
  Private Function I_Load_DB_COMMAND_REPORT(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_T_COMMAND_REPORT.GetWMS_T_COMMAND_REPORTDataListByALL()
      For Each objDB As clsCommandReport In lst_Data
        objDB.Add_Relationship(Me)
      Next

      Return True
    Catch ex As Exception
      RetMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '把Table HOST_T_COMMAND_REPORT 的資料載進記憶體
  Private Function I_Load_DB_HOST_T_COMMAND_REPORT(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = HOST_T_COMMAND_REPORTManagement.GetHOST_T_COMMAND_REPORTDataListByALL()
      For Each objDB As clsHOST_T_COMMAND_REPORT In lst_Data
        objDB.Add_Relationship(Me)
      Next

      Return True
    Catch ex As Exception
      RetMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '把Table WMS_M_Alarm 的資料載進記憶體
  Private Function I_Load_DB_Alarm(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_T_ALARMManagement.GetWMS_T_ALARMDataListByALL()
      For Each objDB As clsALARM In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table CT_PO_DTL 的資料載進記憶體
  Private Function I_Load_DB_CT_PO_DTL(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = HOST_CT_TMP_PO_DTLManagement.GetWCT_PO_DTLDataListByALL()
      For Each objDB As clsHOST_CT_TMP_PO_DTL In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_M_Owner 的資料載進記憶體
  Private Function I_Load_DB_Owner(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_M_OwnerManagement.GetdicOwnerByALL()
      For Each objDB As clsOwner In lst_Data.Values
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  Private Function I_Load_DB_Class(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_M_ClassManagement.GetWMS_M_ClassDataListByALL()
      For Each objDB As clsClass In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_CM_CLASS_ASSIGNATION 的資料載進記憶體
  Private Function I_Load_DB_Class_Assignation(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_CM_CLASS_ASSIGNATIONManagement.GetWMS_CM_ClassAssignationDataListByALL()
      For Each objDB As clsCLASS_ASSIGNATION In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_CM_CLASS_ATTENDANCE 的資料載進記憶體
  Private Function I_Load_DB_Class_Attendance(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_CM_CLASS_ATTENDANCEManagement.GetWMS_CM_ClassAttendanceDataListByALL()
      For Each objDB As clsCLASS_ATTENDANCE In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_T_System_Status 的資料載進記憶體
  Private Function I_Load_DB_System_Status(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_T_SystemStatusManagement.GetWMS_T_System_StatusDataListByALL()
      For Each objDB As clsSystemStatus In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_M_MAINTENANCE 的資料載進記憶體
  Private Function I_Load_DB_Maintenance(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_M_MAINTENANCEManagement.GetWMS_M_MaintenanceDataListByALL()
      For Each objDB As clsMAINTENANCE In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_M_MAINTENANCE_DTL 的資料載進記憶體
  Private Function I_Load_DB_MaintenanceDTL(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_M_MAINTENANCE_DTLManagement.GetWMS_M_MaintenanceDTLDataListByALL()
      For Each objDB As clsMAINTENANCE_DTL In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_T_MAINTENANCE_STATUS 的資料載進記憶體
  Private Function I_Load_DB_MaintenanceStatus(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_T_MAINTENANCE_STATUSManagement.GetWMS_T_MaintenanceStatusDataListByALL()
      For Each objDB As clsMAINTENANCE_STATUS In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_CM_Line_Area 的資料載進記憶體
  Private Function I_Load_CM_Line_Area(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_CM_Line_AreaManagement.GetWMS_CM_Line_AreaDataListByALL()
      For Each objDB As clsLine_Area In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_CT_Line_INFO 的資料載進記憶體
  Private Function I_Load_CT_Line_INFO(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_CT_LINE_INFOManagement.GetWMS_CT_LINE_INFODataListByALL()
      For Each objDB As clsLineInfo In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_CT_Line_PRODUCTION_INFO 的資料載進記憶體
  Private Function I_Load_CT_Line_PRODUCTION_INFO(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_CT_LINE_PRODUCTION_INFOManagement.GetWMS_CT_LINE_PRODUCTION_INFODataListByALL()
      For Each objDB As clsLineProduction_Info In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_CT_Line_Status 的資料載進記憶體
  Private Function I_Load_CT_Line_Status(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_CT_LINE_STATUSManagement.GetWMS_CT_LINE_STATUSDataListByALL()
      For Each objDB As clsLine_Status In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function
  '把Table WMS_CT_Produce_Info 的資料載進記憶體
  Private Function I_Load_CT_Produce_Info(ByRef RetMsg As String) As Boolean
    Try
      Dim lst_Data = WMS_CT_PRODUCE_INFOManagement.GetWMS_C_PRODUCE_INFODataListByALL()
      For Each objDB As clsProduce_Info In lst_Data
        objDB.Add_Relationship(Me)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      RetMsg = ex.ToString
      Return False
    End Try
  End Function

  '取得MaintenanceStatus
  Public Function O_Get_dicMantenaceStatusByFactoryNo_DeviceNo_AreaNo_UnitID_MaintenanceID_FunctionID(ByVal Factory_No As String, ByVal Device_No As String, ByVal Area_No As String,
                                                                                                     ByVal Unit_ID As String, ByVal Maintenance_ID As String, ByVal Function_ID As String,
                                                                                                     Optional ByRef ret_dic As Dictionary(Of String, clsMAINTENANCE_STATUS) = Nothing) As Boolean
    Try
      ret_dic = gdicMaintenance_Status.Where(Function(obj)
                                               If Factory_No.Equals("") = False AndAlso obj.Value.FACTORY_NO <> Factory_No Then
                                                 Return False
                                               End If
                                               If Device_No.Equals("") = False AndAlso obj.Value.DEVICE_NO <> Device_No Then
                                                 Return False
                                               End If
                                               If Area_No.Equals("") = False AndAlso obj.Value.AREA_NO <> Area_No Then
                                                 Return False
                                               End If
                                               If Unit_ID.Equals("") = False AndAlso obj.Value.UNIT_ID <> Unit_ID Then
                                                 Return False
                                               End If
                                               If Maintenance_ID.Equals("") = False AndAlso obj.Value.MAINTENANCE_ID <> Maintenance_ID Then
                                                 Return False
                                               End If
                                               If Function_ID.Equals("") = False AndAlso obj.Value.FUNCTION_ID <> Function_ID Then
                                                 Return False
                                               End If
                                               Return True
                                             End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得Maintenance
  Public Function O_Get_dicMaintenanceByFactoryNo_DeviceNo_AreraNo_UnitID_MaintenanceID(ByVal Factory_No As String,
                                                                                        ByVal Device_No As String,
                                                                                        ByVal Area_No As String,
                                                                                        ByVal Unit_ID As String,
                                                                                        ByVal Maintenance_ID As String,
                                                                                        Optional ByRef ret_dic As Dictionary(Of String, clsMAINTENANCE) = Nothing) As Boolean
    Try
      ret_dic = gdicMaintenance.Where(Function(obj)
                                        If Factory_No.Equals("") = False AndAlso obj.Value.FACTORY_NO <> Factory_No Then
                                          Return False
                                        End If
                                        If Device_No.Equals("") = False AndAlso obj.Value.DEVICE_NO <> Device_No Then
                                          Return False
                                        End If
                                        If Area_No.Equals("") = False AndAlso obj.Value.AREA_NO <> Area_No Then
                                          Return False
                                        End If
                                        If Unit_ID.Equals("") = False AndAlso obj.Value.UNIT_ID <> Unit_ID Then
                                          Return False
                                        End If
                                        If Maintenance_ID.Equals("") = False AndAlso obj.Value.MAINTENANCE_ID <> Maintenance_ID Then
                                          Return False
                                        End If
                                        Return True
                                      End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得MaintenanceDTL
  Public Function O_Get_dicMaintenanceDTLByFactoryNo_DeviceNo_AreaNo_UnitID_MaintenanceID_FunctionID(ByVal Factory_No As String,
                                                                                                     ByVal Device_No As String,
                                                                                                     ByVal Area_No As String,
                                                                                                     ByVal Unit_ID As String,
                                                                                                     ByVal Maintenance_ID As String,
                                                                                                     ByVal Function_ID As String,
                                                                                                     Optional ByRef ret_dic As Dictionary(Of String, clsMAINTENANCE_DTL) = Nothing) As Boolean
    Try
      ret_dic = gdicMaintenance_DTL.Where(Function(obj)
                                            If Factory_No.Equals("") = False AndAlso obj.Value.FACTORY_NO <> Factory_No Then
                                              Return False
                                            End If
                                            If Device_No.Equals("") = False AndAlso obj.Value.DEVICE_NO <> Device_No Then
                                              Return False
                                            End If
                                            If Area_No.Equals("") = False AndAlso obj.Value.AREA_NO <> Area_No Then
                                              Return False
                                            End If
                                            If Unit_ID.Equals("") = False AndAlso obj.Value.UNIT_ID <> Unit_ID Then
                                              Return False
                                            End If
                                            If Maintenance_ID.Equals("") = False AndAlso obj.Value.MAINTENANCE_ID <> Maintenance_ID Then
                                              Return False
                                            End If
                                            If Function_ID.Equals("") = False AndAlso obj.Value.FUNCTION_ID <> Function_ID Then
                                              Return False
                                            End If
                                            Return True
                                          End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得Production_Info的資訊
  Public Function O_Get_dicProductionInfoByPOID(ByVal PO_ID As String,
                                                Optional ByRef ret_dic As Dictionary(Of String, clsProduce_Info) = Nothing) As Boolean
    Try
      ret_dic = gdicProduce_Info.Where(Function(obj)
                                         If PO_ID.Equals("") = False AndAlso obj.Value.PO_ID <> PO_ID Then
                                           Return False
                                         End If
                                         Return True
                                       End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得ClassAttendance
  Public Function O_Get_dicClassAttendanceByClassNo(ByVal Class_No As String,
                                                    Optional ByRef ret_dic As Dictionary(Of String, clsCLASS_ATTENDANCE) = Nothing) As Boolean
    Try
      ret_dic = gdicClassAttendance.Where(Function(obj)
                                            If Class_No.Equals("") = False AndAlso obj.Value.CLASS_NO <> Class_No Then
                                              Return False
                                            End If
                                            Return True
                                          End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得ClassAssgnation
  Public Function O_Get_DicClassAssignationByFactoryNo_AreaNo_ClassNo(ByVal Factory_No As String,
                                                                      ByVal Area_No As String,
                                                                      ByVal Class_No As String,
                                                                      Optional ByRef ret_dic As Dictionary(Of String, clsCLASS_ASSIGNATION) = Nothing) As Boolean
    Try
      ret_dic = gdicClassAssignation.Where(Function(obj)
                                             If Factory_No.Equals("") = False AndAlso obj.Value.FACTORY_NO <> Factory_No Then
                                               Return False
                                             End If
                                             If Area_No.Equals("") = False AndAlso obj.Value.AREA_NO <> Area_No Then
                                               Return False
                                             End If
                                             If Class_No.Equals("") = False AndAlso obj.Value.CLASS_NO <> Class_No Then
                                               Return False
                                             End If
                                             Return True
                                           End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取線別數量的資料
  Public Function O_GetDB_Line_Production_HIST(ByVal StartTime As String, ByVal EndTime As String,
                                 Optional ByRef ret_lst As List(Of clsLineProduction_Hist) = Nothing) As Boolean
    Try
      ret_lst = WMS_CH_LINE_PRODUCTION_HISTManagement.GetclsLineProductionHISTByHistTime(StartTime, EndTime)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取UUID的資料
  Public Function O_GetDB_UUID(ByVal UUID_NO As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsUUID) = Nothing) As Boolean
    Try
      ret_dic = WMS_M_UUIDManagement.GetclsUUIDListByUUID_NO(UUID_NO)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB取得PO的資料，根據傳入的條件
  Public Function O_GetDB_dicPOByPOID_OrderType(ByVal PO_ID As String,
                                                ByVal ORDER_TYPE As enuOrderType,
                                                Optional ByRef ret_dic As Dictionary(Of String, clsPO) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_POManagement.GetclsWMS_T_POListByPO_ID_OrderType(PO_ID, ORDER_TYPE)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO的資料，根據傳入的條件
  Public Function O_GetDB_dicPOByPOID(ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_POManagement.GetPODictionaryByPOID(PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO的資料，根據傳入的條件
  Public Function O_GetDB_dicOutboundDTLByPOID(ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsOUTBOUND_DTL) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_OUTBOUND_DTLManagement.GetDataDicByPO_ID(PO_ID)
      If ret_dic IsNot Nothing AndAlso ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicInboundDTLByPOID(ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsINBOUND_DTL) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_INBOUND_DTLManagement.GetDataDicByPO_ID(PO_ID)
      If ret_dic IsNot Nothing AndAlso ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO的資料，根據傳入的條件
  Public Function O_GetDB_dicPODTLByUniqueID(ByVal UniqueID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_DTL) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_DTLManagement.GetclsWMS_T_PO_DTLListByUniqueID(UniqueID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO的資料，根據傳入的條件
  Public Function O_GetDB_dicPOByALL(Optional ByRef ret_dic As Dictionary(Of String, clsPO) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_POManagement.GetPODictionaryByALL()
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicPOByWO_Type(ByVal WO_TYPE As enuWOType, Optional ByRef ret_dic As Dictionary(Of String, clsPO) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_POManagement.GetPODictionaryByWO_TYPE(WO_TYPE)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Get_dicPOByPO_ID(ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_POManagement.GetclsWMS_T_POListByPO_ID(PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_DTL的資料，根據傳入的條件
  Public Function O_Get_dicPODTLByPOID_POLineNo_SERIAL_NO(ByVal PO_ID As String, ByVal PO_Line_No As String, ByVal PO_SERIAL_NO As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_DTL) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_DTLManagement.GetPO_DTLDictionaryByPOID_LINE_NO_Serial_No(PO_ID, PO_Line_No, PO_SERIAL_NO)
      If ret_dic IsNot Nothing And ret_dic.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_Line的資料，根據傳入的條件
  Public Function O_Get_dicPOLineByPOID_POLineNo(ByVal PO_ID As String, ByVal PO_Line_No As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_LINE) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_LINEManagement.GetPOLineDictionaryByPOID_POLineNo(PO_ID, PO_Line_No)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO的資料，根據傳入的條件
  Public Function O_GetDB_dicHPOBydicPO_ID(ByVal dicPO_ID As Dictionary(Of String, String),
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_POManagement.GetHPODictionaryBydicPOID(dicPO_ID)

      If ret_dic IsNot Nothing Or ret_dic.Count = 0 Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO的資料，根據傳入的條件
  Public Function O_GetDB_dicPOBydicPO_ID(ByVal dicPO_ID As Dictionary(Of String, String),
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_POManagement.GetPODictionaryBydicPOID(dicPO_ID)
      If ret_dic IsNot Nothing Or ret_dic.Count = 0 Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_DTL的資料，根據傳入的條件
  Public Function O_Get_dicPODTLBydicPO_ID(ByVal dicPO_ID As Dictionary(Of String, String),
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_DTL) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_DTLManagement.GetPO_DTLDictionaryBydicPOID(dicPO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_Line的資料，根據傳入的條件
  Public Function O_Get_dicPOLineBydicPO_ID(ByVal dicPO_ID As Dictionary(Of String, String),
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_LINE) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_LINEManagement.GetPOLineDictionaryBydicPOID(dicPO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取UUID的資料
  Public Function O_Get_UUID(ByVal UUID_NO As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsUUID) = Nothing) As Boolean
    Try
      ret_dic = WMS_M_UUIDManagement.GetclsUUIDListByUUID_NO(UUID_NO)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Get_dicPOByPO_ID_ORDER_TYPE(ByVal PO_ID As String, ByVal ORDER_TYPE As enuOrderType,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_POManagement.GetclsWMS_T_POListByPO_ID_OrderType(PO_ID, ORDER_TYPE)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_Line的資料，根據傳入的條件
  Public Function O_GetDB_dicPOLineByPOID_POLineNo(ByVal PO_ID As String, ByVal PO_Line_No As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_LINE) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_LINEManagement.GetPOLineDictionaryByPOID_POLineNo(PO_ID, PO_Line_No)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_Line的資料，根據傳入的條件
  Public Function O_GetDB_dicPOLineByAll(Optional ByRef ret_dic As Dictionary(Of String, clsPO_LINE) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_LINEManagement.GetdicWMS_T_PO_LINEListByALL()
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_Line的資料，根據傳入的條件
  Public Function O_GetDB_dicPOLineBydicPO_ID(ByVal dicPO_ID As Dictionary(Of String, String),
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_LINE) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_LINEManagement.GetPOLineDictionaryBydicPOID(dicPO_ID)
      If ret_dic IsNot Nothing Or ret_dic.Any = False Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_DTL的資料，根據傳入的條件
  Public Function O_GetDB_dicPODTLByPOID_POSerialNo(ByVal PO_ID As String, ByVal PO_Serial_No As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_DTL) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_DTLManagement.GetPO_DTLDictionaryByPOID_PO_Serial_No(PO_ID, PO_Serial_No)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  Public Function O_GetDB_lstWODTLByWO_ID(ByVal WO_ID As String,
                                          Optional ByRef ret_lst As List(Of clsWO_DTL) = Nothing) As Boolean
    Try
      ret_lst = WMS_T_WO_DTLManagement.GetWMS_T_WO_DTLDataListByWO_ID(WO_ID)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取PO_DTL的資料，根據傳入的條件
  Public Function O_GetDB_dicPODTLBydicPO_ID(ByVal dicPO_ID As Dictionary(Of String, String),
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_DTL) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_DTLManagement.GetPO_DTLDictionaryBydicPOID(dicPO_ID)
      If ret_dic IsNot Nothing Or ret_dic.Count = 0 Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_DTL的資料，根據傳入的條件
  Public Function O_GetDB_dicPO_MergeBydicPO_ID(ByVal dicPO_ID As Dictionary(Of String, String),
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_MERGE) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_MERGEManagement.GetPO_MergeDictionaryBydicPOID(dicPO_ID)
      If ret_dic IsNot Nothing Or ret_dic.Count = 0 Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取PO_DTL的資料，根據傳入的條件
  Public Function O_GetDB_dicPO_MergeByPO_ID_PO_Serial_No(ByVal PO_ID As String, ByVal PO_Serial_NO As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPO_MERGE) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_MERGEManagement.GetPO_MergeDictionaryByPOID_PO_Serial_No(PO_ID, PO_Serial_NO)
      If ret_dic IsNot Nothing Or ret_dic.Count = 0 Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_DTL的資料，根據傳入的條件
  Public Function O_GetDB_dicPODTLByALL(Optional ByRef ret_dic As Dictionary(Of String, clsPO_DTL) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_DTLManagement.GetPO_DTLDictionaryByALL()
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_DTL的資料，根據傳入的條件
  Public Function O_GetDB_dicPODTLTRANSACTIONBydicPO_ID(ByVal dicPO_ID As Dictionary(Of String, String),
                                 Optional ByRef ret_dic As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_DTL_TRANSACTIONManagement.GetPO_DTL_TRNASACTONDictionaryBydicPOID(dicPO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_lstSKUBydicSKUNo(ByVal dicSKU_No As Dictionary(Of String, String),
                                       Optional ByRef ret_dic As Dictionary(Of String, clsSKU) = Nothing) As Boolean
    Try
      ret_dic = WMS_M_SKUManagement.GetdicSKUListBydicSKUNo(dicSKU_No)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取Owner的資料，根據傳入的條件
  Public Function O_GetDB_dicOwnerByAll(Optional ByRef ret_dic As Dictionary(Of String, clsOwner) = Nothing) As Boolean
    Try
      ret_dic = WMS_M_OwnerManagement.GetdicOwnerByALL()

      If ret_dic IsNot Nothing And ret_dic.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_dicSKUByAll(Optional ByRef ret_lst As Dictionary(Of String, clsSKU) = Nothing) As Boolean
    Try
      ret_lst = WMS_M_SKUManagement.GetWMS_M_SKUDataListByALL()

      If ret_lst IsNot Nothing And ret_lst.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_GetDB_dicSLByAll(Optional ByRef ret_lst As Dictionary(Of String, clsSL) = Nothing) As Boolean
    Try
      ret_lst = WMS_M_SLManagement.GetWMS_M_SLListByALL()

      If ret_lst IsNot Nothing And ret_lst.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_dicSKUBySKU_ID1_SKU_ID2(ByVal SKU_ID1 As String,
                                                          ByVal SKU_ID2 As String,
                                                          Optional ByRef ret_lst As Dictionary(Of String, clsSKU) = Nothing) As Boolean
    Try
      ret_lst = WMS_M_SKUManagement.GetclsSKUListBySKU_NO_SKU_ID1_SKU_ID2(SKU_ID1, SKU_ID2)

      If ret_lst IsNot Nothing And ret_lst.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取標籤的資料，根據傳入的條件
  Public Function O_GetDB_dicSplit_LabelByAll(Optional ByRef ret_lst As Dictionary(Of String, clsWMS_CM_Split_Label) = Nothing) As Boolean
    Try
      ret_lst = WMS_CM_Split_LabelManagement.GetdicSplit_LabelByALL

      If ret_lst IsNot Nothing And ret_lst.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取VC的資料，根據傳入的條件
  Public Function O_GetDB_dicVCByAll(Optional ByRef ret_dic As Dictionary(Of String, clsWMS_CT_VC) = Nothing) As Boolean
    Try
      ret_dic = WMS_CT_VCManagement.GetdicVCDataByALL

      If ret_dic IsNot Nothing And ret_dic.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取Account的資料，根據傳入的條件
  Public Function O_GetDB_dicAccountByAll(Optional ByRef ret_dic As Dictionary(Of String, clsWMS_CT_ACCOUNT_REPORT) = Nothing) As Boolean
    Try
      ret_dic = WMS_CT_ACCOUNT_REPORTManagement.GetdicAccountDataByALL

      If ret_dic IsNot Nothing And ret_dic.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取VCMapping的資料，根據傳入的條件
  Public Function O_GetDB_dicVCMappingByAll(Optional ByRef ret_dic As Dictionary(Of String, clsWMS_CT_VCMapping) = Nothing) As Boolean
    Try
      ret_dic = WMS_CT_VCMappingManagement.GetdicVCMappingDataByALL

      If ret_dic IsNot Nothing And ret_dic.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicVCMappingBydicPOID(ByVal dicPOID As Dictionary(Of String, String), Optional ByRef ret_dic As Dictionary(Of String, clsWMS_CT_VCMapping) = Nothing) As Boolean
    Try
      ret_dic = WMS_CT_VCMappingManagement.GetdicVCMappingDataByALL

      If ret_dic IsNot Nothing And ret_dic.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_dicSKUBySKUNo(ByVal SKU_No As String,
                                         Optional ByRef ret_lst As Dictionary(Of String, clsSKU) = Nothing) As Boolean
    Try
      ret_lst = WMS_M_SKUManagement.GetdicSKUBySKUNo(SKU_No)

      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_dicSKUPackeStructureBySKU_NO(ByVal SKU_No As String,
                                         Optional ByRef ret_lst As Dictionary(Of String, clsMSKUPackeStructure) = Nothing) As Boolean
    Try
      ret_lst = WMS_M_SKU_Packe_StructureManagement.GetdicSKUPackeStructrueBySKUNo(SKU_No)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_POSTING的資料，根據傳入的條件
  Public Function O_Get_dicPO_POSTING_By_WO_ID(ByVal WO_ID As String, Optional ByRef ret_dic As Dictionary(Of String, clsPO_POSTING) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_POSTINGManagement.GetWMS_T_PO_POSTINGDataListByWO_ID(WO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_POSTING的資料，根據傳入的條件
  Public Function O_Get_dicPO_POSTING_By_dicWO_ID(ByVal dicWO_ID As Dictionary(Of String, String), Optional ByRef ret_dic As Dictionary(Of String, clsPO_POSTING) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_POSTINGManagement.GetWMS_T_PO_POSTINGDataListBydicWO_ID(dicWO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO_POSTING的資料，根據傳入的條件
  Public Function O_Get_dicPO_POSTING_By_All(Optional ByRef ret_dic As Dictionary(Of String, clsPO_POSTING) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_PO_POSTINGManagement.GetWMS_T_PO_POSTINGDataListByAll()
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取DataReportSet的資料，根據傳入的條件
  Public Function O_GetDB_dicDataReportSetByRoleID_FunctionID_DeviceNo_AreaNo_UnitID(ByVal Role_ID As String,
                                                                                   ByVal Fucntion_ID As String,
                                                                                   ByVal Device_No As String,
                                                                                   ByVal Area_No As String,
                                                                                   ByVal Unit_ID As String,
                                                                                   Optional ByRef ret_dic As Dictionary(Of String, clsDATA_REPORT_SET) = Nothing) As Boolean
    Try
      ret_dic = WMS_M_DATA_REPORT_SETManagement.SelectWMS_M_DATA_REPORT_SETDataByROLEID_FUNCTIONID_DEVICENO_AREANO_UNITID(Role_ID, Fucntion_ID, Device_No, Area_No, Unit_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取Stocktaking的資料，根據傳入的條件
  Public Function O_GetDB_dicStocktakingByStocktakingID(ByVal StocktakingID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsTSTOCKTAKING) = Nothing) As Boolean
    Try
      ret_dic = WMS_T_STOCKTAKINGManagement.GetWMS_T_STOCKTAKINGListByStocktaking_ID(StocktakingID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取Stocktaking的資料，根據傳入的條件
  Public Function O_GetDB_dicStocktaking_DTLByStocktakingID(ByVal StocktakingID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsTSTOCKTAKINGDTL) = Nothing) As Boolean
    Try
      'ret_dic = WMS_T_STOCKTAKINGManagement.GetWMS_T_STOCKTAKINGListByStocktaking_ID(StocktakingID)
      ret_dic = WMS_T_STOCKTAKING_DTLManagement.GetWMS_T_STOCKTAKING_DTLListByStocktaking_ID(StocktakingID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取Stocktaking的資料，根據傳入的條件
  Public Function O_GetDB_dicStocktaking_CarrierByStocktakingID(ByVal StocktakingID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsHSTOCKTAKINGCARRIER) = Nothing) As Boolean
    Try
      'ret_dic = WMS_T_STOCKTAKINGManagement.GetWMS_T_STOCKTAKINGListByStocktaking_ID(StocktakingID)
      ret_dic = WMS_H_STOCKTAKING_CARRIERManagement.GetWMS_H_STOCKTAKING_CARRIERdicByStocktaking_ID(StocktakingID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取Stocktaking的資料，根據傳入的條件
  Public Function O_GetDB_dicHStocktaking_DTLByStocktakingID(ByVal StocktakingID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsHSTOCKTAKINGDTL) = Nothing) As Boolean
    Try
      'ret_dic = WMS_T_STOCKTAKINGManagement.GetWMS_T_STOCKTAKINGListByStocktaking_ID(StocktakingID)
      ret_dic = WMS_H_STOCKTAKING_DTLManagement.GetWMS_H_STOCKTAKING_DTLdicByStocktaking_ID(StocktakingID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicHStocktakingByStocktakingID(ByVal StocktakingID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsHSTOCKTAKING) = Nothing) As Boolean
    Try
      'ret_dic = WMS_T_STOCKTAKINGManagement.GetWMS_T_STOCKTAKINGListByStocktaking_ID(StocktakingID)
      ret_dic = WMS_H_STOCKTAKINGManagement.GetWMS_H_STOCKTAKINGdicByStocktaking_ID(StocktakingID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取Stocktaking的資料，根據傳入的條件
  Public Function O_GetDB_dicStocktakingByAll(Optional ByRef ret_dic As Dictionary(Of String, clsTSTOCKTAKING) = Nothing) As Boolean
    Try
      '  ret_dic = WMS_T_STOCKTAKINGManagement.GetWMS_T_STOCKTAKINGListByStocktaking_ID
      '  If ret_dic IsNot Nothing Then
      Return True
      '  End If
      '  Return False
    Catch ex As Exception
      '  SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取CarrierItem的資料，根據傳入的條件
  Public Function GetCarrierItemByACCEPTING_STATUS_ItemCommon2IsNotNull(ByVal ACCEPTING_STATUS As enuAcceptingStatus) As Dictionary(Of String, clsCarrierItem)
    Try
      Dim dicCarrierItem As New Dictionary(Of String, clsCarrierItem)
      dicCarrierItem = WMS_T_Carrier_ItemManagement.GetWMS_T_Carrier_ItemDataListByACCEPTING_STATUS_ItemCommon2IsNotNull(ACCEPTING_STATUS)
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從DB抓取CarrierItem的資料，根據傳入的條件
  Public Function GetCarrierItemByACCEPTING_STATUS(ByVal ACCEPTING_STATUS As enuAcceptingStatus) As Dictionary(Of String, clsCarrierItem)
    Try
      Dim dicCarrierItem As New Dictionary(Of String, clsCarrierItem)
      dicCarrierItem = WMS_T_Carrier_ItemManagement.GetWMS_T_Carrier_ItemDataListByACCEPTING_STATUS(ACCEPTING_STATUS)
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從DB抓取CarrierItem的資料，根據傳入的條件
  Public Function GetCarrierItemByAll(ByRef dicCarrierItem As Dictionary(Of String, clsCarrierItem)) As Boolean
    Try
      dicCarrierItem = WMS_T_Carrier_ItemManagement.GetWMS_T_Carrier_Item_ALL()
      If dicCarrierItem Is Nothing Then
        Return False
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function GetCarrierStatusByAll() As Dictionary(Of String, clsCarrier)
    Try
      Dim dicCarrierItem As New Dictionary(Of String, clsCarrier)
      dicCarrierItem = WMS_T_Carrier_StatusManagement.GetWMS_T_Carrier_StatusDataListByALL()
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_dicItemLabelByGuid(ByVal Guid As String,
                                           Optional ByRef ret_lst As Dictionary(Of String, clsItemLabel) = Nothing) As Boolean
    Try
      'ret_lst = WMS_M_SKUManagement.GetdicSKUBySKUNo(ITEM_LABEL_ID)
      ret_lst = WMS_M_Item_LabelManagement.GetdicItemLabelByGuid(Guid)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_dicItemLabelByPO_ID(ByVal Guid As String,
                                           Optional ByRef ret_lst As Dictionary(Of String, clsItemLabel) = Nothing) As Boolean
    Try
      'ret_lst = WMS_M_SKUManagement.GetdicSKUBySKUNo(ITEM_LABEL_ID)
      ret_lst = WMS_M_Item_LabelManagement.GetdicItemLabelByPO_ID(Guid)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_dicItemLabelByPackage_ID(ByVal Package_ID As String,
                                           Optional ByRef ret_lst As Dictionary(Of String, clsItemLabel) = Nothing) As Boolean
    Try
      'ret_lst = WMS_M_SKUManagement.GetdicSKUBySKUNo(ITEM_LABEL_ID)
      ret_lst = WMS_M_Item_LabelManagement.GetdicItemLabelByPackage_ID(Package_ID)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PackeUnit的資料，根據傳入的條件
  Public Function O_GetDB_dicPackeUnitByPackeUnit(ByVal PackeUnit As String,
                                           Optional ByRef ret_lst As Dictionary(Of String, clsMPackeUnit) = Nothing) As Boolean
    Try
      ret_lst = WMS_M_Packe_UnitManagement.GetdicPackeUnitByPackeUnit(PackeUnit)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取SKUPackeStructure的資料，根據傳入的條件
  Public Function O_GetDB_dicPackeUnitByPackeUnit(ByVal SKU_NO As String,
                                           Optional ByRef ret_lst As Dictionary(Of String, clsMSKUPackeStructure) = Nothing) As Boolean
    Try
      ret_lst = WMS_M_SKU_Packe_StructureManagement.GetdicSKUPackeStructureBySKU(SKU_NO)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取SKU的資料，根據傳入的條件
  Public Function O_GetDB_dicGuid_LabelByUniqueID(ByVal UniqueID As String,
                                           Optional ByRef ret_lst As Dictionary(Of String, clsCTGUIDLabel) = Nothing) As Boolean
    Try
      'ret_lst = WMS_M_SKUManagement.GetdicSKUBySKUNo(ITEM_LABEL_ID)
      'ret_lst = WMS_M_Item_LabelManagement.GetdicItemLabelByUniqueID(UniqueID)
      ret_lst = WMS_CT_GUID_LabelManagement.GetdicGuid_LabelByUniqueID(UniqueID)
      If ret_lst IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '從DB抓取Owner的資料，根據傳入的條件
  Public Function O_GetDB_dicReturnSupplierSettingByAll(Optional ByRef ret_dic As Dictionary(Of String, clsRETURNSUPPLIERSETTING) = Nothing) As Boolean
    Try
      ret_dic = WMS_M_RETURN_SUPPLIER_SETTINGManagement.GetdicWMS_M_RETURN_SUPPLIER_SETTINGByALL()

      If ret_dic IsNot Nothing And ret_dic.Any = True Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '從DB抓取PO的資料，根據傳入的條件
  Public Function O_GetDB_dicPURTCByPOID(ByVal PO_TYPE As String,
                                        ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPURTC) = Nothing) As Boolean
    Try
      ret_dic = PURTCManagement.GetDataDictionaryByTC002(PO_TYPE, PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_GetDB_dicPURTDByPOID(ByVal PO_TYPE As String,
                                         ByVal PO_ID As String,
                                         ByVal No As String,
                                         Optional ByRef ret_dic As Dictionary(Of String, clsPURTD) = Nothing) As Boolean
    Try
      ret_dic = PURTDManagement.GetDataDictionaryByTD002(PO_TYPE, PO_ID, No)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicPURTEByPOID(ByVal PO_TYPE As String, ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsPURTE) = Nothing) As Boolean
    Try
      ret_dic = PURTEManagement.GetDataDictionaryByTE002(PO_TYPE, PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicPURTFByPOID(ByVal PO_TYPE As String,
                                         ByVal PO_ID As String,
                                         ByVal No As String,
                                         Optional ByRef ret_dic As Dictionary(Of String, clsPURTF) = Nothing) As Boolean
    Try
      ret_dic = PURTFManagement.GetDataDictionaryByTF002(PO_TYPE, PO_ID, No)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicMOCTAByPOID(ByVal PO_TYPE As String,
                                         ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsMOCTA) = Nothing) As Boolean
    Try
      ret_dic = MOCTAManagement.GetDataDictionaryByPO_ID(PO_TYPE, PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicMOCTBByPOID(ByVal PO_TYPE As String,
                                         ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsMOCTB) = Nothing) As Boolean
    Try
      ret_dic = MOCTBManagement.GetDataDictionaryByPO_ID(PO_TYPE, PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicMOCTOByPOID(ByVal PO_TYPE As String, ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsMOCTO) = Nothing) As Boolean
    Try
      ret_dic = MOCTOManagement.GetDataDictionaryByPO_ID(PO_TYPE, PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicMOCTPByPOID(ByVal PO_TYPE As String,
                                         ByVal PO_ID As String,
                                 Optional ByRef ret_dic As Dictionary(Of String, clsMOCTP) = Nothing) As Boolean
    Try
      ret_dic = MOCTPManagement.GetDataDictionaryByPO_ID(PO_TYPE, PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicEPSXBByXB010_IS_ZERO(Optional ByRef ret_dic As Dictionary(Of String, clsEPSXB) = Nothing) As Boolean
    Try
      ret_dic = EPSXBManagement.GetDataDictionaryByXB010_IS_ZERO
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicINVXBByXB008_IS_ZERO(Optional ByRef ret_dic As Dictionary(Of String, clsINVXB) = Nothing) As Boolean
    Try
      ret_dic = INVXBManagement.GetDataDictionaryByXB008_IS_ZERO
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicINVXBByXB008_IS_ZEROBY_SKU(ByVal SKU_NO As String, Optional ByRef ret_dic As Dictionary(Of String, clsINVXB) = Nothing) As Boolean
    Try
      ret_dic = INVXBManagement.GetDataDictionaryByXB008_IS_ZEROBySKU(SKU_NO)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_GetDB_dicMOCXDByXD011_IS_ZERO(Optional ByRef ret_dic As Dictionary(Of String, clsMOCXD) = Nothing) As Boolean
    Try
      ret_dic = MOCXDManagement.GetDataDictionaryByXD011_IS_ZERO
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicMOCXDByXD001_XD002(ByVal XD001 As String, ByVal XD002 As String, Optional ByRef ret_dic As Dictionary(Of String, clsMOCXD) = Nothing) As Boolean
    Try
      ret_dic = MOCXDManagement.GetDataDictionaryByXD001_XD002(XD001, XD002)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicMOCXDByXD001_XD002_XD013(ByVal XD001 As String, ByVal XD002 As String, ByVal XD013 As String, Optional ByRef ret_dic As Dictionary(Of String, clsMOCXD) = Nothing) As Boolean
    Try
      ret_dic = MOCXDManagement.GetDataDictionaryByXD001_XD002_XD013(XD001, XD002, XD013)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicMOCXDByKEY(ByVal XD001 As String, ByVal XD002 As String, ByVal XD013 As String, Optional ByRef ret_dic As Dictionary(Of String, clsMOCXD) = Nothing) As Boolean
    Try
      ret_dic = MOCXDManagement.GetDataDictionaryByKEY(XD001, XD002, XD013)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicEPSXBByXB001_XB002(ByVal XB001 As String, ByVal XB002 As String, Optional ByRef ret_dic As Dictionary(Of String, clsEPSXB) = Nothing) As Boolean
    Try
      ret_dic = EPSXBManagement.GetDataDictionaryByXB001_XB002(XB001, XB002)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicEPSXBByXB001_XB002_XB003(ByVal XB001 As String, ByVal XB002 As String, ByVal XB003 As String, Optional ByRef ret_dic As Dictionary(Of String, clsEPSXB) = Nothing) As Boolean
    Try
      ret_dic = EPSXBManagement.GetDataDictionaryByXB001_XB002_003(XB001, XB002, XB003)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicEPSXBByKEY(ByVal XB001 As String, ByVal XB002 As String, ByVal XB003 As String, Optional ByRef ret_dic As Dictionary(Of String, clsEPSXB) = Nothing) As Boolean
    Try
      ret_dic = EPSXBManagement.GetDataDictionaryByKEY(XB001, XB002, XB003)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicINVXDByXD009_IS_ZERO(Optional ByRef ret_dic As Dictionary(Of String, clsINVXD) = Nothing) As Boolean
    Try
      ret_dic = INVXDManagement.GetDataDictionaryByXD009_IS_ZERO
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicINVXDByPO_ID(ByVal PO_ID As String, Optional ByRef ret_dic As Dictionary(Of String, clsINVXD) = Nothing) As Boolean
    Try
      ret_dic = INVXDManagement.GetDataDictionaryByPO_ID(PO_ID)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicINVXFByXF009_IS_ZERO(Optional ByRef ret_dic As Dictionary(Of String, clsINVXF) = Nothing) As Boolean
    Try
      ret_dic = INVXFManagement.GetDataDictionaryByXF009_IS_ZERO
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicINVXFByXF001_XF002(ByVal XF001 As String, ByVal XF002 As String, Optional ByRef ret_dic As Dictionary(Of String, clsINVXF) = Nothing) As Boolean
    Try
      ret_dic = INVXFManagement.GetDataDictionaryByXF001_XF002(XF001, XF002)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicINVXFByKEY(ByVal XF001 As String, ByVal XF002 As String, ByVal XF004 As String, Optional ByRef ret_dic As Dictionary(Of String, clsINVXF) = Nothing) As Boolean
    Try
      ret_dic = INVXFManagement.GetDataDictionaryByKEY(XF001, XF002, XF004)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicINVXBByKEY(ByVal XB001 As String, Optional ByRef ret_dic As Dictionary(Of String, clsINVXB) = Nothing) As Boolean
    Try
      ret_dic = INVXBManagement.GetDataDictionaryByKEY(XB001)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_dicPURXCByKey(ByVal XC008 As String, ByVal XC009 As String, ByVal XC010 As String, ByVal XC016 As String, Optional ByRef ret_dic As Dictionary(Of String, clsPURXC) = Nothing) As Boolean
    Try
      ret_dic = PURXCManagement.GetDataDictionaryByKEY(XC008, XC009, XC010, XC016)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_GetDB_dicMOCXBByKey(ByVal XB010 As String, ByVal XB011 As String, ByVal XB013 As String, Optional ByRef ret_dic As Dictionary(Of String, clsMOCXB) = Nothing) As Boolean
    Try
      ret_dic = MOCXBManagement.GetDataDictionaryByKEY(XB010, XB011, XB013)
      If ret_dic IsNot Nothing Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

#Region "發送郵件"
  Public Function O_GetDB_GUI_Message(ByVal MESSAGE_TYPE As enuMessageType,
                                        Optional ByRef ret_dicMessage_Type As Dictionary(Of String, clsGUI_M_Message_Type) = Nothing,
                                        Optional ByRef ret_dicMessage_Send As Dictionary(Of String, clsGUI_M_Message_Send) = Nothing,
                                        Optional ByRef ret_dicMessage_Send_DTL As Dictionary(Of String, clsGUI_M_Message_Send_DTL) = Nothing) As Boolean
    Try
      ret_dicMessage_Type = GUI_M_Message_TypeManagement.GetGUI_M_Message_TypeDataListByKey_MESSAGE_TYPE(MESSAGE_TYPE)
      If ret_dicMessage_Type.Any = False Then
        Return False
      End If

      ret_dicMessage_Send = GUI_M_Message_SendManagement.GetGUI_M_Message_SendDataListByMessageType(MESSAGE_TYPE)
      If ret_dicMessage_Send.Any = False Then
        Return False
      End If

      Dim lstKey_no As New List(Of String)
      For Each objMessage_Send In ret_dicMessage_Send.Values
        lstKey_no.Add(objMessage_Send.KEY_NO)
      Next
      ret_dicMessage_Send_DTL = GUI_M_Message_Send_DTLManagement.GetGUI_M_Message_Send_DTLDataListBylstKEY_NO(lstKey_no)
      If ret_dicMessage_Send_DTL.Any = False Then
        Return False
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_GUI_M_USER(Optional ByRef ret_dic As Dictionary(Of String, clsGUI_M_User) = Nothing) As Boolean
    Try
      ret_dic = GUI_M_UserManagement.GetGUI_M_UserDataListByALL
      If ret_dic.Any = False Then
        Return False
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_GetDB_INVENTORY_COMPARISON_GetEndFlagByDB() As Boolean
    Try
      Return WMS_CT_INVENTORY_COMPARISONManagement.GetEndFlagByDB
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


#End Region

  '把CommandReport加入gdicCommand_Report
  Public Function O_Add_Command_Report(ByRef obj As clsCommandReport) As Boolean
    Try
      Dim key As String = obj.gid
      If Not gdicCommand_Report.ContainsKey(key) Then
        gdicCommand_Report.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_Command_Report(ByVal REPORT_SYSTEM_TYPE As enuSystemType, ByVal REPORT_SYSTEM_UUID As String, Optional ByRef RetObj As clsCommandReport = Nothing) As Boolean
    Try
      Dim key As String = clsCommandReport.Get_Combination_Key(REPORT_SYSTEM_TYPE, REPORT_SYSTEM_UUID)
      Dim obj As clsCommandReport
      If gdicCommand_Report.ContainsKey(key) Then
        obj = gdicCommand_Report.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '把Business_Rule加入gdicBusiness_Rule
  Public Function O_Remove_Command_Report(ByRef obj As clsCommandReport) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicCommand_Report.ContainsKey(key) Then
        gdicCommand_Report.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '把HOST_COMMAND_REPORT加入gdicHOST_T_COMMAND_REPORT
  Public Function O_Add_HOST_T_COMMAND_REPORT(ByRef obj As clsHOST_T_COMMAND_REPORT) As Boolean
    Try
      Dim key As String = obj.gid
      If Not gdicHOST_T_COMMAND_REPORT.ContainsKey(key) Then
        gdicHOST_T_COMMAND_REPORT.Add(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_HOST_T_COMMAND_REPORT(ByVal REPORT_SYSTEM_TYPE As enuSystemType, ByVal REPORT_SYSTEM_UUID As String, Optional ByRef RetObj As clsHOST_T_COMMAND_REPORT = Nothing) As Boolean
    Try
      Dim key As String = clsHOST_T_COMMAND_REPORT.Get_Combination_Key(REPORT_SYSTEM_TYPE, REPORT_SYSTEM_UUID)
      Dim obj As clsHOST_T_COMMAND_REPORT
      If gdicHOST_T_COMMAND_REPORT.ContainsKey(key) Then
        obj = gdicHOST_T_COMMAND_REPORT.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '把Business_Rule加入gdicBusiness_Rule
  Public Function O_Remove_HOST_T_COMMAND_REPORT(ByRef obj As clsHOST_T_COMMAND_REPORT) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicHOST_T_COMMAND_REPORT.ContainsKey(key) Then
        gdicHOST_T_COMMAND_REPORT.Remove(key)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function




End Class




