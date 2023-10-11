'2019/1/16 上午 11:55:06
'V1.0.0
'Tool
'狀態:Checked

Imports eCA_HostObject
Imports eCA_TransactionMessage


Module Module_T3F5U1_MaintenanceSet
  Public Function O_Process_Message(ByRef Receive_Msg As MSG_T3F5U1_MaintenanceSet,
                                    ByRef Result_Message As String) As Boolean
    Try
      '異動的資料
      Dim dicUpdate_Maintenace As New Dictionary(Of String, clsMAINTENANCE)
      Dim dicUpdate_Maintenace_DTL As New Dictionary(Of String, clsMAINTENANCE_DTL)

      Dim lstSql As New List(Of String) '儲存要更新的SQL，進行一次性更新
      '先進行資料邏輯檢查
      If Check_UpdateData(Receive_Msg, Result_Message) = False Then
        Return False
      End If
      '邏輯
      If Get_UpdateData(Receive_Msg, Result_Message, dicUpdate_Maintenace, dicUpdate_Maintenace_DTL) = False Then
        Return False
      End If
      '取得SQL
      If Get_SQL(Result_Message, dicUpdate_Maintenace, dicUpdate_Maintenace_DTL, lstSql) = False Then
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(Result_Message, dicUpdate_Maintenace, dicUpdate_Maintenace_DTL, lstSql) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Check_UpdateData(ByRef Receive_Msg As MSG_T3F5U1_MaintenanceSet,
                                    ByRef Result_Message As String) As Boolean
    Try
      For Each objWrokDataInfo In Receive_Msg.Body.MaintenanceList.MaintenanceInfo
        '檢查參數
        If objWrokDataInfo.FACTORY_NO = "" Then
          Result_Message = "FACTORY_NO is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If objWrokDataInfo.DEVICE_NO = "" Then
          Result_Message = "DEVICE_NO is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If objWrokDataInfo.AREA_NO = "" Then
          Result_Message = "AREA_NO is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If objWrokDataInfo.UNIT_ID = "" Then
          Result_Message = "UNIT_ID is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If objWrokDataInfo.MAINTENANCE_ID = "" Then
          Result_Message = "MAINTENANCE_ID is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If objWrokDataInfo.MAINTENANCE_NAME = "" Then
          Result_Message = "MAINTENANCE_NAME is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.CONTINUE_SEND) = False Then
          Result_Message = "CONTINUE_SEND is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.SEND_INTERVAL) = False Then
          Result_Message = "SEND_INTERVAL is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.SEND_TYPE) = False Then
          Result_Message = "SEND_TYPE is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.ENABLE) = False Then
          Result_Message = "ENABLE is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If


        '檢查明細參數
        For Each item In objWrokDataInfo.MaintenanceDTLList.MaintenanceDTLInfo
          If item.FUNCTION_ID = "" Then
            Result_Message = "FUNCTION_ID is empty"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If

          If IsNumeric(item.VALUE_TYPE) = False Then
            Result_Message = "VALUE_TYPE is empty"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If IsNumeric(item.NOTICE_TYPE) = False Then
            Result_Message = "NOTICE_TYPE is empty"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If IsNumeric(item.VALUE_UPDATE_TYPE) = False Then
            Result_Message = "VALUE_UPDATE_TYPE is empty"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        Next
      Next


      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Get_UpdateData(ByRef Receive_Msg As MSG_T3F5U1_MaintenanceSet,
                                  ByRef Result_Message As String,
                                  ByRef ret_dicUpdate_Maintenace As Dictionary(Of String, clsMAINTENANCE),
                                  ByRef ret_dicUpdate_Maintenace_DTL As Dictionary(Of String, clsMAINTENANCE_DTL)) As Boolean
    Try
      'logic
      '-先取的資料
      For Each MaintenanceInfo In Receive_Msg.Body.MaintenanceList.MaintenanceInfo
        Dim Facotry_no = MaintenanceInfo.FACTORY_NO
        Dim Device_No = MaintenanceInfo.DEVICE_NO
        Dim Area_No = MaintenanceInfo.AREA_NO
        Dim Unit_ID = MaintenanceInfo.UNIT_ID
        Dim Maintenance_ID = MaintenanceInfo.MAINTENANCE_ID
        Dim Maintenance_Name = MaintenanceInfo.MAINTENANCE_NAME  '給使用者看的名稱
        Dim Continue_Send = MaintenanceInfo.CONTINUE_SEND     '是否持續發送(Mosa設為0)
        Dim Send_Interval = MaintenanceInfo.SEND_INTERVAL     '發送間隔時間(S)
        Dim Send_Type = MaintenanceInfo.SEND_TYPE       '發送類型，固定為0
        Dim Enable As Boolean = IntegerConvertToBoolean(MaintenanceInfo.ENABLE)            '是否啟用，0:禁用/1:啟用
        Dim tmp_dicMaintenace As New Dictionary(Of String, clsMAINTENANCE)
        If gMain.objHandling.O_Get_dicMaintenanceByFactoryNo_DeviceNo_AreraNo_UnitID_MaintenanceID(Facotry_no, Device_No, Area_No, Unit_ID, Maintenance_ID, tmp_dicMaintenace) = False Then
          Result_Message = "Cant Get any Maintenance Data"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '-取得資料
        '修改Maintenance的資料
        If tmp_dicMaintenace.Any = True Then '-更新資料
          For Each objMaintenance As clsMAINTENANCE In tmp_dicMaintenace.Values
            If ret_dicUpdate_Maintenace.ContainsKey(objMaintenance.gid) = False Then
              Dim objNewMaintenance As clsMAINTENANCE = objMaintenance.Clone()
              objNewMaintenance.MAINTENANCE_NAME = Maintenance_Name
              objNewMaintenance.CONTINUE_SEND = Continue_Send
              objNewMaintenance.SEND_INTERVAL = Send_Interval
              objNewMaintenance.SEND_TYPE = Send_Type
              objNewMaintenance.ENABLE = Enable
              ret_dicUpdate_Maintenace.Add(objNewMaintenance.gid, objNewMaintenance)
            End If
          Next
        Else
          Result_Message = "Not Get any Maintenance Data"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '修改MaintenacneDTL的資料
        For Each MaintenanceDTLInfo In MaintenanceInfo.MaintenanceDTLList.MaintenanceDTLInfo
          Dim FUNCTION_ID = MaintenanceDTLInfo.FUNCTION_ID
          Dim VALUE_TYPE = MaintenanceDTLInfo.VALUE_TYPE
          Dim NOTICE_TYPE = MaintenanceDTLInfo.NOTICE_TYPE
          Dim HIGH_WATER_VALUE = MaintenanceDTLInfo.HIGH_WATER_VALUE
          Dim LOW_WATER_VALUE = MaintenanceDTLInfo.LOW_WATER_VALUE
          Dim STANDARD_VALUE = MaintenanceDTLInfo.STANDARD_VALUE
          Dim VALUE_RANGE = MaintenanceDTLInfo.VALUE_RANGE
          Dim MAINTENANCE_MESSAGE = MaintenanceDTLInfo.MAINTENANCE_MESSAGE
          Dim VALUE_SOURCE = MaintenanceDTLInfo.VALUE_SOURCE
          Dim VALUE_UPDATE_TYPE = MaintenanceDTLInfo.VALUE_UPDATE_TYPE
          Dim tmp_dicMaintenaceDTL As New Dictionary(Of String, clsMAINTENANCE_DTL)
          If gMain.objHandling.O_Get_dicMaintenanceDTLByFactoryNo_DeviceNo_AreaNo_UnitID_MaintenanceID_FunctionID(Facotry_no, Device_No, Area_No, Unit_ID, Maintenance_ID, FUNCTION_ID, tmp_dicMaintenaceDTL) = False Then
            Result_Message = "Cant Get any Maintenance DTL Data"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If tmp_dicMaintenaceDTL.Any = True Then
            For Each objMaintenanceDTL As clsMAINTENANCE_DTL In tmp_dicMaintenaceDTL.Values
              If ret_dicUpdate_Maintenace_DTL.ContainsKey(objMaintenanceDTL.gid) = False Then
                Dim objNewMaintenanceDTL As clsMAINTENANCE_DTL = objMaintenanceDTL.Clone()
                objNewMaintenanceDTL.VALUE_TYPE = VALUE_TYPE
                objNewMaintenanceDTL.NOTICE_TYPE = NOTICE_TYPE
                objNewMaintenanceDTL.HIGH_WATER_VALUE = HIGH_WATER_VALUE
                objNewMaintenanceDTL.LOW_WATER_VALUE = LOW_WATER_VALUE
                objNewMaintenanceDTL.STANDARD_VALUE = STANDARD_VALUE
                objNewMaintenanceDTL.VALUE_RANGE = VALUE_RANGE
                objNewMaintenanceDTL.MAINTENANCE_MESSAGE = MAINTENANCE_MESSAGE
                objNewMaintenanceDTL.VALUE_SOURCE = VALUE_SOURCE
                objNewMaintenanceDTL.VALUE_UPDATE_TYPE = VALUE_UPDATE_TYPE
                ret_dicUpdate_Maintenace_DTL.Add(objNewMaintenanceDTL.gid, objNewMaintenanceDTL)
              End If
            Next
          Else
            Result_Message = "Not Get any Maintenance DTL Data"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        Next
      Next
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef ret_dicUpdate_Maintenace As Dictionary(Of String, clsMAINTENANCE),
                           ByRef ret_dicUpdate_Maintenace_DTL As Dictionary(Of String, clsMAINTENANCE_DTL),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得要更新的SQL
      For Each obj In ret_dicUpdate_Maintenace.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          Result_Message = "Get Update Maintenace SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      '-
      For Each obj In ret_dicUpdate_Maintenace_DTL.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          Result_Message = "Get Update MaintenaceDTL SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Execute_DataUpdate(ByRef Result_Message As String,
                                      ByRef ret_dicUpdate_Maintenace As Dictionary(Of String, clsMAINTENANCE),
                                      ByRef ret_dicUpdate_Maintenace_DTL As Dictionary(Of String, clsMAINTENANCE_DTL),
                                      ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSql.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
        '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
        Result_Message = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '更新所有的SQL
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        Result_Message = " Update DB Failed"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '修改記憶體資料
      For Each objNew As clsMAINTENANCE In ret_dicUpdate_Maintenace.Values
        Dim obj As clsMAINTENANCE = Nothing
        If gMain.objHandling.gdicMaintenance.TryGetValue(objNew.gid, obj) = True Then
          obj.Update_To_Memory(objNew)
        End If
      Next
      For Each objNew As clsMAINTENANCE_DTL In ret_dicUpdate_Maintenace_DTL.Values
        Dim obj As clsMAINTENANCE_DTL = Nothing
        If gMain.objHandling.gdicMaintenance_DTL.TryGetValue(objNew.gid, obj) = True Then
          obj.Update_To_Memory(objNew)
        End If
      Next
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
