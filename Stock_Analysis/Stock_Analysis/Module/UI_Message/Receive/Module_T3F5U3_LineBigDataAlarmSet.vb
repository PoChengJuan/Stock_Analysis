'2019/1/16 上午 11:57:33
'V1.0.0
'Benny
'狀態:Open
Imports eCA_HOSTObject
Imports eCA_TransactionMessage
Module Module_T3F5U3_LineBigDataAlarmSet
  Public Function O_Process_Message(ByRef Receive_Msg As MSG_T3F5U3_LineBigDataAlarmSet, ByRef Result_Message As String) As Boolean
    Try
      '異動的資料
      Dim dicUpdate_M_DataReport As New Dictionary(Of String, clsDATA_REPORT_SET)
      Dim lstSql As New List(Of String) '儲存要更新的SQL，進行一次性更新
      '先進行資料邏輯檢查
      If Check_UpdateData(Receive_Msg, Result_Message) = False Then
        Return False
      End If
      '邏輯
      If Get_UpdateData(Receive_Msg, Result_Message, dicUpdate_M_DataReport) = False Then
        Return False
      End If
      '取得SQL
      If Get_SQL(Result_Message, dicUpdate_M_DataReport, lstSql) = False Then
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(Result_Message, lstSql) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Check_UpdateData(ByRef Receive_Msg As MSG_T3F5U3_LineBigDataAlarmSet, ByRef Result_Message As String) As Boolean
    Try
      For Each objWrokDataInfo In Receive_Msg.Body.RoleList.RoleInfo
        '檢查參數
        If objWrokDataInfo.ROLE_ID = "" Then
          Result_Message = "FACTORY_NO is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.ROLE_TYPE) = False Then
          Result_Message = "ROLE_TYPE is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If objWrokDataInfo.FUNCTION_ID = "" Then
          Result_Message = "FUNCTION_ID is empty"
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
        If IsNumeric(objWrokDataInfo.HIGH_WATER_VALUE) = False Then
          Result_Message = "HIGH_WATER_VALUE is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.LOW_WATER_VALUE) = False Then
          Result_Message = "LOW_WATER_VALUE is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.STANDARD_VALUE) = False Then
          Result_Message = "STANDARD_VALUE is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.VALUE_RANGE) = False Then
          Result_Message = "VALUE_RANGE is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.NOTICE_TYPE) = False Then
          Result_Message = "NOTICE_TYPE is empty"
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
        If IsNumeric(objWrokDataInfo.ENABLE) = False Then
          Result_Message = "ENABLE is empty"
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
  Private Function Get_UpdateData(ByRef Receive_Msg As MSG_T3F5U3_LineBigDataAlarmSet,
                                              ByRef Result_Message As String,
                                              ByRef dicUpdate_M_DataReport As Dictionary(Of String, clsDATA_REPORT_SET)) As Boolean
    Try
      'logic

      For Each item In Receive_Msg.Body.RoleList.RoleInfo
        Dim _tmp As New Dictionary(Of String, clsDATA_REPORT_SET)
        Dim ROLE_ID = item.ROLE_ID
        Dim ROLE_TYPE = item.ROLE_TYPE
        Dim FUNCTION_ID = item.FUNCTION_ID
        Dim DEVICE_NO = item.DEVICE_NO
        Dim AREA_NO = item.AREA_NO
        Dim UNIT_ID = item.UNIT_ID
        Dim HIGH_WATER_VALUE = item.HIGH_WATER_VALUE
        Dim LOW_WATER_VALUE = item.LOW_WATER_VALUE
        Dim STANDARD_VALUE = item.STANDARD_VALUE
        Dim VALUE_RANGE = item.VALUE_RANGE
        Dim NOTICE_TYPE = item.NOTICE_TYPE
        Dim CONTINUE_SEND = item.CONTINUE_SEND
        Dim SEND_INTERVAL = item.SEND_INTERVAL
        Dim ENABLE = item.ENABLE
        If gMain.objHandling.O_GetDB_dicDataReportSetByRoleID_FunctionID_DeviceNo_AreaNo_UnitID(ROLE_ID, FUNCTION_ID, DEVICE_NO, AREA_NO, UNIT_ID, _tmp) = False Then
          Result_Message = "Select From LineBigDataAlarmSet False"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        Else
          If _tmp.Count = 0 Then
            Result_Message = "Select From LineBigDataAlarmSet Count=0 False"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          Else
            For Each _item In _tmp
              _item.Value.ROLE_TYPE = ROLE_TYPE
              _item.Value.HIGH_WATER_VALUE = HIGH_WATER_VALUE
              _item.Value.LOW_WATER_VALUE = LOW_WATER_VALUE
              _item.Value.STANDARD_VALUE = STANDARD_VALUE
              _item.Value.VALUE_RANGE = VALUE_RANGE
              _item.Value.NOTICE_TYPE = NOTICE_TYPE
              _item.Value.CONTINUE_SEND = CONTINUE_SEND
              _item.Value.SEND_INTERVAL = SEND_INTERVAL
              _item.Value.ENABLE = ENABLE
              If dicUpdate_M_DataReport.ContainsKey(_item.Key) = False Then
                dicUpdate_M_DataReport.Add(_item.Key, _item.Value)
              Else
                Result_Message = "Select From dicUpdate_M_DataReport False"
                SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If
            Next
          End If
        End If
      Next
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef dicUpdate_M_DataReport As Dictionary(Of String, clsDATA_REPORT_SET),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      For Each obj In dicUpdate_M_DataReport.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          Result_Message = "Get update M_DataReport SQL Failed"
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
  Private Function Execute_DataUpdate(ByRef Result_Message As String, ByRef lstSql As List(Of String)) As Boolean
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
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
