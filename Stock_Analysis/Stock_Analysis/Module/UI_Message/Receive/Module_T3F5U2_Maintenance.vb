'2019/1/16 上午 11:56:57
'V1.0.0
'Tool
'狀態:Checked
Imports eCA_HostObject
Imports eCA_TransactionMessage
Module Module_T3F5U2_Maintenance

  Public Function O_Process_Message(ByRef Receive_Msg As MSG_T3F5U2_Maintenance, ByRef Result_Message As String) As Boolean
    Try
      '異動的資料			
      Dim lstAddLineHist As New List(Of clsLineInfo_Hist)
      Dim dicDeleteLineInfo As New Dictionary(Of String, clsLineInfo)
      Dim dicUpdateMaintenanceStatus As New Dictionary(Of String, clsMAINTENANCE_STATUS)

      Dim lstSql As New List(Of String) '儲存要更新的SQL，進行一次性更新
      Dim lstQueuedSql As New List(Of String) '儲存要更新的SQL，進行一次性更新

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, Result_Message) = False Then
        Return False
      End If
      '邏輯
      If Get_UpdateData(Receive_Msg, Result_Message, dicDeleteLineInfo, dicUpdateMaintenanceStatus, lstAddLineHist) = False Then
        Return False
      End If
      '取得SQL
      If Get_SQL(Result_Message, dicDeleteLineInfo, dicUpdateMaintenanceStatus, lstAddLineHist, lstSql, lstQueuedSql) = False Then
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(Result_Message, dicDeleteLineInfo, dicUpdateMaintenanceStatus, lstSql, lstQueuedSql) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Check_Data(ByRef Receive_Msg As MSG_T3F5U2_Maintenance, ByRef Result_Message As String) As Boolean
    Try
      For Each UnitInfo In Receive_Msg.Body.UnitList.UnitInfo
        If UnitInfo.FACTORY_NO = "" Then
          Result_Message = "FACTORY_NO is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If UnitInfo.DEVICE_NO = "" Then
          Result_Message = "DEVICE_NO is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If UnitInfo.AREA_NO = "" Then
          Result_Message = "AREA_NO is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If UnitInfo.UNIT_ID = "" Then
          Result_Message = "UNIT_ID is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If UnitInfo.MAINTENANCE_ID = "" Then
          Result_Message = "MAINTENANCE_ID is empty"
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
  Private Function Get_UpdateData(ByRef Receive_Msg As MSG_T3F5U2_Maintenance,
                                  ByRef Result_Message As String,
                                  ByRef ret_dicDeleteLineInfo As Dictionary(Of String, clsLineInfo),
                                  ByRef ret_dicUpdateMaintenanceStatus As Dictionary(Of String, clsMAINTENANCE_STATUS),
                                  ByRef ret_lstAddLineHist As List(Of clsLineInfo_Hist)) As Boolean
    Try
      '取得
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim User_ID As String = Receive_Msg.Header.ClientInfo.UserID
      For Each UnitInfo In Receive_Msg.Body.UnitList.UnitInfo
        Dim Factory_No As String = UnitInfo.FACTORY_NO
        Dim Device_No As String = UnitInfo.DEVICE_NO
        Dim Area_No As String = UnitInfo.AREA_NO
        Dim Unit_ID As String = UnitInfo.UNIT_ID
        Dim Maintenance_ID As String = UnitInfo.MAINTENANCE_ID
        Dim OperatorUser As String = UnitInfo.OPERATOR_USER
        Dim Comments As String = UnitInfo.COMMENTS
        '取得該Maintenance_ID對應的MaintenanceStatus
        Dim tmp_MaintenanceStatus As New Dictionary(Of String, clsMAINTENANCE_STATUS)
        '取得該Maintenance對應的資料
        If gMain.objHandling.O_Get_dicMantenaceStatusByFactoryNo_DeviceNo_AreaNo_UnitID_MaintenanceID_FunctionID(Factory_No, Device_No, Area_No, Unit_ID, Maintenance_ID, "", tmp_MaintenanceStatus) = False Then
          Result_Message = String.Format("查無對應的保養資訊, Factory_No={0}, Area_No={1}, Device_No={2}, Unit_ID={3}, Maintenance_ID={4}", Factory_No, Area_No, Device_No, Unit_ID, Maintenance_ID)
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        For Each objMaintenanceStatus As clsMAINTENANCE_STATUS In tmp_MaintenanceStatus.Values
          '檢查是否有資料在LineInfo中，如果有要移除LineInfo並寫入History
          Dim objLineInfo As clsLineInfo = Nothing
          If gMain.objHandling.O_Get_CLineInfo(Factory_No, Area_No, Device_No, Unit_ID, Maintenance_ID, objMaintenanceStatus.FUNCTION_ID, objLineInfo) = True Then
            If ret_dicDeleteLineInfo.ContainsKey(objLineInfo.gid) = False Then
              ret_dicDeleteLineInfo.Add(objLineInfo.gid, objLineInfo)
            End If
            '加入到History
            Dim objLineHistory As New clsLineInfo_Hist(Factory_No, Area_No, Device_No, Unit_ID, objLineInfo.Occur_Time, objLineInfo.Maintenance_Message, User_ID, Now_Time, Maintenance_ID, objMaintenanceStatus.FUNCTION_ID, OperatorUser, Comments)
            ret_lstAddLineHist.Add(objLineHistory)
          Else
            '加入到History
            Dim objLineHistory As New clsLineInfo_Hist(Factory_No, Area_No, Device_No, Unit_ID, "", "", User_ID, Now_Time, Maintenance_ID, objMaintenanceStatus.FUNCTION_ID, OperatorUser, Comments)
            ret_lstAddLineHist.Add(objLineHistory)
          End If
          '更新MaintenanceStatus
          Dim objNewMaintenanceStatus As clsMAINTENANCE_STATUS = objMaintenanceStatus.Clone()
          Dim objMaintenanceDTL As clsMAINTENANCE_DTL = Nothing
          If gMain.objHandling.O_Get_MaintenanceDTL(Factory_No, Device_No, Area_No, Unit_ID, Maintenance_ID, objNewMaintenanceStatus.FUNCTION_ID, objMaintenanceDTL) = True Then
            If objMaintenanceDTL.VALUE_TYPE = enuMaintenanceValueType.IsDate Then '如果是日期就把MaintenanceStatus改成現在日期
              objNewMaintenanceStatus.VALUE = Now_Time
              objNewMaintenanceStatus.UPDATE_TIME = Now_Time
              objNewMaintenanceStatus.MAINTENANCE_SET = False
            Else
              objNewMaintenanceStatus.VALUE = 0
              objNewMaintenanceStatus.UPDATE_TIME = Now_Time
              objNewMaintenanceStatus.MAINTENANCE_SET = False
            End If
          Else
            objNewMaintenanceStatus.VALUE = 0
            objNewMaintenanceStatus.UPDATE_TIME = Now_Time
            objNewMaintenanceStatus.MAINTENANCE_SET = False
          End If
          If ret_dicUpdateMaintenanceStatus.ContainsKey(objNewMaintenanceStatus.gid) = False Then
            ret_dicUpdateMaintenanceStatus.Add(objNewMaintenanceStatus.gid, objNewMaintenanceStatus)
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
                           ByRef ret_dicDeleteLineInfo As Dictionary(Of String, clsLineInfo),
                           ByRef ret_dicUpdateMaintenanceStatus As Dictionary(Of String, clsMAINTENANCE_STATUS),
                           ByRef ret_lstAddLineHist As List(Of clsLineInfo_Hist),
                           ByRef lstSql As List(Of String),
                           ByRef lstQueuedSql As List(Of String)) As Boolean
    Try
      For Each obj In ret_dicDeleteLineInfo.Values
        If obj.O_Add_Delete_SQLString(lstSql) = False Then
          Result_Message = "Get Delete WMS_CT_Line_Info SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj In ret_dicUpdateMaintenanceStatus.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          Result_Message = "Get update WMS_T_MAINTENANCE_STATUS SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj In ret_lstAddLineHist
        If obj.O_Add_Insert_SQLString(lstQueuedSql) = False Then
          Result_Message = "Get Insert WMS_CH_Line_HIST SQL Failed"
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
                                      ByRef ret_dicDeleteLineInfo As Dictionary(Of String, clsLineInfo),
                                      ByRef ret_dicUpdateMaintenanceStatus As Dictionary(Of String, clsMAINTENANCE_STATUS),
                                      ByRef lstSql As List(Of String),
                                      ByRef lstQueuedSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSql.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
        '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
        Result_Message = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        'Result_Message = "SQL Insert DB Failed"
        Result_Message = "WMS 更新资料库失败"
        Return False
      End If

      '執行Queue
      Common_DBManagement.AddQueued_BatchUpdate(lstQueuedSql)
      '修改記憶體資料
      For Each objNew As clsLineInfo In ret_dicDeleteLineInfo.Values
        '移除關聯
        Dim obj As clsLineInfo = Nothing
        If gMain.objHandling.gdicLineInfo.TryGetValue(objNew.gid, obj) = True Then
          obj.Remove_Relationship()
        End If
      Next
      For Each objNew As clsMAINTENANCE_STATUS In ret_dicUpdateMaintenanceStatus.Values
        '移除關聯
        Dim obj As clsMAINTENANCE_STATUS = Nothing
        If gMain.objHandling.gdicMaintenance_Status.TryGetValue(objNew.gid, obj) = True Then
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
