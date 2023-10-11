'2019/1/16 上午 11:57:58
'V1.0.0
'Tool
'狀態:Checked
Imports eCA_HOSTObject
Imports eCA_TransactionMessage
Module Module_T3F5U4_ProductionCountSet
  Public Function O_Process_Message(ByRef Receive_Msg As MSG_T3F5U4_ProductionCountSet,
                                    ByRef Result_Message As String) As Boolean
    SyncLock gMain.objHandling.objLineProduction_InfoLock
      Try
        '異動的資料
        Dim lstCountModifyHist As New List(Of clsCOUNT_MODIFY_HIST)
        Dim dic_UpdateLineProductInfo As New Dictionary(Of String, clsLineProduction_Info)
        Dim lstSql As New List(Of String) '儲存要更新的SQL，進行一次性更新
        Dim lstQueueSql As New List(Of String) '儲存要更新的SQL，進行一次性更新
        '先進行資料邏輯檢查
        If Check_UpdateData(Receive_Msg, Result_Message) = False Then
          Return False
        End If
        '邏輯
        If Get_UpdateData(Receive_Msg, dic_UpdateLineProductInfo, lstCountModifyHist, Result_Message) = False Then
          Return False
        End If
        '取得SQL
        If Get_SQL(Result_Message, dic_UpdateLineProductInfo, lstCountModifyHist, lstSql, lstQueueSql) = False Then
          Return False
        End If
        '執行資料更新
        If Execute_DataUpdate(Result_Message, dic_UpdateLineProductInfo, lstSql, lstQueueSql) = False Then
          Return False
        End If
        Return True
      Catch ex As Exception
        Result_Message = ex.ToString
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End Try
    End SyncLock
  End Function
  Private Function Check_UpdateData(ByRef Receive_Msg As MSG_T3F5U4_ProductionCountSet, ByRef ret_strResultMsg As String) As Boolean
    Try
      'coding
      Dim strLog As String = ""
      For Each AreaInfo In Receive_Msg.Body.AreaList.AreaInfo
        Dim Factory_No As String = AreaInfo.FACTORY_NO
        Dim Area_No As String = AreaInfo.AREA_NO
        Dim Device_No As String = AreaInfo.DEVICE_NO
        Dim Unit_ID As String = AreaInfo.UNIT_ID
        Dim Qty_Modify As String = AreaInfo.QTY_MODIFY
        Dim Qty_NG As String = AreaInfo.QTY_NG
        If Factory_No = "" Then
          ret_strResultMsg = "FACTORY_NO is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If Area_No = "" Then
          ret_strResultMsg = "AREA_NO is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If Device_No = "" Then
          ret_strResultMsg = "DEVICE_NO is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If Unit_ID = "" Then
          ret_strResultMsg = "UNIT_ID is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查數量上報格式是否正確
        If IsNumeric(Qty_Modify) = False Then
          strLog = String.Format("Qty_Modify format error, Factory_No = <{0}>, Area_No = <{1}>, Qty_Modify = <{2}>;", Factory_No, Area_No, Qty_Modify)
          ret_strResultMsg += String.Format("數量格式異常, Factory_No = <{0}>, Area_No = <{1}>, Qty_Modify = <{2}>;", Factory_No, Area_No, Qty_Modify)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        If IsNumeric(Qty_NG) = False Then
          strLog = String.Format("Qty_NG format error, Factory_No = <{0}>, Area_No = <{1}>, Qty_NG = <{2}>;", Factory_No, Area_No, Qty_NG)
          ret_strResultMsg += String.Format("數量格式異常, Factory_No = <{0}>, Area_No = <{1}>, Qty_NG = <{2}>;", Factory_No, Area_No, Qty_NG)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        '檢查位置是否存在
        If gMain.objHandling.O_Get_Line_Area(Factory_No, Area_No) = False Then
          strLog = String.Format("Line Area is not Exist, Factory_No = <{0}>, Area_No = <{1}>;", Factory_No, Area_No)
          ret_strResultMsg = String.Format("查無此生產線位置, Factory_No = <{0}>, Area_No = <{1}>;", Factory_No, Area_No)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If gMain.objHandling.O_Get_CLine(Factory_No, Area_No, Device_No, Unit_ID) = False Then
          strLog = String.Format("CLine is not Exist, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          ret_strResultMsg = String.Format("查無此生產線位置, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Get_UpdateData(ByRef Receive_Msg As MSG_T3F5U4_ProductionCountSet,
                                  ByRef ret_dic_UpdateLineProductionInfo As Dictionary(Of String, clsLineProduction_Info),
                                  ByRef ret_lstCountModifyHist As List(Of clsCOUNT_MODIFY_HIST),
                                  ByRef ret_strResultMsg As String) As Boolean
    Try
      'logic
      Dim strLog As String = ""
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim User_ID As String = Receive_Msg.Header.ClientInfo.UserID
      For Each AreaInfo In Receive_Msg.Body.AreaList.AreaInfo
        Dim Factory_No As String = AreaInfo.FACTORY_NO
        Dim Area_No As String = AreaInfo.AREA_NO
        Dim Device_No As String = AreaInfo.DEVICE_NO
        Dim Unit_ID As String = AreaInfo.UNIT_ID
        Dim Qty_Modify As String = AreaInfo.QTY_MODIFY
        Dim Qty_NG As String = AreaInfo.QTY_NG
        Dim objLineProductionInfo As clsLineProduction_Info = Nothing
        If gMain.objHandling.O_Get_CLineProduction_Info(Factory_No, Area_No, Device_No, Unit_ID, objLineProductionInfo) = False Then
          strLog = String.Format("CLine is not Exist, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          ret_strResultMsg = String.Format("查無此生產線位置, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        Dim objNewLineProductionInfo As clsLineProduction_Info = objLineProductionInfo.Clone()
        objNewLineProductionInfo.Qty_Modify = objLineProductionInfo.Qty_Modify + Qty_Modify
        objNewLineProductionInfo.Qty_NG = objLineProductionInfo.Qty_Modify + Qty_NG
        If ret_dic_UpdateLineProductionInfo.ContainsKey(objNewLineProductionInfo.gid) = False Then
          ret_dic_UpdateLineProductionInfo.Add(objNewLineProductionInfo.gid, objNewLineProductionInfo)
        End If
        '寫入修改記錄
        Dim objCountModifyHist As New clsCOUNT_MODIFY_HIST(Factory_No, Area_No, Device_No, Unit_ID, User_ID, Qty_Modify, Qty_NG, Now_Time)
        ret_lstCountModifyHist.Add(objCountModifyHist)
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef ret_dic_UpdateLineProductionInfo As Dictionary(Of String, clsLineProduction_Info),
                           ByRef ret_lstCountModifyHist As List(Of clsCOUNT_MODIFY_HIST),
                           ByRef lstSql As List(Of String),
                           ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '取得要更新的SQL
      For Each obj In ret_dic_UpdateLineProductionInfo.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          Result_Message = "Get Update Line Product Info SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      '-
      For Each obj In ret_lstCountModifyHist
        If obj.O_Add_Insert_SQLString(lstQueueSql) = False Then
          Result_Message = "Get Insert Count Modify Hist SQL Failed"
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
                           ByRef ret_dic_UpdateLineProductionInfo As Dictionary(Of String, clsLineProduction_Info),
                           ByRef lstSql As List(Of String),
                           ByRef lstQueueSql As List(Of String)) As Boolean
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
      Common_DBManagement.AddQueued_BatchUpdate(lstQueueSql)
      '修改記憶體資料
      For Each objNew As clsLineProduction_Info In ret_dic_UpdateLineProductionInfo.Values
        '移除關聯
        Dim obj As clsLineProduction_Info = Nothing
        If gMain.objHandling.gdicLineProduction_Info.TryGetValue(objNew.gid, obj) = True Then
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
