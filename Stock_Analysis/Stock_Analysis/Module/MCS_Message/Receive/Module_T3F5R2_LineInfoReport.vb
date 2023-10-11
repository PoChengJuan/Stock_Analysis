Imports eCA_TransactionMessage
Imports eCA_HostObject
''' <summary>
''' 20181119
''' Mark
''' 狀態:Open(進行初步確認)
''' 設備上報設備保養訊息
''' </summary>
Module Module_T3F5R2_LineInfoReport
  Public Function O_Process_Message(ByVal Receive_Msg As MSG_T3F5R2_LineInfoReport,
                                    ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim dicAddLineInfo As New Dictionary(Of String, clsLineInfo)
      Dim dicDeleteLineInfo As New Dictionary(Of String, clsLineInfo)
      Dim lstSQL As New List(Of String)
      Dim lstQueueSQL As New List(Of String)
      '檢查資料是否正確
      If I_Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '取得要更新的資料
      If I_Get_Data(Receive_Msg, dicAddLineInfo, dicDeleteLineInfo, ret_strResultMsg) = False Then
        Return False
      End If
      '取得要更新的SQL
      If I_Get_SQL(ret_strResultMsg, dicAddLineInfo, dicDeleteLineInfo, lstSQL, lstQueueSQL) = False Then
        Return False
      End If
      '執行資料更新
      If I_Execute_DataUpdate(ret_strResultMsg, dicAddLineInfo, dicDeleteLineInfo, lstSQL, lstQueueSQL) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '檢查傳入的資料是否正確
  Private Function I_Check_Data(ByVal Receive_Msg As MSG_T3F5R2_LineInfoReport,
                                ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim strLog As String = ""
      For Each MessageInfo In Receive_Msg.Body.MessageList.MessageInfo
        Dim Factory_No As String = MessageInfo.FACTORY_NO
        Dim Area_No As String = MessageInfo.AREA_NO
        Dim Device_No As String = MessageInfo.DEVICE_NO
        Dim Unit_ID As String = MessageInfo.UNIT_ID
        Dim Time As String = MessageInfo.TIME
        Dim Message As String = MessageInfo.MESSAGE
        Dim Set_Flag As String = MessageInfo.SET_FLAG
        '檢查位置是否存在
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
  '取得資料庫新
  Private Function I_Get_Data(ByVal Receive_Msg As MSG_T3F5R2_LineInfoReport,
                              ByRef ret_dicAddLineInfo As Dictionary(Of String, clsLineInfo),
                              ByRef ret_dicDeleteLineInfo As Dictionary(Of String, clsLineInfo),
                              ByRef ret_strResultMsg As String) As Boolean
    Try
			'  Dim strLog As String = ""
			'  Dim Now_Time As String = GetNewTime_DBFormat()
			'  For Each MessageInfo In Receive_Msg.Body.MessageList.MessageInfo
			'    Dim Factory_No As String = MessageInfo.FACTORY_NO
			'    Dim Area_No As String = MessageInfo.AREA_NO
			'    Dim Device_No As String = MessageInfo.DEVICE_NO
			'    Dim Unit_ID As String = MessageInfo.UNIT_ID
			'    Dim Time As String = MessageInfo.TIME
			'Dim Message As String = MessageInfo.MESSAGE

			'Dim Set_Flag As Boolean = IntegerConvertToBoolean(MessageInfo.SET_FLAG)
			'    If Set_Flag = False Then  '清除原有的Message
			'      Dim objLineInfo As clsLineInfo = Nothing
			'      If gMain.objHandling.O_Get_CLineInfo(Factory_No, Area_No, Device_No, Unit_ID, Message, objLineInfo) = True Then
			'        If ret_dicDeleteLineInfo.ContainsKey(objLineInfo.gid) = False Then
			'          ret_dicDeleteLineInfo.Add(objLineInfo.gid, objLineInfo)
			'        End If
			'      Else
			'        strLog = String.Format("Clear Line Info, but Not Get Original Line Info, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Maintenance_Message = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Message)
			'        SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
			'      End If
			'    Else  '新增Message
			'      If gMain.objHandling.O_Get_CLineInfo(Factory_No, Area_No, Device_No, Unit_ID, Message) = False Then
			'        Dim objLineInfo As New clsLineInfo(Factory_No, Area_No, Device_No, Unit_ID, Now_Time, Message)
			'        If ret_dicAddLineInfo.ContainsKey(objLineInfo.gid) = False Then
			'          ret_dicAddLineInfo.Add(objLineInfo.gid, objLineInfo)
			'        End If
			'      Else
			'        strLog = String.Format("Set Line Info, but Not Line Info is exist, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Maintenance_Message = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Message)
			'        SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
			'      End If
			'    End If
			'  Next
			Return True
    Catch ex As Exception
      ret_strResultMsg = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得更新的SQL
  Private Function I_Get_SQL(ByRef ret_strResultMsg As String,
                             ByRef ret_dicAddLineInfo As Dictionary(Of String, clsLineInfo),
                             ByRef ret_dicDeleteLineInfo As Dictionary(Of String, clsLineInfo),
                             ByRef lstSQL As List(Of String),
                             ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '取得新增LineInfo的Insert SQL
      For Each obj In ret_dicAddLineInfo.Values
        If Not obj.O_Add_Insert_SQLString(lstSQL) Then
          ret_strResultMsg = "Get Insert WMS_CT_LINE_INFO SQL Failed"
          Return False
        End If
      Next
      '取得刪除LineInfo的Delete SQL
      For Each obj In ret_dicDeleteLineInfo.Values
        If Not obj.O_Add_Delete_SQLString(lstSQL) Then
          ret_strResultMsg = "Get Delete WMS_CT_LINE_INFO SQL Failed"
          Return False
        End If
      Next
      '取得要Insert进历史资料SQL
      Module_Process_History.Get_LineInfo_HIST_SQL(lstQueueSql, "WMS", ret_dicDeleteLineInfo)
      Return True
    Catch ex As Exception
      ret_strResultMsg = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行資料更新
  Private Function I_Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                        ByRef ret_dicAddLineInfo As Dictionary(Of String, clsLineInfo),
                                        ByRef ret_dicDeleteLineInfo As Dictionary(Of String, clsLineInfo),
                                        ByRef lstSQL As List(Of String),
                                        ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSQL.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
        '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
        ret_strResultMsg = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        '更新DB失敗則回傳False
        'ret_strResult_Message = "WMS Update DB Failed"
        ret_strResultMsg = "WMS 更新资料库失败"
        Return False
      End If
      Common_DBManagement.AddQueued(lstQueueSql)
      '修改記憶體資料
      '刪除LineInfo資訊
      For Each objNew As clsLineInfo In ret_dicDeleteLineInfo.Values
        Dim obj As clsLineInfo = Nothing
        If gMain.objHandling.gdicLineInfo.TryGetValue(objNew.gid, obj) Then
          obj.Remove_Relationship()
        End If
      Next
      '新增LineInfo資訊
      For Each obj As clsLineInfo In ret_dicAddLineInfo.Values
        obj.Add_Relationship(gMain.objHandling)
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
