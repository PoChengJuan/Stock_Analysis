Imports eCA_TransactionMessage
Imports eCA_HostObject

''' <summary>
''' 20181117
''' V1.0.0
''' Mark
''' 狀態:Open(進行初步確認)
''' 設備上報生產線運行狀態
''' </summary>
Module Module_T3F5R1_LineStatusChangeReport
  Public Function O_Process_Message(ByVal Receive_Msg As MSG_T3F5R1_LineStatusChangeReport,
                                    ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim dicUpdateCLine As New Dictionary(Of String, clsLine_Status)
      Dim lstSQL As New List(Of String)
      Dim lstQueueSQL As New List(Of String)
      '檢查資料是否正確
      If I_Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '取得要更新的資料
      If I_Get_Data(Receive_Msg, dicUpdateCLine, ret_strResultMsg) = False Then
        Return False
      End If
      '取得要更新的SQL
      If I_Get_SQL(ret_strResultMsg, dicUpdateCLine, lstSQL, lstQueueSQL) = False Then
        Return False
      End If
      '執行資料更新
      If I_Execute_DataUpdate(ret_strResultMsg, dicUpdateCLine, lstSQL, lstQueueSQL) = False Then
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
  Private Function I_Check_Data(ByVal Receive_Msg As MSG_T3F5R1_LineStatusChangeReport,
                                ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim strLog As String = ""
      Dim Factory_No = Receive_Msg.Body.DeviceInfo.FACTORY_NO
      Dim Device_No = Receive_Msg.Body.DeviceInfo.DEVICE_NO
      For Each LineInfo In Receive_Msg.Body.LineList.LineInfo
        Dim Area_No As String = LineInfo.AREA_NO
        Dim Unit_ID As String = LineInfo.UNIT_ID
        Dim Status As String = LineInfo.STATUS
        '檢查Unit_ID是否存在
        If gMain.objHandling.O_Get_CLine(Factory_No, Area_No, Device_No, Unit_ID) = False Then
          strLog = String.Format("CLine is not Exist, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          ret_strResultMsg = String.Format("查無此生產線位置, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查傳入的狀態是否正確
        If CheckValueInEnum(Of enuLineStatus)(Status) = False Then
          strLog = String.Format("Line Status not defined, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Status = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Status)
          ret_strResultMsg = String.Format("此生產線狀態未定義, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Status = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Status)
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
  Private Function I_Get_Data(ByVal Receive_Msg As MSG_T3F5R1_LineStatusChangeReport,
                              ByRef ret_dicUpdateLine As Dictionary(Of String, clsLine_Status),
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim strLog As String = ""
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim Factory_No = Receive_Msg.Body.DeviceInfo.FACTORY_NO
      Dim Device_No = Receive_Msg.Body.DeviceInfo.DEVICE_NO
      For Each LineInfo In Receive_Msg.Body.LineList.LineInfo
        Dim Area_No As String = LineInfo.AREA_NO
        Dim Unit_ID As String = LineInfo.UNIT_ID
        Dim Status As enuLineStatus = CInt(LineInfo.STATUS)
        Dim objLine As clsLine_Status = Nothing
        '檢查Unit_ID是否存在
        If gMain.objHandling.O_Get_CLine(Factory_No, Area_No, Device_No, Unit_ID, objLine) = True Then
          Dim objNew_Line As clsLine_Status = Nothing
          If ret_dicUpdateLine.TryGetValue(objLine.gid, objNew_Line) = False Then
            objNew_Line = objLine.Clone()
            ret_dicUpdateLine.Add(objNew_Line.gid, objNew_Line)
          End If
          objNew_Line.Status = Status
          objNew_Line.Update_Time = Now_Time
        Else
          strLog = String.Format("CLine is not Exist, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          ret_strResultMsg = String.Format("查無此生產線位置, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得更新的SQL
  Private Function I_Get_SQL(ByRef ret_strResultMsg As String,
                             ByRef ret_dicUpdateLine As Dictionary(Of String, clsLine_Status),
                             ByRef lstSQL As List(Of String),
                             ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '取得新增Carrier的Update SQL
      For Each obj In ret_dicUpdateLine.Values
        If Not obj.O_Add_Update_SQLString(lstSQL) Then
          ret_strResultMsg = "Get Update WMS_CT_LINE_STATUS SQL Failed"
          Return False
        End If
      Next
      '取得要Insert进历史资料SQL
      Module_Process_History.Get_LineStatus_HIST_SQL(lstQueueSql, ret_dicUpdateLine)
      Return True
    Catch ex As Exception
      ret_strResultMsg = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行資料更新
  Private Function I_Execute_DataUpdate(ByRef ret_strResult_Message As String,
                                        ByRef ret_dicUpdateLine As Dictionary(Of String, clsLine_Status),
                                        ByRef lstSQL As List(Of String),
                                        ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSQL.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
        '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
        ret_strResult_Message = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
        SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        '更新DB失敗則回傳False
        'ret_strResult_Message = "WMS Update DB Failed"
        ret_strResult_Message = "HostHandler 更新资料库失败"
        Return False
      End If
      Common_DBManagement.AddQueued(lstQueueSql)
      '修改記憶體資料
      '更新Line資訊
      For Each objNew As clsLine_Status In ret_dicUpdateLine.Values
        Dim obj As clsLine_Status = Nothing
        If gMain.objHandling.gdicLine.TryGetValue(objNew.gid, obj) Then
          obj.Update_To_Memory(objNew)
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResult_Message = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
