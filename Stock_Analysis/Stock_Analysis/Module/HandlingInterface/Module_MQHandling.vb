Imports System.Collections.Concurrent

Imports eCA_HostObject
Imports eCAMQTool.eCARabbitMQ

Module Module_MQHandling
#Region "WMS的部分"
  ''' <summary>
  ''' 處理使用MQ交握WMS_Commad的部份(分3部份中的第1部份)
  ''' </summary>
  Public Sub O_thr_WMSMQHandling()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        '顯示在FORM上的CLOCK
        'gMain.date_tWMS_CommandDBHandle = Now
        'If WMS_StopGetWMSCommand Then
        '  SendMessageToLog("Process WMS Command3 Stop...", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  While WMS_StopGetWMSCommand
        '    Threading.Thread.Sleep(2000)
        '  End While
        'End If
        'If Count < 10 Then
        '  Count += 1
        'Else
        '  Count = 0
        '  If gMain.int_tWMSCommandDBHandle > 99 Then
        '    gMain.int_tWMSCommandDBHandle = 0
        '  Else
        '    gMain.int_tWMSCommandDBHandle += 1
        '  End If
        'End If
        I_ProcessFromWMSMessage()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_WMS_CommandMQHandling3 End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Private Function I_ProcessFromWMSMessage() As Boolean
    Try
      While gMain.dicWMS_TO_HANDLING_Queue.Any = True
        Dim objQueue = gMain.dicWMS_TO_HANDLING_Queue.First.Value
        '取得headers '這四個固定要有的  在監控MQ THREAD已經先CHECK過DATA才放到各系統的待處理THREAD中，因此DECODE會成功
        Dim FUNCTION_ID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        Dim UUID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
        Dim SYSTEM_TYPE As enuSystemType = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.SEND_SYSTEM.ToString))
        Dim UserID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.USER_ID.ToString))
        Dim ClientID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.CLIENT_ID.ToString))
        Dim ip = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.IP.ToString))
        Dim XmlMessage As String = clsRabbitMQ.O_Decode_Message(objQueue.Body)
        Dim Result As String = ""
        Dim ResultMessage As String = ""
        Dim Report As String = ""
        Dim Wait_UUID As String = ""
        Dim strLog As String = ""

        '如果為等待回覆的 則跳過 
        If gMain.objHandling.O_Get_Command_Report(SYSTEM_TYPE, UUID) = True Then
          '把事件加入Report中
          If gMain.dicWMS_TO_HANDLING_Queue_R.ContainsKey(objQueue.DeliveryTag) = False Then
            gMain.dicWMS_TO_HANDLING_Queue_R.Add(objQueue.DeliveryTag, objQueue)
          End If
          gMain.dicWMS_TO_HANDLING_Queue.Remove(objQueue.DeliveryTag)
          Continue While
        End If

        '把取得的WMSCommand送出去進行處理
        If O_ProcessCommand(FUNCTION_ID, XmlMessage, ResultMessage, enuHTTPContentType.XML, Wait_UUID) = True Then
          '如果Wait_UUID不為空時才把Wait_UUID填入
          If Wait_UUID = "" Then
            Result = "0"
            'ResultMessage = ""
          End If
        Else
          '執行失敗
          Result = "1"
        End If

        '根據W有沒有ait_UUID進行對應的處理
        If Wait_UUID = "" Then
          '回覆處理結果
          If gMain.RabbitMQ.ResultSecondaryMessage(Result, ResultMessage, enuRabbitMQ.HOST_TO_WMS.ToString, "", "", objQueue, strLog) = False Then
            SendMessageToLog($"ResultSecondaryMessage Falied : {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If

          'Ack訊息 回覆MQ這筆已處理完成的訊息可以結束掉
          If gMain.RabbitMQ.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, strLog) = False Then
            SendMessageToLog($"Ack_ReceiveMessage Falied : {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If

          Dim Headers = objQueue.BasicProperties.Headers
          I_Write_HOST_T_COMMAND_HIS(XmlMessage, Headers, enuSystemType.GUI, Result, ResultMessage, Wait_UUID)
        Else '如果有Wait UUID
          Dim lstSQL As New List(Of String)
          Dim dicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)
          Dim objHOST_T_COMMAND_REPORT As New clsHOST_T_COMMAND_REPORT(Wait_UUID, SYSTEM_TYPE, UUID, GetNewTime_DBFormat)
          If dicHOST_T_COMMAND_REPORT.ContainsKey(objHOST_T_COMMAND_REPORT.gid) = False Then
            dicHOST_T_COMMAND_REPORT.Add(objHOST_T_COMMAND_REPORT.gid, objHOST_T_COMMAND_REPORT)
          End If
          For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
            If objHOST_T_COMMAND_REPORT.O_Add_Insert_SQLString(lstSQL) = False Then
              SendMessageToLog("HOST_T_COMMAND_REPORT O_Add_Insert_SQLString Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
            End If
          Next
          objHOST_T_COMMAND_REPORT.O_Add_Insert_SQLString(lstSQL)
          '寫入DB
          If Common_DBManagement.BatchUpdate(lstSQL) = False Then
            Return False
          End If

          '加入記憶體
          For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
            gMain.objHandling.O_Add_HOST_T_COMMAND_REPORT(objHOST_T_COMMAND_REPORT)
          Next

          '把事件加入Report中，等回覆
          If gMain.dicWMS_TO_HANDLING_Queue_R.ContainsKey(objQueue.DeliveryTag) = False Then
            gMain.dicWMS_TO_HANDLING_Queue_R.Add(objQueue.DeliveryTag, objQueue)
          End If
        End If

        SendMessageToLog("Set WMS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & Result & ", Result_Message=" & ResultMessage & ", Wait_UUID=" & Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        gMain.dicWMS_TO_HANDLING_Queue.Remove(objQueue.DeliveryTag)
      End While

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Sub O_thr_ToWMSMQHandling_Result()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        If Count < 10 Then
          Count += 1
        Else
          Count = 0
          'If gMain.int_tWMSToWMSResultDBHandle > 99 Then
          '  gMain.int_tWMSToWMSResultDBHandle = 0
          'Else
          '  gMain.int_tWMSToWMSResultDBHandle += 1
          'End If
        End If
        I_ProcessToWMSMessageResult()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_ToWMSMQHandling_Result End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Private Function I_ProcessToWMSMessageResult() As Boolean
    Try
      While gMain.dicWMS_TO_HANDLING_Queue_S.Any = True
        Dim objQueue = gMain.dicWMS_TO_HANDLING_Queue_S.First.Value
        '取得headers '這四個固定要有的
        Dim FUNCTION_ID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        Dim UUID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
        Dim XmlMessageResult As String = clsRabbitMQ.O_Decode_Message(objQueue.Body)
        Dim Wait_UUID As String = ""
        Dim str_ResultMsg As String = ""
        Dim ReplyResult As Boolean = IIf(clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.RESULT.ToString)) = "0", True, False)
        Dim RejectReason As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.RESULT_MESSAGE.ToString))

        Dim XmlMessage As String = ""
        Dim dicProcessFromHostCommand As New Dictionary(Of String, clsFromHostCommand)

        Dim dicFromHostCommandResult = HOST_T_CommandManagement.GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(enuSystemType.WMS)
        For Each objFromHostCommandResult In dicFromHostCommandResult.Values
          '不相等表示這一筆不是當前要處理的命令， 下一次再處理
          If UUID <> objFromHostCommandResult.UUID OrElse FUNCTION_ID <> objFromHostCommandResult.Function_ID Then
            Continue For
          End If
          If dicProcessFromHostCommand.ContainsKey(objFromHostCommandResult.gid) = False Then
            dicProcessFromHostCommand.Add(objFromHostCommandResult.gid, objFromHostCommandResult)
          Else
            SendMessageToLog(String.Format("WMSCommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objFromHostCommandResult.UUID, objFromHostCommandResult.Function_ID, objFromHostCommandResult.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
          XmlMessage = XmlMessage & objFromHostCommandResult.Message
        Next

        '記錄發生給WMS後回覆的給結果
        If ReplyResult = True Then
          SendMessageToLog("Get Send To WMS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog("Get Send To WMS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        '把取得的ToWMSCommand送出去進行處理
        If O_ProcessCommandResult(UUID, FUNCTION_ID, XmlMessage, RejectReason, ReplyResult, str_ResultMsg) = True Then
          If ProcessHostCommandResult_Move_to_Hist(dicProcessFromHostCommand) = True Then

          End If
        End If
        'Ack訊息
        If gMain.RabbitMQ.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, str_ResultMsg) = False Then
          SendMessageToLog($"Ack_ReceiveMessage Falied : {str_ResultMsg}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        gMain.dicWMS_TO_HANDLING_Queue_S.Remove(objQueue.DeliveryTag)
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Set_WMS_MQWaitCommandResult(ByRef objHOST_T_COMMAND_REPORT As clsHOST_T_COMMAND_REPORT, ByVal WAIT_UUID As String, ByVal bln_Result As Boolean, ByVal str_RejectReason As String) As Boolean
    Try
      Dim Result As String = IIf(bln_Result = True, 0, 1)
      Dim strLog As String = ""
      Dim lstSQL As New List(Of String)
      '取得MQ_R資訊
      Dim objQueue As RabbitMQ.Client.Events.BasicDeliverEventArgs = Nothing
      If O_Get_objMQEventByUUID(gMain.dicWMS_TO_HANDLING_Queue_R, objHOST_T_COMMAND_REPORT.REPORT_SYSTEM_UUID, objQueue) = True Then
        '發Secondary
        Dim FUNCTION_ID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        If gMain.RabbitMQ.ResultSecondaryMessage(Result, str_RejectReason, enuRabbitMQ.HOST_TO_WMS.ToString, "", "", objQueue, strLog, WAIT_UUID:=WAIT_UUID) = False Then
          SendMessageToLog($"ResultSecondaryMessage Failed {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        'HOST_H_COMMAND_HIST 歷史紀錄
        Dim Headers = objQueue.BasicProperties.Headers
        I_Write_HOST_T_COMMAND_HIS(strLog, Headers, strLog, WAIT_UUID, Result, str_RejectReason)

        '刪除HOST_T_COMMAND_REPORT的DB紀錄
        Dim dicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)
        If dicHOST_T_COMMAND_REPORT.ContainsKey(objHOST_T_COMMAND_REPORT.gid) = False Then
          dicHOST_T_COMMAND_REPORT.Add(objHOST_T_COMMAND_REPORT.gid, objHOST_T_COMMAND_REPORT)
        End If
        For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
          If objHOST_T_COMMAND_REPORT.O_Add_Delete_SQLString(lstSQL) = False Then
            SendMessageToLog("HOST_T_COMMAND_REPORT O_Add_Delete_SQLString Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
          End If
        Next
        If Common_DBManagement.BatchUpdate(lstSQL) = False Then
          SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
        '刪除HOST_T_COMMAND_REPORT的GDIC紀錄
        For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
          gMain.objHandling.O_Remove_HOST_T_COMMAND_REPORT(objHOST_T_COMMAND_REPORT)
        Next
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
#End Region

#Region "MCS的部分"
  ''' <summary>
  ''' 處理使用MQ交握MCS_Commad的部份(分3部份中的第1部份)
  ''' </summary>
  Public Sub O_thr_MCSMQHandling()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        '顯示在FORM上的CLOCK
        'gMain.date_tMCS_CommandDBHandle = Now
        'If WMS_StopGetMCSCommand Then
        '  SendMessageToLog("Process MCS Command3 Stop...", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  While WMS_StopGetMCSCommand
        '    Threading.Thread.Sleep(2000)
        '  End While
        'End If
        'If Count < 10 Then
        '  Count += 1
        'Else
        '  Count = 0
        '  If gMain.int_tMCSCommandDBHandle > 99 Then
        '    gMain.int_tMCSCommandDBHandle = 0
        '  Else
        '    gMain.int_tMCSCommandDBHandle += 1
        '  End If
        'End If
        I_ProcessFromMCSMessage()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_MCS_CommandMQHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Private Function I_ProcessFromMCSMessage() As Boolean
    Try
      While gMain.dicMCS_TO_HANDLING_Queue.Any = True
        Dim objQueue = gMain.dicMCS_TO_HANDLING_Queue.First.Value
        '取得headers '這四個固定要有的  在監控MQ THREAD已經先CHECK過DATA才放到各系統的待處理THREAD中，因此DECODE會成功
        Dim FUNCTION_ID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        Dim UUID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
        Dim SYSTEM_TYPE As enuSystemType = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.SEND_SYSTEM.ToString))
        Dim UserID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.USER_ID.ToString))
        Dim ClientID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.CLIENT_ID.ToString))
        Dim ip = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.IP.ToString))
        Dim XmlMessage As String = clsRabbitMQ.O_Decode_Message(objQueue.Body)
        Dim Result As String = ""
        Dim ResultMessage As String = ""
        Dim Report As String = ""
        Dim Wait_UUID As String = ""
        Dim strLog As String = ""

        '如果為等待回覆的 則跳過 
        If gMain.objHandling.O_Get_Command_Report(SYSTEM_TYPE, UUID) = True Then
          '把事件加入Report中
          If gMain.dicMCS_TO_HANDLING_Queue_R.ContainsKey(objQueue.DeliveryTag) = False Then
            gMain.dicMCS_TO_HANDLING_Queue_R.Add(objQueue.DeliveryTag, objQueue)
          End If
          gMain.dicMCS_TO_HANDLING_Queue.Remove(objQueue.DeliveryTag)
          Continue While
        End If

        '把取得的MCSCommand送出去進行處理
        If O_ProcessCommand(FUNCTION_ID, XmlMessage, ResultMessage, Wait_UUID) = True Then
          '如果Wait_UUID不為空時才把Wait_UUID填入
          If Wait_UUID = "" Then
            Result = "0"
            'ResultMessage = ""
          End If
        Else
          '執行失敗
          Result = "1"
        End If

        '根據Wait_UUID進行對應的處理
        If Wait_UUID = "" Then
          '回覆處理結果
          If gMain.RabbitMQ.ResultSecondaryMessage(Result, ResultMessage, enuRabbitMQ.HOST_TO_MCS.ToString, "", "", objQueue, strLog) = False Then
            SendMessageToLog($"ResultSecondaryMessage Falied : {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If

          'Ack訊息 回覆MQ這筆已處理完成的訊息可以結束掉
          If gMain.RabbitMQ.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, strLog) = False Then
            SendMessageToLog($"Ack_ReceiveMessage Falied : {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If

          Dim Headers = objQueue.BasicProperties.Headers
          I_Write_HOST_T_COMMAND_HIS(ResultMessage, Headers, enuSystemType.GUI, Result, ResultMessage, Wait_UUID)
        Else '如果有Wait UUID
          Dim lstSQL As New List(Of String)
          Dim dicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)
          Dim objHOST_T_COMMAND_REPORT As New clsHOST_T_COMMAND_REPORT(Wait_UUID, SYSTEM_TYPE, UUID, GetNewTime_DBFormat)
          If dicHOST_T_COMMAND_REPORT.ContainsKey(objHOST_T_COMMAND_REPORT.gid) = False Then
            dicHOST_T_COMMAND_REPORT.Add(objHOST_T_COMMAND_REPORT.gid, objHOST_T_COMMAND_REPORT)
          End If
          For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
            If objHOST_T_COMMAND_REPORT.O_Add_Insert_SQLString(lstSQL) = False Then
              SendMessageToLog("HOST_T_COMMAND_REPORT O_Add_Insert_SQLString Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
            End If
          Next
          objHOST_T_COMMAND_REPORT.O_Add_Insert_SQLString(lstSQL)
          '寫入DB
          If Common_DBManagement.BatchUpdate(lstSQL) = False Then
            Return False
          End If

          '加入記憶體
          For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
            gMain.objHandling.O_Add_HOST_T_COMMAND_REPORT(objHOST_T_COMMAND_REPORT)
          Next

          '把事件加入Report中
          If gMain.dicMCS_TO_HANDLING_Queue_R.ContainsKey(objQueue.DeliveryTag) = False Then
            gMain.dicMCS_TO_HANDLING_Queue_R.Add(objQueue.DeliveryTag, objQueue)
          End If
        End If

        SendMessageToLog("Set MCS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & Result & ", Result_Message=" & ResultMessage & ", Wait_UUID=" & Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        gMain.dicMCS_TO_HANDLING_Queue.Remove(objQueue.DeliveryTag)
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Sub O_thr_ToMCSMQHandling_Result()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        If Count < 10 Then
          Count += 1
        Else
          Count = 0
          'If gMain.int_tWMSToMCSResultDBHandle > 99 Then
          '  gMain.int_tWMSToMCSResultDBHandle = 0
          'Else
          '  gMain.int_tWMSToMCSResultDBHandle += 1
          'End If
        End If
        I_ProcessToMCSMessageResult()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_WMSToMCSResultDBHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Private Function I_ProcessToMCSMessageResult() As Boolean
    Try
      While gMain.dicMCS_TO_HANDLING_Queue_S.Any = True
        Dim objQueue = gMain.dicMCS_TO_HANDLING_Queue_S.First.Value
        '取得headers '這四個固定要有的
        Dim FUNCTION_ID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        Dim UUID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
        Dim XmlMessageResult As String = clsRabbitMQ.O_Decode_Message(objQueue.Body)
        Dim Wait_UUID As String = ""
        Dim str_ResultMsg As String = ""
        Dim ReplyResult As Boolean = IIf(clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.RESULT.ToString)) = "0", True, False)
        Dim RejectReason As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.RESULT_MESSAGE.ToString))

        'MQ收完訊息，再到HOST_T_Command取RECEIVE = MCS的資料
        Dim XmlMessage As String = ""
        Dim dicProcessFromHostCommand As New Dictionary(Of String, clsFromHostCommand)

        Dim dicFromHostCommandResult = HOST_T_CommandManagement.GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(enuSystemType.MCS)
        For Each objFromHostCommandResult In dicFromHostCommandResult.Values
          '不相等表示這一筆不是當前要處理的命令，下一次再處理
          If UUID <> objFromHostCommandResult.UUID OrElse FUNCTION_ID <> objFromHostCommandResult.Function_ID Then
            Continue For
          End If
          If dicProcessFromHostCommand.ContainsKey(objFromHostCommandResult.gid) = False Then
            dicProcessFromHostCommand.Add(objFromHostCommandResult.gid, objFromHostCommandResult)
          Else
            SendMessageToLog(String.Format("MCSCommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objFromHostCommandResult.UUID, objFromHostCommandResult.Function_ID, objFromHostCommandResult.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
          XmlMessage = XmlMessage & objFromHostCommandResult.Message
        Next

        '記錄發生給MCS後回覆的給結果
        If ReplyResult = True Then
          SendMessageToLog("Get Send To MCS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog("Get Send To MCS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        '把取得的ToMCSCommand送出去進行處理
        If O_ProcessCommandResult(UUID, FUNCTION_ID, XmlMessage, RejectReason, ReplyResult, str_ResultMsg) = True Then
          If ProcessHostCommandResult_Move_to_Hist(dicProcessFromHostCommand) = True Then

          End If
        End If
        'Ack訊息
        If gMain.RabbitMQ.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, str_ResultMsg) = False Then
          SendMessageToLog($"Ack_ReceiveMessage Falied : {str_ResultMsg}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        gMain.dicMCS_TO_HANDLING_Queue_S.Remove(objQueue.DeliveryTag)
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Set_MCS_MQWaitCommandResult(ByRef objHOST_T_COMMAND_REPORT As clsHOST_T_COMMAND_REPORT, ByVal WAIT_UUID As String, ByVal bln_Result As Boolean, ByVal str_RejectReason As String) As Boolean
    Try
      Dim Result As String = IIf(bln_Result = True, 0, 1)
      Dim strLog As String = ""
      Dim lstSQL As New List(Of String)
      '取得MQ_R資訊
      Dim objQueue As RabbitMQ.Client.Events.BasicDeliverEventArgs = Nothing
      If O_Get_objMQEventByUUID(gMain.dicMCS_TO_HANDLING_Queue_R, objHOST_T_COMMAND_REPORT.REPORT_SYSTEM_UUID, objQueue) = True Then
        '發Secondary
        Dim FUNCTION_ID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        If gMain.RabbitMQ.ResultSecondaryMessage(Result, str_RejectReason, enuRabbitMQ.HOST_TO_MCS.ToString, "", "", objQueue, strLog, WAIT_UUID:=WAIT_UUID) = False Then
          SendMessageToLog($"ResultSecondaryMessage Failed {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        'HOST_H_COMMAND_HIST 歷史紀錄
        Dim Headers = objQueue.BasicProperties.Headers
        I_Write_HOST_T_COMMAND_HIS(strLog, Headers, strLog, WAIT_UUID, Result, str_RejectReason)

        '刪除HOST_T_COMMAND_REPORT的DB紀錄
        Dim dicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)
        If dicHOST_T_COMMAND_REPORT.ContainsKey(objHOST_T_COMMAND_REPORT.gid) = False Then
          dicHOST_T_COMMAND_REPORT.Add(objHOST_T_COMMAND_REPORT.gid, objHOST_T_COMMAND_REPORT)
        End If
        For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
          If objHOST_T_COMMAND_REPORT.O_Add_Delete_SQLString(lstSQL) = False Then
            SendMessageToLog("HOST_T_COMMAND_REPORT O_Add_Delete_SQLString Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
          End If
        Next
        If Common_DBManagement.BatchUpdate(lstSQL) = False Then
          SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
        '刪除HOST_T_COMMAND_REPORT的GDIC紀錄
        For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
          gMain.objHandling.O_Remove_HOST_T_COMMAND_REPORT(objHOST_T_COMMAND_REPORT)
        Next
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
#End Region

#Region "GUI的部分"
  ''' <summary>
  ''' 處理使用MQ交握GUI_Commad的部份(分3部份中的第1部份)
  ''' </summary>
  Public Sub O_thr_GUIMQHandling()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        '顯示在FORM上的CLOCK
        'gMain.date_tGUI_CommandDBHandle = Now
        'If WMS_StopGetGUICommand Then
        '  SendMessageToLog("Process GUI Command3 Stop...", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  While WMS_StopGetGUICommand
        '    Threading.Thread.Sleep(2000)
        '  End While
        'End If
        'If Count < 10 Then
        '  Count += 1
        'Else
        '  Count = 0
        '  If gMain.int_tGUICommandDBHandle > 99 Then
        '    gMain.int_tGUICommandDBHandle = 0
        '  Else
        '    gMain.int_tGUICommandDBHandle += 1
        '  End If
        'End If
        I_ProcessFromGUIMessage()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_GUI_CommandMQHandling3 End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Private Function I_ProcessFromGUIMessage() As Boolean
    Try
      While gMain.dicGUI_TO_HANDLING_Queue.Any = True
        Dim objQueue = gMain.dicGUI_TO_HANDLING_Queue.First.Value
        '取得headers '這四個固定要有的  在監控MQ THREAD已經先CHECK過DATA才放到各系統的待處理THREAD中，因此DECODE會成功
        Dim FUNCTION_ID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        Dim UUID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
        Dim SYSTEM_TYPE As enuSystemType = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.SEND_SYSTEM.ToString))
        Dim UserID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.USER_ID.ToString))
        Dim ClientID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.CLIENT_ID.ToString))
        Dim ip = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.IP.ToString))
        Dim XmlMessage As String = clsRabbitMQ.O_Decode_Message(objQueue.Body)
        Dim Result As String = ""
        Dim ResultMessage As String = ""
        Dim Report As String = ""
        Dim Wait_UUID As String = ""
        Dim strLog As String = ""

        '如果為等待回覆的 則跳過 
        If gMain.objHandling.O_Get_Command_Report(SYSTEM_TYPE, UUID) = True Then
          '把事件加入Report中
          If gMain.dicGUI_TO_HANDLING_Queue_R.ContainsKey(objQueue.DeliveryTag) = False Then
            gMain.dicGUI_TO_HANDLING_Queue_R.Add(objQueue.DeliveryTag, objQueue)
          End If
          gMain.dicGUI_TO_HANDLING_Queue.Remove(objQueue.DeliveryTag)
          Continue While
        End If

        '把取得的GUICommand送出去進行處理
        If O_ProcessCommand(FUNCTION_ID, XmlMessage, ResultMessage, Wait_UUID) = True Then
          '如果Wait_UUID不為空時才把Wait_UUID填入
          If Wait_UUID = "" Then
            Result = "0"
            'ResultMessage = ""
          End If
        Else
          '執行失敗
          Result = "1"
        End If

        '根據Wait_UUID進行對應的處理
        If Wait_UUID = "" Then
          '回覆處理結果
          If gMain.RabbitMQ.ResultSecondaryMessage(Result, ResultMessage, enuRabbitMQ.HOST_TO_GUI.ToString, "", "", objQueue, strLog) = False Then
            SendMessageToLog($"ResultSecondaryMessage Falied : {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If

          'Ack訊息 回覆MQ這筆已處理完成的訊息可以結束掉
          If gMain.RabbitMQ.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, strLog) = False Then
            SendMessageToLog($"Ack_ReceiveMessage Falied : {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If

          Dim Headers = objQueue.BasicProperties.Headers
          I_Write_HOST_T_COMMAND_HIS(XmlMessage, Headers, enuSystemType.GUI, Result, ResultMessage, Wait_UUID)
        Else '如果有Wait UUID
          Dim lstSQL As New List(Of String)
          Dim dicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)
          Dim objHOST_T_COMMAND_REPORT As New clsHOST_T_COMMAND_REPORT(Wait_UUID, SYSTEM_TYPE, UUID, GetNewTime_DBFormat)
          If dicHOST_T_COMMAND_REPORT.ContainsKey(objHOST_T_COMMAND_REPORT.gid) = False Then
            dicHOST_T_COMMAND_REPORT.Add(objHOST_T_COMMAND_REPORT.gid, objHOST_T_COMMAND_REPORT)
          End If
          For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
            If objHOST_T_COMMAND_REPORT.O_Add_Insert_SQLString(lstSQL) = False Then
              SendMessageToLog("HOST_T_COMMAND_REPORT O_Add_Insert_SQLString Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
            End If
          Next
          objHOST_T_COMMAND_REPORT.O_Add_Insert_SQLString(lstSQL)
          '寫入DB
          If Common_DBManagement.BatchUpdate(lstSQL) = False Then
            Return False
          End If

          '加入記憶體
          For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
            gMain.objHandling.O_Add_HOST_T_COMMAND_REPORT(objHOST_T_COMMAND_REPORT)
          Next

          '把事件加入Report中
          If gMain.dicGUI_TO_HANDLING_Queue_R.ContainsKey(objQueue.DeliveryTag) = False Then
            gMain.dicGUI_TO_HANDLING_Queue_R.Add(objQueue.DeliveryTag, objQueue)
          End If
        End If

        SendMessageToLog("Set GUI Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & Result & ", Result_Message=" & ResultMessage & ", Wait_UUID=" & Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        gMain.dicGUI_TO_HANDLING_Queue.Remove(objQueue.DeliveryTag)
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Sub O_thr_ToGUIMQHandling_Result()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        If Count < 10 Then
          Count += 1
        Else
          Count = 0
          'If gMain.int_tWMSToMCSResultDBHandle > 99 Then
          '  gMain.int_tWMSToMCSResultDBHandle = 0
          'Else
          '  gMain.int_tWMSToMCSResultDBHandle += 1
          'End If
        End If
        I_ProcessToGUIMessageResult()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_ToGUIMQHandling_Result End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Private Function I_ProcessToGUIMessageResult() As Boolean
    Try
      While gMain.dicGUI_TO_HANDLING_Queue_S.Any = True
        Dim objQueue = gMain.dicGUI_TO_HANDLING_Queue_S.First.Value
        '取得headers
        Dim FUNCTION_ID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        Dim UUID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
        Dim XmlMessageResult As String = clsRabbitMQ.O_Decode_Message(objQueue.Body)
        Dim Wait_UUID As String = ""
        Dim str_ResultMsg As String = ""
        Dim ReplyResult As Boolean = IIf(clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.RESULT.ToString)) = "0", True, False)
        Dim RejectReason As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.RESULT_MESSAGE.ToString))
        Dim XmlMessage As String = ""
        Dim dicProcessFromHostCommand As New Dictionary(Of String, clsFromHostCommand)

        ''處理時參考WMS有將發出的COMMAND寫到HOST_T_COMMAND
        Dim dicFromHostCommandResult = HOST_T_CommandManagement.GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(enuSystemType.GUI)
        For Each objFromHostCommandResult In dicFromHostCommandResult.Values
          '不相等表示這一筆不是當前要處理的命令，下一次再處理
          If UUID <> objFromHostCommandResult.UUID OrElse FUNCTION_ID <> objFromHostCommandResult.Function_ID Then
            Continue For
          End If
          If dicProcessFromHostCommand.ContainsKey(objFromHostCommandResult.gid) = False Then
            dicProcessFromHostCommand.Add(objFromHostCommandResult.gid, objFromHostCommandResult)
          Else
            SendMessageToLog(String.Format("FromHostCommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objFromHostCommandResult.UUID, objFromHostCommandResult.Function_ID, objFromHostCommandResult.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
          XmlMessage = XmlMessage & objFromHostCommandResult.Message
        Next

        '記錄發生給GUI後回覆的給結果
        If ReplyResult = True Then
          SendMessageToLog("Get Send To GUI Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog("Get Send To GUI Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If

        If O_ProcessCommandResult(UUID, FUNCTION_ID, XmlMessage, RejectReason, ReplyResult, str_ResultMsg) = True Then
          If ProcessHostCommandResult_Move_to_Hist(dicProcessFromHostCommand) = True Then

          End If
        End If
        'Ack訊息
        If gMain.RabbitMQ.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, str_ResultMsg) = False Then
          SendMessageToLog($"Ack_ReceiveMessage Falied : {str_ResultMsg}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        gMain.dicGUI_TO_HANDLING_Queue_S.Remove(objQueue.DeliveryTag)
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Set_GUI_MQWaitCommandResult(ByRef objHOST_T_COMMAND_REPORT As clsHOST_T_COMMAND_REPORT, ByVal WAIT_UUID As String, ByVal bln_Result As Boolean, ByVal str_RejectReason As String) As Boolean
    Try
      Dim Result As String = IIf(bln_Result = True, 0, 1)
      Dim strLog As String = ""
      Dim lstSQL As New List(Of String)
      '取得MQ_R資訊
      Dim objQueue As RabbitMQ.Client.Events.BasicDeliverEventArgs = Nothing
      If O_Get_objMQEventByUUID(gMain.dicGUI_TO_HANDLING_Queue_R, objHOST_T_COMMAND_REPORT.REPORT_SYSTEM_UUID, objQueue) = True Then
        '發Secondary
        Dim FUNCTION_ID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        If gMain.RabbitMQ.ResultSecondaryMessage(Result, str_RejectReason, enuRabbitMQ.HOST_TO_GUI.ToString, "", "", objQueue, strLog, WAIT_UUID:=WAIT_UUID) = False Then
          SendMessageToLog($"ResultSecondaryMessage Failed {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        'HOST_H_COMMAND_HIST 歷史紀錄
        Dim Headers = objQueue.BasicProperties.Headers
        I_Write_HOST_T_COMMAND_HIS(strLog, Headers, strLog, WAIT_UUID, Result, str_RejectReason)

        '刪除HOST_T_COMMAND_REPORT的DB紀錄
        Dim dicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)
        If dicHOST_T_COMMAND_REPORT.ContainsKey(objHOST_T_COMMAND_REPORT.gid) = False Then
          dicHOST_T_COMMAND_REPORT.Add(objHOST_T_COMMAND_REPORT.gid, objHOST_T_COMMAND_REPORT)
        End If
        For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
          If objHOST_T_COMMAND_REPORT.O_Add_Delete_SQLString(lstSQL) = False Then
            SendMessageToLog("HOST_T_COMMAND_REPORT O_Add_Delete_SQLString Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
          End If
        Next
        If Common_DBManagement.BatchUpdate(lstSQL) = False Then
          SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
        '刪除HOST_T_COMMAND_REPORT的GDIC紀錄
        For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
          gMain.objHandling.O_Remove_HOST_T_COMMAND_REPORT(objHOST_T_COMMAND_REPORT)
        Next
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
#End Region

#Region "NS的部分"
  ''' <summary>
  ''' 處理使用MQ交握NS_Commad的部份(分3部份中的第1部份)
  ''' </summary>
  Public Sub O_thr_NSMQHandling()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        '顯示在FORM上的CLOCK
        'gMain.date_tNS_CommandDBHandle = Now
        'If WMS_StopGetNSCommand Then
        '  SendMessageToLog("Process NS Command3 Stop...", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  While WMS_StopGetNSCommand
        '    Threading.Thread.Sleep(2000)
        '  End While
        'End If
        'If Count < 10 Then
        '  Count += 1
        'Else
        '  Count = 0
        '  If gMain.int_tNSCommandDBHandle > 99 Then
        '    gMain.int_tNSCommandDBHandle = 0
        '  Else
        '    gMain.int_tNSCommandDBHandle += 1
        '  End If
        'End If
        I_ProcessFromNSMessage()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_NS_CommandMQHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Private Function I_ProcessFromNSMessage() As Boolean
    Try
      While gMain.dicNS_TO_HANDLING_Queue.Any = True
        Dim objQueue = gMain.dicNS_TO_HANDLING_Queue.First.Value
        '取得headers '這四個固定要有的  在監控MQ THREAD已經先CHECK過DATA才放到各系統的待處理THREAD中，因此DECODE會成功
        Dim FUNCTION_ID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        Dim UUID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
        Dim SYSTEM_TYPE As enuSystemType = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.SEND_SYSTEM.ToString))
        Dim UserID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.USER_ID.ToString))
        Dim ClientID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.CLIENT_ID.ToString))
        Dim ip = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.IP.ToString))
        Dim XmlMessage As String = clsRabbitMQ.O_Decode_Message(objQueue.Body)
        Dim Result As String = ""
        Dim ResultMessage As String = ""
        Dim Report As String = ""
        Dim Wait_UUID As String = ""
        Dim strLog As String = ""

        '如果為等待回覆的 則跳過 
        If gMain.objHandling.O_Get_Command_Report(SYSTEM_TYPE, UUID) = True Then
          '把事件加入Report中
          If gMain.dicNS_TO_HANDLING_Queue_R.ContainsKey(objQueue.DeliveryTag) = False Then
            gMain.dicNS_TO_HANDLING_Queue_R.Add(objQueue.DeliveryTag, objQueue)
          End If
          gMain.dicNS_TO_HANDLING_Queue.Remove(objQueue.DeliveryTag)
          Continue While
        End If

        '把取得的NSCommand送出去進行處理
        If O_ProcessCommand(FUNCTION_ID, XmlMessage, ResultMessage, Wait_UUID) = True Then
          '如果Wait_UUID不為空時才把Wait_UUID填入
          If Wait_UUID = "" Then
            Result = "0"
            'ResultMessage = ""
          End If
        Else
          '執行失敗
          Result = "1"
        End If

        '根據Wait_UUID進行對應的處理
        If Wait_UUID = "" Then
          '回覆處理結果
          If gMain.RabbitMQ.ResultSecondaryMessage(Result, ResultMessage, enuRabbitMQ.HOST_TO_NS.ToString, "", "", objQueue, strLog) = False Then
            SendMessageToLog($"ResultSecondaryMessage Falied : {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If

          'Ack訊息 回覆MQ這筆已處理完成的訊息可以結束掉
          If gMain.RabbitMQ.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, strLog) = False Then
            SendMessageToLog($"Ack_ReceiveMessage Falied : {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If

          Dim Headers = objQueue.BasicProperties.Headers
          I_Write_HOST_T_COMMAND_HIS(XmlMessage, Headers, enuSystemType.NS, Result, ResultMessage, Wait_UUID)
        Else '如果有Wait UUID
          Dim lstSQL As New List(Of String)
          Dim dicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)
          Dim objHOST_T_COMMAND_REPORT As New clsHOST_T_COMMAND_REPORT(Wait_UUID, SYSTEM_TYPE, UUID, GetNewTime_DBFormat)
          If dicHOST_T_COMMAND_REPORT.ContainsKey(objHOST_T_COMMAND_REPORT.gid) = False Then
            dicHOST_T_COMMAND_REPORT.Add(objHOST_T_COMMAND_REPORT.gid, objHOST_T_COMMAND_REPORT)
          End If
          For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
            If objHOST_T_COMMAND_REPORT.O_Add_Insert_SQLString(lstSQL) = False Then
              SendMessageToLog("HOST_T_COMMAND_REPORT O_Add_Insert_SQLString Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
            End If
          Next
          objHOST_T_COMMAND_REPORT.O_Add_Insert_SQLString(lstSQL)
          '寫入DB
          If Common_DBManagement.BatchUpdate(lstSQL) = False Then
            Return False
          End If

          '加入記憶體
          For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
            gMain.objHandling.O_Add_HOST_T_COMMAND_REPORT(objHOST_T_COMMAND_REPORT)
          Next

          '把事件加入Report中
          If gMain.dicNS_TO_HANDLING_Queue_R.ContainsKey(objQueue.DeliveryTag) = False Then
            gMain.dicNS_TO_HANDLING_Queue_R.Add(objQueue.DeliveryTag, objQueue)
          End If
        End If

        SendMessageToLog("Set NS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & Result & ", Result_Message=" & ResultMessage & ", Wait_UUID=" & Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        gMain.dicNS_TO_HANDLING_Queue.Remove(objQueue.DeliveryTag)
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Sub O_thr_ToNSMQHandling_Result()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        If Count < 10 Then
          Count += 1
        Else
          Count = 0
          'If gMain.int_tWMSToNSResultDBHandle > 99 Then
          '  gMain.int_tWMSToNSResultDBHandle = 0
          'Else
          '  gMain.int_tWMSToNSResultDBHandle += 1
          'End If
        End If
        I_ProcessToNSMessageResult()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_WMSToNSResultDBHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Private Function I_ProcessToNSMessageResult() As Boolean
    Try
      While gMain.dicNS_TO_HANDLING_Queue_S.Any = True
        Dim objQueue = gMain.dicNS_TO_HANDLING_Queue_S.First.Value
        '取得headers '這四個固定要有的
        Dim FUNCTION_ID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        Dim UUID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
        Dim XmlMessageResult As String = clsRabbitMQ.O_Decode_Message(objQueue.Body)
        Dim Wait_UUID As String = ""
        Dim str_ResultMsg As String = ""
        Dim ReplyResult As Boolean = IIf(clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.RESULT.ToString)) = "0", True, False)
        Dim RejectReason As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.RESULT_MESSAGE.ToString))

        'MQ收完訊息，再到HOST_T_Command取RECEIVE = NS的資料
        Dim XmlMessage As String = ""
        Dim dicProcessFromHostCommand As New Dictionary(Of String, clsFromHostCommand)

        'WMS那隻是讀這個TABLE
        Dim dicFromHostCommandResult = HOST_T_CommandManagement.GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(enuSystemType.NS)
        For Each objFromHostCommandResult In dicFromHostCommandResult.Values
          '不相等表示這一筆不是當前要處理的命令，下一次再處理
          If UUID <> objFromHostCommandResult.UUID OrElse FUNCTION_ID <> objFromHostCommandResult.Function_ID Then
            Continue For
          End If
          If dicProcessFromHostCommand.ContainsKey(objFromHostCommandResult.gid) = False Then
            dicProcessFromHostCommand.Add(objFromHostCommandResult.gid, objFromHostCommandResult)
          Else
            SendMessageToLog(String.Format("NSCommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objFromHostCommandResult.UUID, objFromHostCommandResult.Function_ID, objFromHostCommandResult.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
          XmlMessage = XmlMessage & objFromHostCommandResult.Message
        Next

        '記錄發生給NS後回覆的給結果
        If ReplyResult = True Then
          SendMessageToLog("Get Send To NS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog("Get Send To NS Command Data Report, UUID=" & UUID & ", Function_ID=" & FUNCTION_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        '把取得的ToNSCommand送出去進行處理
        If O_ProcessCommandResult(UUID, FUNCTION_ID, XmlMessage, RejectReason, ReplyResult, str_ResultMsg) = True Then
          If ProcessHostCommandResult_Move_to_Hist(dicProcessFromHostCommand) = True Then

          End If
        End If

        'Ack訊息
        If gMain.RabbitMQ.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, str_ResultMsg) = False Then
          SendMessageToLog($"Ack_ReceiveMessage Falied : {str_ResultMsg}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        gMain.dicNS_TO_HANDLING_Queue_S.Remove(objQueue.DeliveryTag)
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Set_NS_MQWaitCommandResult(ByRef objHOST_T_COMMAND_REPORT As clsHOST_T_COMMAND_REPORT, ByVal WAIT_UUID As String, ByVal bln_Result As Boolean, ByVal str_RejectReason As String) As Boolean
    Try
      Dim Result As String = IIf(bln_Result = True, 0, 1)
      Dim strLog As String = ""
      Dim lstSQL As New List(Of String)
      '取得MQ_R資訊
      Dim objQueue As RabbitMQ.Client.Events.BasicDeliverEventArgs = Nothing
      If O_Get_objMQEventByUUID(gMain.dicNS_TO_HANDLING_Queue_R, objHOST_T_COMMAND_REPORT.REPORT_SYSTEM_UUID, objQueue) = True Then
        '發Secondary
        Dim FUNCTION_ID = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
        If gMain.RabbitMQ.ResultSecondaryMessage(Result, str_RejectReason, enuRabbitMQ.HOST_TO_NS.ToString, "", "", objQueue, strLog, WAIT_UUID:=WAIT_UUID) = False Then
          SendMessageToLog($"ResultSecondaryMessage Failed {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        'HOST_H_COMMAND_HIST 歷史紀錄
        Dim Headers = objQueue.BasicProperties.Headers
        I_Write_HOST_T_COMMAND_HIS(strLog, Headers, strLog, WAIT_UUID, Result, str_RejectReason)

        '刪除HOST_T_COMMAND_REPORT的DB紀錄
        Dim dicHOST_T_COMMAND_REPORT As New Dictionary(Of String, clsHOST_T_COMMAND_REPORT)
        If dicHOST_T_COMMAND_REPORT.ContainsKey(objHOST_T_COMMAND_REPORT.gid) = False Then
          dicHOST_T_COMMAND_REPORT.Add(objHOST_T_COMMAND_REPORT.gid, objHOST_T_COMMAND_REPORT)
        End If
        For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
          If objHOST_T_COMMAND_REPORT.O_Add_Delete_SQLString(lstSQL) = False Then
            SendMessageToLog("HOST_T_COMMAND_REPORT O_Add_Delete_SQLString Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
          End If
        Next
        If Common_DBManagement.BatchUpdate(lstSQL) = False Then
          SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
        '刪除HOST_T_COMMAND_REPORT的GDIC紀錄
        For Each objHOST_T_COMMAND_REPORT In dicHOST_T_COMMAND_REPORT.Values
          gMain.objHandling.O_Remove_HOST_T_COMMAND_REPORT(objHOST_T_COMMAND_REPORT)
        Next
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
#End Region

#Region "HostHandler 送給其他系統的部分"
  ''' <summary>
  ''' 用MQ送給其他系統，同時寫一筆資料到HOST_T_COMMAND當記錄
  ''' </summary>
  ''' <param name="strMessage"></param>
  ''' <param name="HeaderInfo"></param>
  ''' <param name="sendTo"></param>
  ''' <returns></returns>
  Public Function O_Send_MessageToOther_ByMQ(ByVal strMessage As String,
                                             ByVal HeaderInfo As eCA_TransactionMessage.clsHeader,
                                             ByVal sendTo As enuSystemType) As Boolean
    Try
      Dim strMQMessage As String = strMessage

      Dim dicToOtherCommand As New Dictionary(Of String, clsFromHostCommand)
      Dim DBMaxLength As Long = 3500
      '如果strMessage超過DBMaxLength個字就要分成多個Message(如果只有一個Message則SEQ填0)
      If strMessage.Length < DBMaxLength Then
        Dim objToCommand As clsFromHostCommand = Nothing
        If HeaderInfo.ClientInfo IsNot Nothing Then
          objToCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, 0, HeaderInfo.ClientInfo.UserID, HeaderInfo.ClientInfo.ClientID, HeaderInfo.ClientInfo.IP, GetNewTime_DBFormat, strMessage, "", "", "")
        Else
          objToCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, 0, "", "", "", GetNewTime_DBFormat, strMessage, "", "", "")
        End If
        If dicToOtherCommand.ContainsKey(objToCommand.gid) = False Then
          dicToOtherCommand.Add(objToCommand.gid, objToCommand)
        End If
      Else
        Dim Seq As Long = 1
        Do
          Dim NewMessage As String = ""
          If strMessage.Length > DBMaxLength Then
            NewMessage = strMessage.Substring(0, 3500)
            strMessage = strMessage.Substring(3500)
          Else
            NewMessage = strMessage.Substring(0, strMessage.Length)
            strMessage = ""
          End If
          Dim objToCommand As clsFromHostCommand = Nothing
          If HeaderInfo.ClientInfo IsNot Nothing Then
            objToCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, Seq, HeaderInfo.ClientInfo.UserID, HeaderInfo.ClientInfo.ClientID, HeaderInfo.ClientInfo.IP, GetNewTime_DBFormat, NewMessage, "", "", "")
          Else
            objToCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, Seq, "", "", "", GetNewTime_DBFormat, NewMessage, "", "", "")
          End If
          If dicToOtherCommand.ContainsKey(objToCommand.gid) = False Then
            dicToOtherCommand.Add(objToCommand.gid, objToCommand)
          End If
          Seq = Seq + 1
        Loop While (strMessage.Length > 0)
      End If

      '取得HostCommand的Insert SQL
      Dim lstSQL As New List(Of String)
      For Each objToOtherCommand As clsFromHostCommand In dicToOtherCommand.Values
        objToOtherCommand.O_Add_Insert_SQLString(lstSQL)
      Next

      '發送MQ事件
      If gMain.RabbitMQ.SendMessage(HeaderInfo.UUID, HeaderInfo.EventID, strMQMessage, Msg_Direction_Primary, "Handler", sendTo) = False Then
        Return False
      End If

      '寫入DB   留下HOST_H_COMMAND 事後要用來比對
      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        SendMessageToLog($"Common_DBManagement.BatchUpdate = False", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
#End Region

  Private Sub I_Write_HOST_T_COMMAND_HIS(ByVal strMessage As String,
                                         ByVal HeaderInfo As eCA_TransactionMessage.clsHeader,
                                         ByVal sendTo As enuSystemType,
                                         ByVal Result As String,
                                         ByVal Result_Str As String,
                                         ByVal Wait_UUID As String)
    Try
      '取得HostCommand的Insert SQL
      Dim lstSQL As New List(Of String)
      Dim strMQMessage As String = strMessage
      Dim DBMaxLength As Long = 3500
      Dim Now_Time As String = GetNewTime_DBFormat()

      '如果strMessage超過DBMaxLength個字就要分成多個Message(如果只有一個Message則SEQ填0)
      If strMessage.Length < DBMaxLength Then
        Dim objFromHostCommandHist As clsFromHostCommandHist = Nothing
        If HeaderInfo.ClientInfo IsNot Nothing Then
          objFromHostCommandHist = New clsFromHostCommandHist(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, 0, HeaderInfo.ClientInfo.UserID, Now_Time, strMessage, Result, Result_Str, Wait_UUID, Now_Time)
        Else
          objFromHostCommandHist = New clsFromHostCommandHist(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, 0, "", Now_Time, strMessage, Result, Result_Str, Wait_UUID, Now_Time)
        End If

        objFromHostCommandHist.O_Add_Insert_SQLString(lstSQL)
      Else
        Dim Seq As Long = 1
        Do
          Dim NewMessage As String = ""
          If strMessage.Length > DBMaxLength Then
            NewMessage = strMessage.Substring(0, 3500)
            strMessage = strMessage.Substring(3500)
          Else
            NewMessage = strMessage.Substring(0, strMessage.Length)
            strMessage = ""
          End If
          Dim objFromHostCommandHist As clsFromHostCommandHist = Nothing
          If HeaderInfo.ClientInfo IsNot Nothing Then
            objFromHostCommandHist = New clsFromHostCommandHist(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, Seq, HeaderInfo.ClientInfo.UserID, Now_Time, strMessage, Result, Result_Str, Wait_UUID, Now_Time)
          Else
            objFromHostCommandHist = New clsFromHostCommandHist(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, Seq, "", Now_Time, strMessage, Result, Result_Str, Wait_UUID, Now_Time)
          End If

          objFromHostCommandHist.O_Add_Insert_SQLString(lstSQL)
          Seq = Seq + 1
        Loop While (strMessage.Length > 0)
      End If

      '寫入DB   留下HOST_H_COMMAND_HIST的送出紀錄
      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        SendMessageToLog($"Common_DBManagement.BatchUpdate = False", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      End If

      SendMessageToLog($"Write HOST_H_COMMAND_HIST Finish", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Catch ex As Exception
      SendMessageToLog($"寫入HOST_H_COMMAND_HIST Failed : {ex.ToString}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
    End Try
  End Sub

  Private Function ProcessHostCommandResult_Move_to_Hist(ByVal dicFromHostCommand As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try
      Dim lstSQL As New List(Of String)
      Dim lstQueueSQL As New List(Of String)
      Dim Now_Time As String = GetNewTime_DBFormat()
      For Each objFromHostCommand As clsFromHostCommand In dicFromHostCommand.Values
        With objFromHostCommand
          '將WMS_T_MCSCommand寫到History
          Dim objFromHostCommandHist As New clsFromHostCommandHist(.UUID, .Send_System, .Receive_System, .Function_ID, .SEQ, .User_ID, .Create_Time, .Message, .Result, .Result_Message, .Wait_UUID, Now_Time)
          objFromHostCommandHist.O_Add_Insert_SQLString(lstQueueSQL)
        End With
        objFromHostCommand.O_Add_Delete_SQLString(lstSQL)
      Next
      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      End If
      Common_DBManagement.AddQueued_BatchUpdate(lstQueueSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function O_Get_objMQEventByUUID(ByRef dic As Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs),
                                               ByVal UUID As String,
                                               ByRef ret_obj As RabbitMQ.Client.Events.BasicDeliverEventArgs) As Boolean
    Try
      Dim tmp_dic = dic.ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
      Dim ret_dic = tmp_dic.Where(Function(obj)
                                    If obj.Value.BasicProperties.Headers.ContainsKey(enuMQHeaders.UUID.ToString) = False Then
                                      Return False
                                    End If
                                    If clsRabbitMQ.O_Decode_Message(obj.Value.BasicProperties.Headers(enuMQHeaders.UUID.ToString)) <> UUID Then
                                      Return False
                                    End If
                                    Return True
                                  End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        ret_obj = ret_dic.First.Value '只會有一個
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module