Imports eCA_HostObject
Imports eCAMQTool.eCARabbitMQ
Imports eCA_TransactionMessage

Public Class clsMQ
  Private _blnEnable As Boolean = False
  Private _strMQIp As String = ""
  Private _strMQPort As String = ""
  Private _strUserName As String = "Admin"
  Private _strPassword As String = "Admin"
  Private _strExchangeName As String = "WMSDispatcher"
  Private _lstQueueNqme As New List(Of String)
  Public objRabbitMQ As clsRabbitMQ

  'GUI單據操作或花費長時間
  Private lstGUIFunctionID1 As New List(Of String) From {"T5F1U52_CancelPOToWO", "T5F1U54_BatchPOToWOBySortPOList", "T5F1U55_ResetPOByPO", "T5F1U61_CutBatchExecution", "T5F1U65_CutBatchSetSowingGroup", "T5F1U66_CutBatchSetTeminal", "T5F1U92_WOEmergencyOut_Customized", "T5F1U93_POBindCD_Customized", "T5F5U6_TransactionPOToWO", "T5F5U11_TransactionSourceAutoSet", "T5F5U12_TransactionSourceManagement", "T10F2U1_StocktakingManagement", "T10F2U2_StocktakingExecute", "T10F2U5_StocktakingClose"}
  'GUI揀貨相關操作
  Private lstGUIFunctionID2 As New List(Of String) From {"T5F3U111_PickUpPackage_BindDestCarrier", "T5F3U112_PickUpPackage_UnbindDestCarrier", "T5F3U113_PickUpPackage_BindSourceCarrier", "T5F3U114_PickUpPackage_AW", "T5F3U115_PickUpPackage_SourceCarrierOut", "T5F3U119_PickUpPackage_CarrierCompleted", "T5F3U121_PickUpPackage_NotAW", "T5F3U122_PickUpPackage_BindPackageAndLabel", "T5F3U123_PickUpPackage_ByCarrier", "T5F3U124_PickUpPackage_FullCarrier", "T5F3U131_PickUpBulk_NotAW", "T5F3U151_PickUpBulk_BindPackage", "T5F3U152_PickUpBulk_PackageToCarrier", "T5F3U161_PickUpTally_PackageToCarrier", "T5F3U162_PickUpTally_CarrierCompleted", "T5F3U171_PickUpDelivery_ByPackage", "T5F3U181_PickUpSowing_PackageToCarrier", "T5F3U191_PickUp_RestackPackage"}
  'WMS單據操作
  Private lstWMSFunctionID1 As New List(Of String) From {"T5F1U1_POManagement", "T5F1U16_POUpdate"}

  Sub New(ByVal Enable As Boolean,
          ByVal RabbitMQIp As String,
          ByVal RabbitMQPort As String)
    Try
      _blnEnable = Enable
      _strMQIp = RabbitMQIp
      _strMQPort = RabbitMQPort

      'GUI Queue Name (接收)
      _lstQueueNqme.Add(enuRabbitMQ.GUI_TO_HOST.ToString)

      'WMS Queue Name (接收)
      _lstQueueNqme.Add(enuRabbitMQ.WMS_TO_HOST.ToString)

      'MCS Queue Name (接收)
      _lstQueueNqme.Add(enuRabbitMQ.MCS_TO_HOST.ToString)

      'NS Queue Name (接收)
      _lstQueueNqme.Add(enuRabbitMQ.NS_TO_HOST.ToString)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub


  ''' <summary>
  ''' 开启 Rabbit MQ CreateConnection
  ''' </summary>
  Public Function I_Init_RabbitMQService() As Boolean
    Try
      If _blnEnable = False Then
        SendMessageToLog($"Rabbit MQ Enable is False", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If

      SendMessageToLog($"Rabbit MQ CreateConnection From {_strMQIp}:{_strMQPort}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Dim strLog As String = ""

      If objRabbitMQ IsNot Nothing Then
        objRabbitMQ.Dispose()
      End If

      objRabbitMQ = New clsRabbitMQ(_strMQIp, _strMQPort, "")
      objRabbitMQ.SetAccount(_strUserName, _strPassword)

      'Dim lstQueue As New List(Of String) From {
      '  strQueueName, 'Queue name '接收的
      '  strReultQueueName  '回覆的
      '  }

      If Not _lstQueueNqme.Any Then
        MsgBox("No Queue does Consumer Handling", MsgBoxStyle.Question)
        Return False
      End If

      If Not objRabbitMQ.CreateConnection(strLog) Then
        Debug.Print(strLog)
        Return False
      End If

      '讀取Queue 且不ack
      If objRabbitMQ.Register_ReciveQueue(_lstQueueNqme, False, strLog) = False Then
        SendMessageToLog($"Register_ReciveQueue Failed. {strLog}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Return False
      End If

      '把讀到的Queue存起來
      AddHandler objRabbitMQ.ReceivedData, Sub(model, objQueue)
                                             I_Process_MQ_QueeData(objQueue)
                                           End Sub

      SendMessageToLog($"Rabbit MQ is Running on {_strMQIp}:{_strMQPort}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      Return True
    Catch ex As Exception
      SendMessageToLog($"Rabbit MQ CreateConnection fail on {_strMQIp}:{_strMQPort}", eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function I_Check_MQ_Header(ByVal objQueue As RabbitMQ.Client.Events.BasicDeliverEventArgs) As Boolean
    Try
      Dim check_result As Boolean = True

      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.UUID.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.UUID.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If
      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.FUNCTION_ID.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.FUNCTION_ID.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If
      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.SEND_SYSTEM.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.SEND_SYSTEM.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If
      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.DIRECTION.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.DIRECTION.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If
      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.USER_ID.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.USER_ID.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If
      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.RECEIVE_SYSTEM.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.RECEIVE_SYSTEM.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If
      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.CLIENT_ID.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.CLIENT_ID.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If
      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.IP.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.IP.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If
      If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.CREATE_TIME.ToString) = False Then
        SendMessageToLog($"Headers {enuMQHeaders.CREATE_TIME.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        check_result = False
      End If

      Return check_result
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      Return False
    End Try
  End Function

  Private Sub I_Process_MQ_QueeData(ByVal objQueue As RabbitMQ.Client.Events.BasicDeliverEventArgs)

    Try
      If I_Check_MQ_Header(objQueue) = False Then
        '訊息已在處理時SENDTOLOG
        'CHECK成功才會被收進來
        Exit Sub
      End If

      Dim strMsg As String = clsRabbitMQ.O_Decode_Message(objQueue.Body, clsRabbitMQ.enuEncodingType.UTF8)

      Dim FUNCTION_ID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.FUNCTION_ID.ToString))
      Dim UUID As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.UUID.ToString))
      Dim SEND_SYSTEM As enuSystemType = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.SEND_SYSTEM.ToString))
      Dim DIRECTION As String = clsRabbitMQ.O_Decode_Message(objQueue.BasicProperties.Headers(enuMQHeaders.DIRECTION.ToString))
      SendMessageToLog($"Get MQ Message: UUID:{UUID}, FUNCTION_ID:{FUNCTION_ID}, SEND_SYSTEM:{SEND_SYSTEM}, DIRECTION:{DIRECTION}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      '根據系統分類

      Select Case SEND_SYSTEM
#Region "GUI, RF"
        Case enuSystemType.GUI, enuSystemType.RF
          If DIRECTION = Msg_Direction_Primary Then
            '根據事件進行分類。
            '參考的那隻WMS有啟動3個THREAD，因此有註冊三個LST去分開項目的處理QUEUE
            If lstGUIFunctionID1.Contains(FUNCTION_ID) = True Then
              If gMain.dicGUI_TO_HANDLING_Queue.ContainsKey(objQueue.DeliveryTag) = False Then
                gMain.dicGUI_TO_HANDLING_Queue.Add(objQueue.DeliveryTag, objQueue)
              End If
            ElseIf lstGUIFunctionID2.Contains(FUNCTION_ID) = True Then
              If gMain.dicGUI_TO_HANDLING_Queue.ContainsKey(objQueue.DeliveryTag) = False Then
                gMain.dicGUI_TO_HANDLING_Queue.Add(objQueue.DeliveryTag, objQueue)
              End If
            Else  '不在上述LST中的放到ELSE的THREAD處理，公版就只起一個THREAD，所以大家都放一起，要用的再自己調整
              If gMain.dicGUI_TO_HANDLING_Queue.ContainsKey(objQueue.DeliveryTag) = False Then
                gMain.dicGUI_TO_HANDLING_Queue.Add(objQueue.DeliveryTag, objQueue)
              End If
            End If
          Else 'Secondary
            If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.RESULT.ToString) = False Then
              SendMessageToLog($"Headers {enuMQHeaders.RESULT.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Exit Sub
            End If
            If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.RESULT_MESSAGE.ToString) = False Then
              SendMessageToLog($"Headers {enuMQHeaders.RESULT_MESSAGE.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Exit Sub
            End If
            If gMain.dicGUI_TO_HANDLING_Queue_S.ContainsKey(objQueue.DeliveryTag) = False Then
              gMain.dicGUI_TO_HANDLING_Queue_S.Add(objQueue.DeliveryTag, objQueue)
            End If
          End If
#End Region
#Region "WMS"
        Case enuSystemType.WMS
          If DIRECTION = Msg_Direction_Primary Then
            '根據事件進行分類
            If lstWMSFunctionID1.Contains(FUNCTION_ID) = True Then
              If gMain.dicWMS_TO_HANDLING_Queue.ContainsKey(objQueue.DeliveryTag) = False Then
                gMain.dicWMS_TO_HANDLING_Queue.Add(objQueue.DeliveryTag, objQueue)
              End If
            Else
              If gMain.dicWMS_TO_HANDLING_Queue.ContainsKey(objQueue.DeliveryTag) = False Then
                gMain.dicWMS_TO_HANDLING_Queue.Add(objQueue.DeliveryTag, objQueue)
              End If
            End If
          Else 'Secondary
            If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.RESULT.ToString) = False Then
              SendMessageToLog($"Headers {enuMQHeaders.RESULT.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Exit Sub
            End If
            If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.RESULT_MESSAGE.ToString) = False Then
              SendMessageToLog($"Headers {enuMQHeaders.RESULT_MESSAGE.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Exit Sub
            End If
            If lstWMSFunctionID1.Contains(FUNCTION_ID) = True Then
              If gMain.dicWMS_TO_HANDLING_Queue_S.ContainsKey(objQueue.DeliveryTag) = False Then
                gMain.dicWMS_TO_HANDLING_Queue_S.Add(objQueue.DeliveryTag, objQueue)
              End If
            Else
              If gMain.dicWMS_TO_HANDLING_Queue_S.ContainsKey(objQueue.DeliveryTag) = False Then
                gMain.dicWMS_TO_HANDLING_Queue_S.Add(objQueue.DeliveryTag, objQueue)
              End If
            End If
          End If
#End Region
#Region "MCS"
        Case enuSystemType.MCS
          If DIRECTION = Msg_Direction_Primary Then
            '根據事件進行分類
            If gMain.dicMCS_TO_HANDLING_Queue.ContainsKey(objQueue.DeliveryTag) = False Then
              gMain.dicMCS_TO_HANDLING_Queue.Add(objQueue.DeliveryTag, objQueue)
            End If
          Else 'Secondary
            If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.RESULT.ToString) = False Then
              SendMessageToLog($"Headers {enuMQHeaders.RESULT.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Exit Sub
            End If
            If objQueue.BasicProperties.Headers.ContainsKey(enuMQHeaders.RESULT_MESSAGE.ToString) = False Then
              SendMessageToLog($"Headers {enuMQHeaders.RESULT_MESSAGE.ToString} not Exist", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Exit Sub
            End If
            If gMain.dicMCS_TO_HANDLING_Queue_S.ContainsKey(objQueue.DeliveryTag) = False Then
              gMain.dicMCS_TO_HANDLING_Queue_S.Add(objQueue.DeliveryTag, objQueue)
            End If
          End If
#End Region
      End Select

    Catch ex As Exception
      SendMessageToLog($"依照系統或DIRECTION分配QUEUE時發生錯誤，objDeliveryTag : {objQueue.DeliveryTag}。ex : {ex.ToString}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
    End Try
  End Sub

  ''' <summary>
  ''' 关闭WCF (WebService)
  ''' </summary>
  Sub CloseConnection()
    Try
      Dim strLog As String = ""
      objRabbitMQ.CloseConnection(strLog)
      SendMessageToLog($"Rabbit MQ is Close on {_strMQIp}:{_strMQPort}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub


  ''' <summary>
  ''' <para>發送事件</para>
  ''' <para></para>
  ''' <para name="gUUID">gUUID :訊息的UUID</para>
  ''' <para name="gMessageID">gMessageID : Function_ID / Event_ID</para>
  ''' <para name="gSendQueueName">gSendQueueName : 要去的指定QueueName 若使用EXCHANGE就填空白</para>
  ''' <para name="gExchangeName">gExchangeName : 要透過的ExchangeName 一般會搭配 RoutingKey使用</para>
  ''' <para name="gRoutingKey">gRoutingKey : RoutingKey配合Exchange使用</para>
  ''' <para name="gReplyQueueName">gReplyQueueName  填空白</para>
  ''' <para name="gMessage">gMessage</para>
  ''' <para name="gHeaders">gHeaders</para>
  ''' <para name="retMsg">retMsg : Byref的最後結果</para>
  ''' </summary>
  ''' <returns></returns>
  Private Function MQSendMessage(ByVal gUUID As String,
                                 ByVal gMessageID As String,
                                 ByVal gSendQueueName As String,
                                 ByVal gExchangeName As String,
                                 ByVal gRoutingKey As String,
                                 ByVal gReplyQueueName As String,
                                 ByVal gMessage As String,
                                 ByVal gHeaders As Dictionary(Of String, Object),
                                 ByRef retMsg As String) As Boolean
    Try
      Dim objSendMessage = New clsRabbitMQ.SendMessage With {
        .gUUID = gUUID,
        .gMessageID = gMessageID,
        .gSendQueueName = gSendQueueName,
        .gExchangeName = gExchangeName,
        .gRoutingKey = gRoutingKey,
        .gReplyQueueName = gReplyQueueName, '可給回覆的QueueName
        .gMessage = gMessage,
        .gHeaders = gHeaders
      }
      If objRabbitMQ.Send(objSendMessage, retMsg) = False Then
        SendMessageToLog(retMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''' <summary>
  ''' <para>把結果到回覆的QUEUE，並把MQ的該筆紀錄ACK掉。</para>
  ''' <para>這裡呼叫 MQSendMessage 時要配合MQ的重送方式調整，可以只走QUEUE也可以走EXCHANGE</para>
  ''' </summary>
  ''' <param name="Result"></param>
  ''' <param name="ResultMessage"></param>
  ''' <param name="RoutingKey"></param>
  ''' <param name="objQueue"></param>
  ''' <param name="ret_Msg"></param>
  ''' <param name="WAIT_UUID"></param>
  ''' <returns></returns>
  Public Function ResultSecondaryMessage(ByVal Result As String,
                                         ByVal ResultMessage As String,
                                         ByVal QueueName As String,
                                         ByVal Exchange As String,
                                         ByVal RoutingKey As String,
                                         ByRef objQueue As RabbitMQ.Client.Events.BasicDeliverEventArgs,
                                         ByRef ret_Msg As String,
                                         Optional ByVal WAIT_UUID As String = "") As Boolean
    Try
      Dim Headers = objQueue.BasicProperties.Headers
      Dim UUID As String = clsRabbitMQ.O_Decode_Message(Headers(enuMQHeaders.UUID.ToString))
      Dim EventID As String = clsRabbitMQ.O_Decode_Message(Headers(enuMQHeaders.FUNCTION_ID.ToString))
      Dim SEND_SYSTEM As String = clsRabbitMQ.O_Decode_Message(Headers(enuMQHeaders.SEND_SYSTEM.ToString))
      Dim RECEIVE_SYSTEM As String = clsRabbitMQ.O_Decode_Message(Headers(enuMQHeaders.RECEIVE_SYSTEM.ToString))
      Dim USER_ID As String = clsRabbitMQ.O_Decode_Message(Headers(enuMQHeaders.USER_ID.ToString))
      Dim CLIENT_ID As String = clsRabbitMQ.O_Decode_Message(Headers(enuMQHeaders.CLIENT_ID.ToString))
      Dim IP As String = clsRabbitMQ.O_Decode_Message(Headers(enuMQHeaders.IP.ToString))
      Dim CREATE_TIME As String = clsRabbitMQ.O_Decode_Message(Headers(enuMQHeaders.CREATE_TIME.ToString))

      Dim objResult As New MSG_Secondary_Message
      objResult.Header.UUID = UUID
      objResult.Header.EventID = EventID
      objResult.Header.Direction = Msg_Direction_Secondary
      objResult.Body.ResultInfo.Result = Result
      objResult.Body.ResultInfo.ResultMessage = ResultMessage

      '統一處理 
      'DIRECTION
      If Headers.ContainsKey(enuMQHeaders.DIRECTION.ToString) = False Then
        Headers.Add(enuMQHeaders.DIRECTION.ToString, objResult.Header.Direction)
      Else
        Headers(enuMQHeaders.DIRECTION.ToString) = objResult.Header.Direction
      End If
      'RESULT
      If Headers.ContainsKey(enuMQHeaders.RESULT.ToString) = False Then
        Headers.Add(enuMQHeaders.RESULT.ToString, objResult.Body.ResultInfo.Result)
      Else
        Headers(enuMQHeaders.RESULT.ToString) = objResult.Body.ResultInfo.Result
      End If
      'RESULT_MESSAGE
      If Headers.ContainsKey(enuMQHeaders.RESULT_MESSAGE.ToString) = False Then
        Headers.Add(enuMQHeaders.RESULT_MESSAGE.ToString, objResult.Body.ResultInfo.ResultMessage)
      Else
        Headers(enuMQHeaders.RESULT_MESSAGE.ToString) = objResult.Body.ResultInfo.ResultMessage
      End If

      Dim strReportXML As String = ""
      '取得XML
      If PrepareMessage_Secondary_Message(strReportXML, objResult, ret_Msg) = False Then
        Return False
      End If

      '發送到指定的路徑
      If MQSendMessage(UUID, EventID, QueueName, Exchange, RoutingKey, "", strReportXML, Headers, ret_Msg) = False Then
        Return False
      End If

      '拆到外層去做，SEND歸SEND，ACK還是獨立拆成一個動作，不然會搞混函式的定位
      ''Ack訊息 回覆MQ這筆已處理完成的訊息可以結束掉
      'If Me.objRabbitMQ.Ack_ReceiveMessage(objQueue.DeliveryTag, ret_Msg) = False Then
      '  SendMessageToLog($"Ack_ReceiveMessage Falied : {ret_Msg}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  ''' <summary>
  ''' 發送事件
  ''' </summary>
  ''' <param name="gUUID"></param>
  ''' <param name="gFunction_ID"></param>
  ''' <param name="gMessage"></param>
  ''' <param name="gDirection"></param>
  ''' <param name="gUser_ID"></param>
  ''' <returns></returns>
  Public Function SendMessage(ByVal gUUID As String,
                              ByVal gFunction_ID As String,
                              ByVal gMessage As String,
                              ByVal gDirection As String,
                              ByVal gUser_ID As String,
                              ByVal RECEIVE_SYSTEM As String) As Boolean
    Try
      Dim strLog As String = ""
      Dim gSendQueueName As String = ""
      Dim gReplyQueueName As String = ""
      Dim gRoutingKey As String = ""
      Dim SEND_SYSTEM As String = enuSystemType.HostHandler '要轉成字串

      '建立Headers
      Dim gHeaders As New Dictionary(Of String, Object)
      gHeaders.Add(enuMQHeaders.FUNCTION_ID.ToString, gFunction_ID)
      gHeaders.Add(enuMQHeaders.UUID.ToString, gUUID)
      gHeaders.Add(enuMQHeaders.SEND_SYSTEM.ToString, SEND_SYSTEM)
      gHeaders.Add(enuMQHeaders.DIRECTION.ToString, gDirection)
      gHeaders.Add(enuMQHeaders.USER_ID.ToString, gUser_ID)
      gHeaders.Add(enuMQHeaders.RECEIVE_SYSTEM.ToString, RECEIVE_SYSTEM)
      gHeaders.Add(enuMQHeaders.CLIENT_ID.ToString, gUser_ID)
      gHeaders.Add(enuMQHeaders.IP.ToString, "")
      gHeaders.Add(enuMQHeaders.CREATE_TIME.ToString, GetNewTime_DBFormat)

      Select Case RECEIVE_SYSTEM
        Case enuSystemType.WMS
          gSendQueueName = enuRabbitMQ.HOST_TO_WMS.ToString
        Case enuSystemType.GUI
          gSendQueueName = enuRabbitMQ.HOST_TO_GUI.ToString
        Case enuSystemType.MCS
          gSendQueueName = enuRabbitMQ.HOST_TO_MCS.ToString
        Case enuSystemType.NS
          gSendQueueName = enuRabbitMQ.HOST_TO_NS.ToString
      End Select
      '發送
      If MQSendMessage(gUUID, gFunction_ID, gSendQueueName, "", "", "", gMessage, gHeaders, strLog) = False Then
        Return False
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Class
