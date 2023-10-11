Imports System.ServiceModel
Imports WCFService_WMS
Imports System.ServiceModel.Description
Imports System.Reflection
Imports eCA_TransactionMessage
Imports eCA_HostObject

Module Mod_WCFHost
  Public Host As ServiceHost '-20180730 Jerry

  Public HosteWMSMessage As ServiceHost               'Vito_20421
  '-call別人家的 '-待測試
  Public Function ERP_REPORT_Release(ByVal clsSendTransferDataToERP As MSG_SendTransferDataToERP, Optional ByRef ret_msg As String = "") As Boolean
    Try
      Dim test1 As ServiceReference1.TIPTOPServiceGateWayPortTypeClient = New ServiceReference1.TIPTOPServiceGateWayPortTypeClient
      test1.ClientCredentials.UserName.UserName = ClientCredentialsUserName
      test1.ClientCredentials.UserName.Password = ClientCredentialsPassword
      test1.InnerChannel.OperationTimeout = New TimeSpan(0, 0, WebServiceTimeOut)  '-設定Timeout 時-分-秒



      '塞參數 取得結果
      'Dim request As New ServiceReference1.XMLAdapterRequest
      Dim request As New ServiceReference1.UpdateWOIssueDataRequest
      Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, clsSendTransferDataToERP, ReturnMessage) = False Then
        Return False '將obj轉成xml
      End If

      SendMessageToLog("[Address]" & test1.Endpoint.Address.ToString, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      SendMessageToLog("[Request]" & strXML, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      'Dim Body As New ServiceReference1.UpdateWOIssueDataRequest
      'Body.XMLStrigFromWMS = strXML
      request.request = strXML

      ''傳送併取得回覆
      Try
        'Dim response = test1.ServiceReference1_DSCERPWSServiceSoap_XMLAdapter(request)
        'Dim response = test2.ServiceReference3_WS_E2bGWSoap_WMSToERP(request)
        Dim response = test1.ServiceReference1_TIPTOPServiceGateWayPortType_UpdateWOIssueData(request)
        '對方的回覆資訊
        SendMessageToLog("[Responce]" & response.response.ToString, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        'Dim clsresponse = Tools.ParseXmlStringToClass(Of MSG_ERPReport)(response.Body.XMLAdapterResult)

        'If clsresponse.Result = ERP_Result Then '-成功收到訊息
        '  Return True
        'Else
        '  Return False
        'End If
      Catch ex As Exception
        SendMessageToLog("STD_IN 无法连线", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return True
      End Try
      Return True


    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      MsgBox("NG")
      Return False
    End Try
  End Function
  Public Function ERP_REPORT_Stocktaking(ByVal clsSendTransferDataToERP As MSG_SendTransferDataToERP, Optional ByRef ret_msg As String = "") As Boolean
    Try
      Dim test1 As ServiceReference1.TIPTOPServiceGateWayPortTypeClient = New ServiceReference1.TIPTOPServiceGateWayPortTypeClient
      test1.ClientCredentials.UserName.UserName = ClientCredentialsUserName
      test1.ClientCredentials.UserName.Password = ClientCredentialsPassword
      test1.InnerChannel.OperationTimeout = New TimeSpan(0, 0, WebServiceTimeOut)  '-設定Timeout 時-分-秒



      '塞參數 取得結果
      'Dim request As New ServiceReference1.XMLAdapterRequest
      Dim request As New ServiceReference1.UpdateCountingLabelDataRequest
      Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, clsSendTransferDataToERP, ReturnMessage) = False Then
        Return False '將obj轉成xml
      End If

      SendMessageToLog("[Address]" & test1.Endpoint.Address.ToString, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      SendMessageToLog("[Request]" & strXML, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      'Dim Body As New ServiceReference1.UpdateWOIssueDataRequest
      'Body.XMLStrigFromWMS = strXML
      request.request = strXML

      ''傳送併取得回覆
      Try
        'Dim response = test1.ServiceReference1_DSCERPWSServiceSoap_XMLAdapter(request)
        'Dim response = test2.ServiceReference3_WS_E2bGWSoap_WMSToERP(request)
        Dim response = test1.ServiceReference1_TIPTOPServiceGateWayPortType_UpdateCountingLabelData(request)
        '對方的回覆資訊
        SendMessageToLog("[Responce]" & response.response.ToString, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        'Dim clsresponse = Tools.ParseXmlStringToClass(Of MSG_ERPReport)(response.Body.XMLAdapterResult)

        'If clsresponse.Result = ERP_Result Then '-成功收到訊息
        '  Return True
        'Else
        '  Return False
        'End If
      Catch ex As Exception
        SendMessageToLog("STD_IN 無法連線", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return True
      End Try
      Return True


    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      MsgBox("NG")
      Return False
    End Try
  End Function
  'SFCAB01-移轉單匯入WMS
  Public Function STD_IN(ByVal clsSendTransferDataToERP As MSG_SendTransferDataToERP, Optional ByRef ret_msg As String = "") As Boolean
    Try
      'Dim test1 As ServiceReference1.DSCERPWSServiceSoapClient = New ServiceReference1.DSCERPWSServiceSoapClient
      'test1.ClientCredentials.UserName.UserName = ClientCredentialsUserName
      'test1.ClientCredentials.UserName.Password = ClientCredentialsPassword
      'test1.InnerChannel.OperationTimeout = New TimeSpan(0, 0, WebServiceTimeOut)  '-設定Timeout 時-分-秒

      'Dim test2 As ServiceReference2.WS_E2bGWSoapClient = New ServiceReference2.WS_E2bGWSoapClient
      'Dim test2 As ServiceReference3.WS_E2bGWSoapClient = New ServiceReference3.WS_E2bGWSoapClient
      'test2.ClientCredentials.UserName.UserName = ClientCredentialsUserName
      'test2.ClientCredentials.UserName.Password = ClientCredentialsPassword
      'test2.InnerChannel.OperationTimeout = New TimeSpan(0, 0, WebServiceTimeOut)

      ''塞參數 取得結果
      ''Dim request As New ServiceReference1.XMLAdapterRequest
      'Dim request As New ServiceReference3.WMSToERPRequest
      'Dim strXML = ""
      'Dim ReturnMessage = ""
      'If PrepareMessage_SendTransferDataToERP(strXML, clsSendTransferDataToERP, ReturnMessage) = False Then
      '  Return False '將obj轉成xml
      'End If

      'SendMessageToLog("[Address]" & test2.Endpoint.Address.ToString, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      'SendMessageToLog("[Request]" & strXML, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      'Dim Body As New ServiceReference3.WMSToERPRequestBody
      'Body.XMLStrigFromWMS = strXML
      'request.Body = Body

      ''傳送併取得回覆
      'Try
      '  'Dim response = test1.ServiceReference1_DSCERPWSServiceSoap_XMLAdapter(request)
      '  Dim response = test2.ServiceReference3_WS_E2bGWSoap_WMSToERP(request)
      '  '對方的回覆資訊
      '  SendMessageToLog("[Responce]" & response.Body.WMSToERPResult, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      '  'Dim clsresponse = Tools.ParseXmlStringToClass(Of MSG_ERPReport)(response.Body.XMLAdapterResult)

      '  'If clsresponse.Result = ERP_Result Then '-成功收到訊息
      '  '  Return True
      '  'Else
      '  '  Return False
      '  'End If
      'Catch ex As Exception
      '  SendMessageToLog("STD_IN 无法连线", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return True
      'End Try
      Return True


    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      MsgBox("NG")
      Return False
    End Try
  End Function
  'WMS回報給ERP
  Public Function STD_IN(ByVal strXML As String, ByRef Result_Message As String, Optional ByRef ret_msg As String = "") As Boolean
    Try
      'Dim test1 As ServiceReference1.DSCERPWSServiceSoapClient = New ServiceReference1.DSCERPWSServiceSoapClient
      'test1.ClientCredentials.UserName.UserName = ClientCredentialsUserName
      'test1.ClientCredentials.UserName.Password = ClientCredentialsPassword
      'test1.InnerChannel.OperationTimeout = New TimeSpan(0, 0, WebServiceTimeOut)  '-設定Timeout 時-分-秒

      'Dim test2 As ServiceReference3.WS_E2bGWSoapClient = New ServiceReference3.WS_E2bGWSoapClient
      'test2.ClientCredentials.UserName.UserName = ClientCredentialsUserName
      'test2.ClientCredentials.UserName.Password = ClientCredentialsPassword
      'test2.InnerChannel.OperationTimeout = New TimeSpan(0, 0, WebServiceTimeOut)
      ''塞參數 取得結果
      'Dim request As New ServiceReference3.WMSToERPRequest
      ''Dim strXML = ""
      ''Dim ReturnMessage = ""
      ''If PrepareMessage_InventoryToERP(strXML, clsInventoryDataToERP, ReturnMessage) = False Then
      ''  Return False '將obj轉成xml
      ''End If

      'SendMessageToLog("[Request]" & strXML, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      'Dim Body As New ServiceReference3.WMSToERPRequestBody
      'Body.XMLStrigFromWMS = strXML
      'request.Body = Body

      ''傳送併取得回覆
      'Try
      '  Dim response = test2.ServiceReference3_WS_E2bGWSoap_WMSToERP(request)
      '  '對方的回覆資訊
      '  SendMessageToLog("[Responce]" & response.Body.WMSToERPResult, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      '  Dim xmlStr = response.Body.WMSToERPResult.ToString
      '  Dim startindex = xmlStr.IndexOf("<")
      '  Dim endindex = xmlStr.IndexOf(">")
      '  xmlStr = xmlStr.Remove(startindex, (endindex - startindex) + 1)
      '  xmlStr = xmlStr.Trim()
      '  Dim Result = ParseXmlStringToClass(Of MSG_Response)(xmlStr)

      '  For Each obj In Result.ReportData
      '    If obj.Result = "N" Then
      '      Result_Message = obj.ResultMessage
      '      Return False
      '    End If
      '  Next
      '  'Dim clsresponse = Tools.ParseXmlStringToClass(Of MSG_ERPReport)(response.Body.XMLAdapterResult)

      '  'If clsresponse.Result = ERP_Result Then '-成功收到訊息
      '  '  Return True
      '  'Else
      '  '  Return False
      '  'End If
      'Catch ex As Exception
      '  SendMessageToLog("STD_IN 无法连线", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End Try
      Return True


    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      MsgBox("NG")
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 开启WCF WebService
  ''' </summary>
  ''' <returns></returns>
  Function WCFHostOpen() As Integer
    Try
      gMain.SendMessageToLog("WCF Host Service Init From [" & WCFHostIpPort & "]", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      eWMSMessage.eWMSMessage.getWMSService = AddressOf HandleeWMSMessage                                       'Vito_20421
      'ProcessWCF_WebService() '實作Message內容

#Region "設定服務空間並啟用服務"
      Dim myeWMSMessage As New ServiceHost(GetType(eWMSMessage.eWMSMessage))
      Dim bindingeWMSMessage As BasicHttpBinding = New BasicHttpBinding
      bindingeWMSMessage.MaxReceivedMessageSize = Integer.MaxValue - 1
      bindingeWMSMessage.MaxBufferSize = Integer.MaxValue - 1
      bindingeWMSMessage.MaxBufferPoolSize = Long.MaxValue - 1
      bindingeWMSMessage.ReaderQuotas.MaxStringContentLength = Integer.MaxValue - 1
      bindingeWMSMessage.ReaderQuotas.MaxArrayLength = Integer.MaxValue - 1
      bindingeWMSMessage.ReaderQuotas.MaxDepth = Integer.MaxValue - 1
      bindingeWMSMessage.ReaderQuotas.MaxBytesPerRead = Integer.MaxValue - 1
      bindingeWMSMessage.ReaderQuotas.MaxNameTableCharCount = Integer.MaxValue - 1

      '建立ServiceHost物件裝載服務，並透過程式加入服務Endpoint
      myeWMSMessage.AddServiceEndpoint(GetType(eWMSMessage.IeWMSMessage), bindingeWMSMessage, "http://" & WCFHostIpPort & "/eWMSMessage")
      'myHostPickingList.AddServiceEndpoint(GetType(PickingList.IPickingList), bindingPickingList, "http://" & WCFHostIpPort & "/PickingList/")

      '可以在URI看到WSDL
      Dim metadatabehavioreWMSMessage = New ServiceMetadataBehavior()
      'Dim metadatabehaviorPickingList = New ServiceMetadataBehavior()

      metadatabehavioreWMSMessage.HttpGetEnabled = True
      metadatabehavioreWMSMessage.HttpGetUrl = New Uri("http://" & WCFHostIpPort & "/eWMSMessage/mex")
      myeWMSMessage.Description.Behaviors.Add(metadatabehavioreWMSMessage)

      SendMessageToLog("WebService eWMSMessage List服務地址: " & metadatabehavioreWMSMessage.HttpGetUrl.ToString, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      myeWMSMessage.Open()
      HosteWMSMessage = myeWMSMessage
#End Region


      gMain.SendMessageToLog("WCF Host Service is Running on [" & WCFHostIpPort & "]", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Return 0
    Catch ex As Exception
      gMain.SendMessageToLog("WCF Host Service fail on [" & WCFHostIpPort & "]", eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return -1
    End Try
  End Function


  ''' <summary>
  ''' 关闭WCF (WebService)
  ''' </summary>
  Sub WCFHostClose()
    Try
      Host.Close()
      gMain.SendMessageToLog("WCF Host Service is Close on [" & WCFHostIpPort & "]", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Catch ex As Exception
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  ''' <summary>
  ''' PickingLis 20200121
  ''' </summary>
  ''' <param name="input"></param>
  ''' <returns></returns> OK
  ''' Vito_20421
  Public Function HandleeWMSMessage(ByVal input As eWMSMessage.clsWMSService) As eWMSMessage.ReportData
    'Public Function HandleeWMSMessage(ByVal input As PickingList.clsPickingList) As PickingList.ReportData
    Dim code As String = "Y" '0成功 1失敗
    Dim msg As String = "" '失敗則回訊息
    Dim ReportData As New eWMSMessage.ReportData '回覆的資訊
    SendMessageToLog("eWMSMessgae From WebService", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Try
      '紀錄輸入參數
      If input IsNot Nothing Then
        SendMessageToLog("eWMSMessgae， 輸入參數： " & input.XML_Message.ToString, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        Dim Result_Message As String = ""
        Dim objMSG = Nothing
        Dim tmp_objMSG As MSG_Header = Nothing
        Dim xmlStr = input.XML_Message.ToString
        Dim EventID As String = ""

        '#Region "處理字串"
        '        xmlStr = xmlStr.Replace("<![CDATA[", "")
        '        xmlStr = xmlStr.Replace(" ]]>", "")
        '#Region "使用SOAP UI會碰到的"
        '        xmlStr = xmlStr.Replace("&", "")
        '        xmlStr = xmlStr.Replace(vbLf, "")
        '        xmlStr = xmlStr.Replace(vbTab, "")
        '        xmlStr = xmlStr.Replace(Chr(34), "")
        '#End Region
        '        xmlStr = xmlStr.Trim()
        '#End Region

        EventID = ParseXmlStringToClass(Of MSG_Header)(xmlStr.ToString).EventID

        'If O_ProcessCommand(EventID, xmlStr, msg) = True Then
        If I_Process_WCFCommand(EventID, xmlStr, msg) = True Then '原本使用的，等確認完要用的功能之後要改到O_ProcessCommand
          O_Handle_HS_To_Host_Command_Hist(EventID, enuConnectionType.WebService, enuSystemType.HostHandler, "", GetNewTime_DBFormat, xmlStr, "0", msg, "", "", "", "")
          code = "Y"
        Else
          O_Handle_HS_To_Host_Command_Hist(EventID, enuConnectionType.WebService, enuSystemType.HostHandler, "", GetNewTime_DBFormat, xmlStr, "1", msg, "", "", "", "")
          code = "N"
        End If

        SendMessageToLog("訊息執行完畢", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Else
        msg = "輸入XML為空值"
        O_Handle_HS_To_Host_Command_Hist("", enuConnectionType.WebService, enuSystemType.HostHandler, "", GetNewTime_DBFormat, "", "0", msg, "", "", "", "")
        code = "N"
      End If

      SendMessageToLog("eWMSMessgae 回報 " & msg, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
    Catch ex As Exception
      code = "N"
      msg = ex.InnerException.Message
      ReportData.RtnCode = code
      ReportData.RtnReason = msg
      Return ReportData
    End Try

    ReportData.RtnCode = code
    ReportData.RtnReason = msg  '沒錯誤的話基本上是空白，I_Process_WCHCommand return前也可以自由填入
    Return ReportData
  End Function

  Private Function I_Process_WCFCommand(ByVal EventID As String,
                                        ByVal xmlStr As String,
                                        ByRef ret_strResultMsg As String) As Boolean

    Try
      Dim objMSG = Nothing

      Select Case EventID
#Region "品號資料"
        Case "SendSKUData"
          objMSG = ParseXmlStringToClass(Of MSG_SendSKUData)(xmlStr.ToString)
          If O_SKUManagement_SendSKUData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region

#Region "入庫單 可用 MSG_SendInData 公版解出來的才放這邊，有多內容的就當成客製化另外一個CASE"
          'SendInboundData : 採購入庫
          'SendProduceInData : 生產入庫
          'SendSellReturnData : 退貨入庫
          'SendReturnData : 退料入庫
          'SendOtherInData : 雜收單
          '新增CASE請到 O_POManagement_SendInData 調整 POType_2 和 H_PO_ORDER_TYPE
        Case "SendBuyboundData", "SendInboundData", "SendProduceInData",
             "SendSellReturnData", "SendReturnData", "SendOtherInData"
          objMSG = ParseXmlStringToClass(Of MSG_SendInData)(xmlStr.ToString)
          If O_POManagement_SendInData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region

#Region "出庫單 可用 MSG_SendOutData 公版解出來的才放這邊，有多內容的就當成客製化另外一個CASE"
          'SendSellData : 產品出貨(銷貨)
          'SendPickUpData : 原料出庫
          'SendOtherOutData : 雜發單
          '新增CASE請到 O_POManagement_SendOutData 調整POType_2和Comment1~3
        Case = "SendSellData", "SendPickUpData", "SendOtherOutData"
          objMSG = ParseXmlStringToClass(Of MSG_SendOutData)(xmlStr.ToString)
          If O_POManagement_SendOutData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region

#Region "調撥單"
        Case "SendTransferData"
          objMSG = ParseXmlStringToClass(Of MSG_SendTransferData)(xmlStr.ToString)
          If O_POManagement_SendTransferData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region
#Region "未使用"
#Region "品號變更資料"
        Case "SendSKUChangeData"
          objMSG = ParseXmlStringToClass(Of MSG_SendSKUChangeData)(xmlStr.ToString)
          If O_SKUManagement_SendSKUChangeData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region
#Region "倉別資料"
        Case "SendWarehouseData"
          objMSG = ParseXmlStringToClass(Of MSG_SendWarehouseData)(xmlStr.ToString)
          If O_SendWarehouseData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region
#Region "託外進貨單"
          'Case "SendOutsourcePurchaseData"
          '  objMSG = ParseXmlStringToClass(Of MSG_SendOutsourceData)(xmlStr.ToString)
          '  If O_POManagement_SendOutsourcePurchaseData(objMSG, ret_strResultMsg) = False Then
          '    Return False
          '  End If
#End Region
#Region "庫存異動"
          'Case "SendTransactionData"
          '  objMSG = ParseXmlStringToClass(Of MSG_SendTransactionData)(xmlStr.ToString)
          '  If O_POManagement_SendTransactionData(objMSG, ret_strResultMsg) = False Then
          '    Return False
          '  End If
#End Region
#Region "轉播單"
          'Case "SendInventoryChangeData"
          '  objMSG = ParseXmlStringToClass(Of MSG_SendInventoryChangeData)(xmlStr.ToString)
          '  If O_POManagement_SendInventoryChangeData(objMSG, ret_strResultMsg) = False Then
          '    Return False
          '  End If
#End Region

#Region "暫出入單"
          'Case "SendTempInOutData"
          '  objMSG = ParseXmlStringToClass(Of MSG_SendTempInOutData)(xmlStr.ToString)
          '  If O_POManagement_SendTempInOutData(objMSG, ret_strResultMsg) = False Then
          '    Return False
          '  End If
#End Region
#Region "暫出入歸還單"
          'Case "SendTempInOutrReturnData"
          '  objMSG = ParseXmlStringToClass(Of MSG_SendTempInOutrReturnData)(xmlStr.ToString)
          '  If O_POManagement_SendTempInOutReturnData(objMSG, ret_strResultMsg) = False Then
          '    Return False
          '  End If
#End Region

#Region "採購入庫單"
        'Case "SendInboundData"
        '  objMSG = ParseXmlStringToClass(Of MSG_SendInboundData)(xmlStr.ToString)
        '  If O_POManagement_SendInboundData(objMSG, ret_strResultMsg) = False Then
        '    Return False
        '  End If
#End Region
#Region "退廠商單"
        Case "SendInboundReturnData"
          objMSG = ParseXmlStringToClass(Of MSG_SendInboundReturnData)(xmlStr.ToString)
          If O_POManagement_SendInboundReturnData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region

#Region "生產入庫單"
        'Case "SendProduceInData"
        '  objMSG = ParseXmlStringToClass(Of MSG_SendProduceInData)(xmlStr.ToString)
        '  If O_POManagement_SendProduceInData(objMSG, ret_strResultMsg) = False Then
        '    Return False
        '  End If
#End Region
#Region "雜發單"
        'Case "SendOtherOutData"
        '  objMSG = ParseXmlStringToClass(Of MSG_SendOtherOutData)(xmlStr.ToString)
        '  If O_POManagement_SendOtherOutData(objMSG, ret_strResultMsg) = False Then
        '    Return False
        '  End If
#End Region
#Region "雜收單"
        'Case "SendOtherInData"
        '  objMSG = ParseXmlStringToClass(Of MSG_SendOtherInData)(xmlStr.ToString)
        '  If O_POManagement_SendOtherInData(objMSG, ret_strResultMsg) = False Then
        '    Return False
        '  End If
#End Region

#Region "銷退單"
        'Case = "SendSellReturnData"
        '  objMSG = ParseXmlStringToClass(Of MSG_SendSellReturnData)(xmlStr.ToString)
        '  If O_POManagement_SendSellReturnData(objMSG, ret_strResultMsg) = False Then
        '    Return False
        '  End If
#End Region
#Region "貨主調撥單"
        Case "SendTransferOwnerData"
          objMSG = ParseXmlStringToClass(Of MSG_SendTransferOwnerData)(xmlStr.ToString)
          If O_POManagement_SendTransferOwnerData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region
#Region "盤點資料"

        Case "SendInventoryData"
          objMSG = ParseXmlStringToClass(Of MSG_SendInventoryData)(xmlStr.ToString)
          If O_StocktakingManagement_SendInventoryData(objMSG, ret_strResultMsg) = False Then
            Return False
          End If
#End Region
#Region "退料單"
          'Case "SendReturnData"
          '  objMSG = ParseXmlStringToClass(Of MSG_SendReturnData)(xmlStr.ToString)
          '  If O_POManagement_SendReturnData(objMSG, ret_strResultMsg) = False Then
          '    Return False
          '  End If
#End Region
#Region "品號換算單位"
          'Case "SendSKUUnitConversionData"
          '  objMSG = ParseXmlStringToClass(Of MSG_SendSKUUnitConversionData)(xmlStr.ToString)
          '  If O_POManagement_SendSKUUnitConversionData(objMSG, ret_strResultMsg) = False Then
          '    Return False
          '  End If
#End Region
#Region "領料單"
          'Case = "SendPickUpData"
          '  objMSG = ParseXmlStringToClass(Of MSG_SendPickUpData)(xmlStr.ToString)
          '  If O_POManagement_SendPickUpData(objMSG, ret_strResultMsg) = False Then
          '    Return False
          '  End If
#End Region
#End Region
        Case Else
          ret_strResultMsg = $"沒有對應的EventID : {EventID}"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
      End Select

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  '用WMS_WCFServiceClient套件發送
  'Public Function O_SendWCFToWMS(ByVal strXML As String, Optional ByRef strResultMsg As String = "") As Boolean
  '  Try
  '    Dim Binding As New BasicHttpBinding()
  '    Dim Address As New EndpointAddress(WMS_Service_Address)
  '    Binding.OpenTimeout = New TimeSpan(0, 0, 5)
  '    Binding.CloseTimeout = New TimeSpan(0, 0, 5)
  '    Binding.SendTimeout = New TimeSpan(0, 0, 10)
  '    Binding.ReceiveTimeout = New TimeSpan(0, 0, 5)

  '    Dim tmpWMS_WCFService As WMS_WCFService.WMS_WCFServiceClient = New WMS_WCFService.WMS_WCFServiceClient(Binding, Address)
  '    tmpWMS_WCFService.ClientCredentials.UserName.UserName = ClientCredentialsUserName
  '    tmpWMS_WCFService.ClientCredentials.UserName.Password = ClientCredentialsPassword

  '    'Xml的表頭
  '    Dim objMSG_Header As MSG_WMS_Message_Header = Nothing
  '    If Xml_To_Obj(strXML, objMSG_Header, strResultMsg) = True Then
  '      '建立事件紀錄
  '      TO_HS_Command_Record(enuSystemType.WMS, objMSG_Header.Header.EventID, objMSG_Header.Header.UUID, strXML)
  '    End If

  '    Try
  '      SendMessageToLog("[Request] " & strXML, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
  '      '輸入參數 取得結果
  '      Dim response = tmpWMS_WCFService.WCF_Command(strXML)
  '      SendMessageToLog("[Responce] " & response, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

  '      '結果轉為obj
  '      Dim objMessageResult As MSG_WMS_XML_Message_Result = Nothing
  '      If Xml_To_Obj(response, objMessageResult, strResultMsg) = False Then
  '        SendMessageToLog("response to xml failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If

  '      '回填處理結果並刪除
  '      TO_HS_Command_Result(enuSystemType.WMS, objMessageResult.Header.UUID, objMessageResult.Body.ResultInfo.Result, objMessageResult.Body.ResultInfo.ResultMessage)

  '      '檢查結果
  '      If objMessageResult.Body.ResultInfo.Result = enuResultInfo.OK Then '-成功收到訊息
  '        Return True
  '      Else
  '        strResultMsg = objMessageResult.Body.ResultInfo.ResultMessage
  '        Return False
  '      End If
  '    Catch ex As Exception
  '      SendMessageToLog("STD_IN 无法连线", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End Try
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
End Module
