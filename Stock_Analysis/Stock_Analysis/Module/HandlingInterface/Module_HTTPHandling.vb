Imports eCA_HostObject
Imports eCA_TransactionMessage
Imports System.Net

Module Module_HTTPHandling
  ''' <summary>
  ''' 對上位系統的WebAPI發送
  ''' <para name="strMessage">strMessage : 訊息內容>。已將OBJ轉成STRING</para>
  ''' <para name="url">sendTo : 目標IP</para>
  ''' <para name="dicCustomizeHeader">dicCustomizeHeader : 依照各API填Header傳進來(該廠固定項目可寫死到程式裡面)</para>
  ''' <para name="enuContentType">enuContentType : 發送的內容型態(預設走XML)</para>
  ''' </summary>
  ''' <returns></returns>
  Public Function O_Send_MessageToHost_ByHTTP(ByVal strMessage As String,
                                              ByVal url As String,
                                              ByVal dicCustomizeHeader As Dictionary(Of String, String),
                                              Optional ByVal enuContentType As enuHTTPContentType = enuHTTPContentType.XML) As Boolean
    Try
      Dim req As WebRequest = Nothing
      Dim method As String = "POST"
      Dim contentType As String = String.Empty
      Dim response As String = String.Empty

      Select Case enuContentType
        Case enuHTTPContentType.JSON
          contentType = Msg_Application_JSON
        Case enuHTTPContentType.XML
          contentType = Msg_Application_XML
      End Select

      '發送HTTP REQUET
      If SendHTTPRequest(req, New Uri(url), strMessage, contentType, method, dicCustomizeHeader) = False Then
        Return False
      End If

      SendMessageToLog($"Send WebAPI URL : {url} 。strMessage : {strMessage}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      '等待並取得回覆
      If GetHTTPResponse(req, response) = False Then
        Return False
      End If

      SendMessageToLog($"Receive WebAPI URL : {url} 。strMessage : {response}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''' <summary>
  ''' 用HTTP POST發送訊息
  ''' <para name="strMessage">strMessage : 訊息內容>。已將OBJ轉成STRING</para>
  ''' <para name="HeaderInfo">HeaderInfo : clsHeader的內容</para>
  ''' <para name="sendTo">sendTo : 發送的目標系統</para>
  ''' <para name="enuContentType">enuContentType : 發送的內容型態(目前內部系統好像沒有取來用，所以沒寫應該沒差，對外部系統舊看看對方有沒有在用)</para>
  ''' </summary>
  ''' <returns></returns>
  Public Function O_Send_MessageToOther_ByHTTP(ByVal strMessage As String,
                                               ByVal HeaderInfo As eCA_TransactionMessage.clsHeader,
                                               ByVal sendTo As enuSystemType,
                                               Optional ByVal enuContentType As enuHTTPContentType = enuHTTPContentType.JSON,
                                               Optional ByVal timeout As Integer = 30000) As Boolean
    Try
      Dim req As WebRequest = Nothing
      Dim url As String = String.Empty
      Dim dicHeader As New Dictionary(Of String, String)
      Dim contentType As String = String.Empty
      Dim method As String = "POST"
      Dim response As String = String.Empty
      Dim ret_strResultMsg As String = String.Empty
      Dim objResponse = Nothing

      If I_Get_Common_HTTP_Header(HeaderInfo, dicHeader) = False Then
        SendMessageToLog("I_Get_Common_HTTP_Header 取得基本HEADER失敗", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

      '假如要發送的系統有除了clsHeader以外的HEADER要讀取，另外加即可 EX : dicHeader.Add("PASSWORD", "password")
      Select Case sendTo
        Case enuSystemType.WMS
          url = TO_WMS_API_URL
        Case enuSystemType.MCS
          url = TO_MCS_API_URL
        Case enuSystemType.GUI
          url = TO_GUI_API_URL
      End Select

      'contentType的測試結果，帶JSON或TEXT都沒差，主要是看接收端那邊有沒有要參考，沒有的話就沒影響
      Select Case enuContentType
        Case enuHTTPContentType.JSON
          contentType = Msg_Application_JSON
        Case enuHTTPContentType.XML
          contentType = Msg_Application_XML
      End Select

      '發送HTTP REQUET
      If SendHTTPRequest(req, New Uri(url), strMessage, contentType, method, dicHeader, timeout) = False Then
        Return False
      End If

      SendMessageToLog($"Send WebAPI URL : {url} 。strMessage : {strMessage}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      '等待並取得回覆
      If GetHTTPResponse(req, response) = False Then
        Return False
      End If

      SendMessageToLog($"Receive WebAPI URL : {url} 。strMessage : {response}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      '處理回覆後的動作，解MSG_TO_CLASS
      Select Case enuContentType
        Case enuHTTPContentType.JSON
          objResponse = ParseJSONStringToClass(Of MSG_WMS_JSON_Message_Result)(response)
        Case enuHTTPContentType.XML
          objResponse = ParseXmlStringToClass(Of MSG_WMS_XML_Message_Result)(response)
      End Select

      If objResponse IsNot Nothing Then
        If O_ProcessCommandResult(HeaderInfo.UUID, HeaderInfo.EventID, response, objResponse.Body.ResultInfo.ResultMessage, objResponse.Body.ResultInfo.Result, ret_strResultMsg) = True Then

        End If
      Else
        SendMessageToLog("obj is nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''' <summary>
  ''' 將clsHeader轉成DIC方便操作
  ''' </summary>
  ''' <param name="header"></param>
  ''' <param name="dicHeader"></param>
  ''' <returns></returns>
  Private Function I_Get_Common_HTTP_Header(ByVal header As clsHeader, ByRef dicHeader As Dictionary(Of String, String)) As Boolean
    Try
      dicHeader.Add("UUID", header.UUID)
      dicHeader.Add("SEND_SYSTEM_ID", enuSystemType.HostHandler.ToString)
      dicHeader.Add("SEND_SYSTEM_NO", enuSystemType.HostHandler)
      dicHeader.Add("DIRECTION", header.Direction)
      dicHeader.Add("FUNCTION_ID", header.EventID)
      dicHeader.Add("USER_ID", header.ClientInfo.UserID)
      dicHeader.Add("CLIENT_ID", header.ClientInfo.ClientID)
      dicHeader.Add("IP", header.ClientInfo.IP)
      dicHeader.Add("CREATE_TIME", GetNewTime_DBFormat)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
