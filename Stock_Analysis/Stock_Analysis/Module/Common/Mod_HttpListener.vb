Imports System.Net
Imports System.Globalization
Imports System.Threading
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Newtonsoft.Json
Imports eCA_HostObject



Module HttpListener
  Sub Create_HttpListener()
    Dim prefixes As New List(Of String)

    prefixes.Add("http://127.0.0.1:8881/HttpListener/Outbound_Info/") 'prefixes(0)
    prefixes.Add("http://127.0.0.1:8881/HttpListener/PostingCheck/")  'prefixes(1)

    NonblockingListener_Other(prefixes)
  End Sub

  Sub Create_HttpListener_Noneblocking_WMS()
    Dim prefixes As New List(Of String)

    prefixes.Add(FROM_WMS_API_URL)

    NonblockingListener_WMS(prefixes)
  End Sub

  Sub Create_HttpListener_Noneblocking_GUI()
    Dim prefixes As New List(Of String)

    prefixes.Add(FROM_GUI_API_URL)

    NonblockingListener_GUI(prefixes)
  End Sub

  Sub Create_HttpListener_Noneblocking_MCS()
    Dim prefixes As New List(Of String)

    prefixes.Add(FROM_MCS_API_URL)

    NonblockingListener_MCS(prefixes)
  End Sub

  Sub Create_HttpListener_Noneblocking_NS()
    Dim prefixes As New List(Of String)

    prefixes.Add(FROM_NS_API_URL)

    NonblockingListener_NS(prefixes)
  End Sub

#Region "舊的WebAPI啟動方式"
  'Sub Create_HttpListener()
  '  Dim prefixes(1) As String
  '  prefixes(0) = "http://192.168.1.50:8881/HttpListener/Outbound_Info/"
  '  prefixes(1) = "http://192.168.1.50:8881/HttpListener/PostingCheck/"
  '  ProcessRequests(prefixes)
  'End Sub
  'Private Sub ProcessRequests(ByVal prefixes() As String)
  '  If Not System.Net.HttpListener.IsSupported Then
  '    SendMessageToLog("Windows XP SP2, Server 2003, or higher is required to " &
  '        "use the HttpListener class.", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
  '    Exit Sub
  '  End If

  '  ' URI prefixes are required,
  '  If prefixes Is Nothing OrElse prefixes.Length = 0 Then
  '    Throw New ArgumentException("prefixes")
  '  End If

  '  ' Create a listener and add the prefixes.
  '  Dim listener As System.Net.HttpListener =
  '      New System.Net.HttpListener()
  '  For Each s As String In prefixes
  '    listener.Prefixes.Add(s)
  '  Next

  '  While True
  '    Try
  '      ' Start the listener to begin listening for requests.
  '      listener.Start()
  '      SendMessageToLog("Listening...", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

  '      ' Set the number of requests this application will handle.
  '      Dim numRequestsToBeHandled As Integer = 10

  '      For i As Integer = 0 To numRequestsToBeHandled
  '        Try
  '          ' Note: GetContext blocks while waiting for a request.
  '          Dim context As HttpListenerContext = listener.GetContext()

  '          '得到key
  '          'context.Request.Headers.GetValues("ID")
  '          Dim body = New StreamReader(context.Request.InputStream).ReadToEnd()

  '          Dim Result_Message As String = ""
  '          '根據收到的訊息的URL作相對應的事件處理
  '          Dim Url = context.Request.Url.ToString
  '          If Url.Chars(Url.Length - 1) <> "/" Then
  '            Url += "/"
  '          End If
  '          Select Case Url
  '            Case prefixes(0) '給庫單
  '              '反陣列化 解json內容
  '              Dim Outbound_Info As New List(Of eCA_TransactionMessage.MSG_Outbound_Info) '根據之前先轉換好的類別宣告List
  '              eCA_TransactionMessage.ParseMessage_MSG_Outbound_Info(body, Outbound_Info, "")
  '              'Module_Send_Outbound_Info.O_Send_Outbound_Info(Outbound_Info, Result_Message)
  '              'Case prefixes(1) '回覆過帳結果
  '              '  '反陣列化 解json內容
  '              '  Dim PostingCheck As New List(Of eCA_TransactionMessage.MSG_PostingCheck) '根據之前先轉換好的類別宣告List
  '              '  eCA_TransactionMessage.ParseMessage_MSG_PostingCheck(body, PostingCheck, "")
  '              '  Module_PostingCheck.O_PostingCheck(PostingCheck, Result_Message)
  '          End Select


  '          '回覆資訊 '有需要再做調整
  '          ProcessRequest(context, Result_Message)

  '        Catch ex As HttpListenerException
  '          Console.WriteLine(ex.Message)
  '        End Try
  '      Next

  '    Catch ex As HttpListenerException
  '      Console.WriteLine(ex.Message)
  '      Exit While
  '    Finally
  '      ' Stop listening for requests.
  '      listener.Close()
  '      SendMessageToLog("Done Listening...", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

  '    End Try
  '    Thread.Sleep(100)
  '  End While

  'End Sub

  'Public Function SendRequest(uri As Uri, jsonDataBytes As Byte(), contentType As String, method As String) As String

  '  'Dim data = Encoding.UTF8.GetBytes(jsonSring)
  '  'Dim result_post = SendRequest(uri, data, "application/json", "POST")


  '  Dim req As WebRequest = WebRequest.Create(uri)
  '  req.ContentType = contentType
  '  req.Method = method
  '  req.ContentLength = jsonDataBytes.Length

  '  Dim stream = req.GetRequestStream()
  '  stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
  '  stream.Close()

  '  Dim response = req.GetResponse().GetResponseStream()

  '  Dim reader As New StreamReader(response)
  '  Dim res = reader.ReadToEnd()
  '  reader.Close()
  '  response.Close()

  '  Return res
  'End Function

  'Private Sub ProcessRequest(ByRef context As HttpListenerContext, ByVal Result_Message As String)
  '  Try
  '    Dim response As HttpListenerResponse = Nothing
  '    'Dim HreadString = ""
  '    'For Each key In context.Request.Headers.AllKeys
  '    '  HreadString += key + vbCrLf
  '    'Next
  '    'MsgBox(HreadString)

  '    ' Create the response.
  '    response = context.Response
  '    Dim responseString As String = ""

  '    '回覆的訊息
  '    If Result_Message = "" Then
  '      responseString =
  '        "<HTML><BODY>OK" &
  '        "</BODY></HTML>"
  '    Else
  '      responseString =
  '        "<HTML><BODY>NG: " &
  '        Result_Message &
  '        "</BODY></HTML>"
  '    End If


  '    Dim buffer() As Byte =
  '        System.Text.Encoding.UTF8.GetBytes(responseString)
  '    response.ContentLength64 = buffer.Length
  '    Dim output As System.IO.Stream = response.OutputStream
  '    output.Write(buffer, 0, buffer.Length)


  '    context.Response.Close()


  '  Catch ex As Exception
  '    Console.WriteLine(ex.Message)
  '  End Try
  'End Sub

#End Region
  Public Sub NonblockingListener_Other(ByVal prefixes As List(Of String))
    Dim listener As System.Net.HttpListener = New System.Net.HttpListener()
    For Each s As String In prefixes
      listener.Prefixes.Add(s)
      SendMessageToLog($"NonblockingListener:{s}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Next

    listener.Start()
    While True
      While listener.IsListening
        Dim result As IAsyncResult = listener.BeginGetContext(New AsyncCallback(AddressOf NonblockingListner_Other_Callback), listener)
        result.AsyncWaitHandle.WaitOne()
      End While
      Thread.Sleep(500) 'listener 可能被暫停了 稍等一下
    End While
  End Sub

  Public Sub NonblockingListner_Other_Callback(ByVal AsyncResult As IAsyncResult)
    Dim listener As Net.HttpListener = CType(AsyncResult.AsyncState, Net.HttpListener)
    Dim context As HttpListenerContext = listener.EndGetContext(AsyncResult)
    Dim Response_Message As String = String.Empty
    Thread.Sleep(10000)
    Try
      Dim body = New StreamReader(context.Request.InputStream).ReadToEnd()

      Dim Result_Message As String = ""

      Dim SystemID As enuSystemType = enuSystemType.None
      SendMessageToLog($"HttpListener, body = {body}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      Dim UUID = IIf(context.Request.Headers("UUID") IsNot Nothing, context.Request.Headers("UUID"), "")
      Dim FUNCTION_ID = IIf(context.Request.Headers("FUNCTION_ID") IsNot Nothing, context.Request.Headers("FUNCTION_ID"), "")
      Dim Result As String = "0"  '0:成功、1:失敗

      '確認來源訊息的內容格式，暫時還沒有用到，未來可以用這裡判斷轉JSON或XML
      Dim contentType As enuHTTPContentType = enuHTTPContentType.XML
      If context.Request.ContentType IsNot Nothing Then
        If context.Request.ContentType.ToUpper.IndexOf(Msg_Send_Type_JSON) > -1 Then
          contentType = enuHTTPContentType.JSON
        Else
          contentType = enuHTTPContentType.XML
        End If
      End If

#Region "這版的舊的啟動方式貼過來的內容"
      'context.Request.Url.ToString : Listener被觸發的網址
      Select Case context.Request.Url.ToString
        Case listener.Prefixes(0) '給庫單
          '反陣列化 解json內容
          Dim Outbound_Info As New List(Of eCA_TransactionMessage.MSG_Outbound_Info) '根據之前先轉換好的類別宣告List
          eCA_TransactionMessage.ParseMessage_MSG_Outbound_Info(body, Outbound_Info, "")
          I_Prepare_ResultMsg_Success("", Response_Message)
          'Module_Send_Outbound_Info.O_Send_Outbound_Info(Outbound_Info, Result_Message)
          'Case prefixes(1) '回覆過帳結果
          '  '反陣列化 解json內容
          '  Dim PostingCheck As New List(Of eCA_TransactionMessage.MSG_PostingCheck) '根據之前先轉換好的類別宣告List
          '  eCA_TransactionMessage.ParseMessage_MSG_PostingCheck(body, PostingCheck, "")
          '  Module_PostingCheck.O_PostingCheck(PostingCheck, Result_Message)
      End Select
#End Region

      '看HOSTHANDLER需不需要寫紀錄
      'O_Handle_HS_To_Host_Command_Hist(FUNCTION_ID, enuConnectionType.HttpListener, SystemID, UUID, GetNewTime_DBFormat, body, Result, Result_Message, "", "", "", "")

      '回覆資訊
      I_ProcessRequest_N(context, Response_Message)

    Catch ex As Exception
      '失敗了也要吐訊息回去，不然發送方會等到TIMEOUT
      I_Prepare_ResultMsg_NG(ex.ToString, Response_Message)
      I_ProcessRequest_N(context, Response_Message)

      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub


#Region "對內部系統"
  Public Sub NonblockingListener_WMS(ByVal prefixes As List(Of String))
    Dim listener As System.Net.HttpListener = New System.Net.HttpListener()
    For Each s As String In prefixes
      listener.Prefixes.Add(s)
      SendMessageToLog($"NonblockingListener:{s}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Next

    'listener.Start()
    While True
      While listener.IsListening
        Dim result As IAsyncResult = listener.BeginGetContext(New AsyncCallback(AddressOf ListenerCallback), listener)
        result.AsyncWaitHandle.WaitOne()
      End While
      Thread.Sleep(500) 'listener 可能被暫停了 稍等一下
    End While
  End Sub

  Public Sub NonblockingListener_GUI(ByVal prefixes As List(Of String))
    Dim listener As System.Net.HttpListener = New System.Net.HttpListener()
    For Each s As String In prefixes
      listener.Prefixes.Add(s)
      SendMessageToLog($"NonblockingListener:{s}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Next

    listener.Start()
    While True
      While listener.IsListening
        Dim result As IAsyncResult = listener.BeginGetContext(New AsyncCallback(AddressOf ListenerCallback), listener)
        result.AsyncWaitHandle.WaitOne()
      End While
      Thread.Sleep(500) 'listener 可能被暫停了 稍等一下
    End While
  End Sub

  Public Sub NonblockingListener_MCS(ByVal prefixes As List(Of String))
    Dim listener As System.Net.HttpListener = New System.Net.HttpListener()
    For Each s As String In prefixes
      listener.Prefixes.Add(s)
      SendMessageToLog($"NonblockingListener:{s}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Next

    listener.Start()
    While True
      While listener.IsListening
        Dim result As IAsyncResult = listener.BeginGetContext(New AsyncCallback(AddressOf ListenerCallback), listener)
        result.AsyncWaitHandle.WaitOne()
      End While
      Thread.Sleep(500) 'listener 可能被暫停了 稍等一下
    End While
  End Sub

  Public Sub NonblockingListener_NS(ByVal prefixes As List(Of String))
    Dim listener As System.Net.HttpListener = New System.Net.HttpListener()
    For Each s As String In prefixes
      listener.Prefixes.Add(s)
      SendMessageToLog($"NonblockingListener:{s}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
    Next

    listener.Start()
    While True
      While listener.IsListening
        Dim result As IAsyncResult = listener.BeginGetContext(New AsyncCallback(AddressOf ListenerCallback), listener)
        result.AsyncWaitHandle.WaitOne()
      End While
      Thread.Sleep(500) 'listener 可能被暫停了 稍等一下
    End While
  End Sub

  Public Sub ListenerCallback(ByVal AsyncResult As IAsyncResult)
    Dim listener As Net.HttpListener = CType(AsyncResult.AsyncState, Net.HttpListener)
    Dim context As HttpListenerContext = listener.EndGetContext(AsyncResult)
    Dim Response_Message As String = String.Empty
    Try
      Dim body = New StreamReader(context.Request.InputStream).ReadToEnd()

      Dim Result_Message As String = ""
      SendMessageToLog($"HttpListener, body = {body}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      Dim UUID = IIf(context.Request.Headers("UUID") IsNot Nothing, context.Request.Headers("UUID"), "")
      Dim FUNCTION_ID = IIf(context.Request.Headers("FUNCTION_ID") IsNot Nothing, context.Request.Headers("FUNCTION_ID"), "")
      Dim Result As String = "0"  '0:成功、1:失敗

      Dim SystemID As enuSystemType = enuSystemType.None
      Select Case context.Request.Url.ToString
        Case FROM_WMS_API_URL
          SystemID = enuSystemType.WMS
        Case FROM_GUI_API_URL
          SystemID = enuSystemType.GUI
        Case FROM_MCS_API_URL
          SystemID = enuSystemType.MCS
        Case FROM_NS_API_URL
          SystemID = enuSystemType.NS
      End Select

      '確認來源訊息的內容格式，暫時還沒有用到
      Dim contentType As enuHTTPContentType = enuHTTPContentType.XML
      If context.Request.ContentType.ToUpper.IndexOf(Msg_Send_Type_JSON) > -1 Then
        contentType = enuHTTPContentType.JSON
      Else
        contentType = enuHTTPContentType.XML
      End If

      If O_ProcessCommand(FUNCTION_ID, body, Result_Message, contentType) = True Then
        I_Prepare_ResultMsg_Success(Result_Message, Response_Message)
        Result = 0
      Else
        I_Prepare_ResultMsg_NG(Result_Message, Response_Message)
        Result = 1
      End If

      '看HOSTHANDLER需不需要寫紀錄
      O_Handle_HS_To_Host_Command_Hist(FUNCTION_ID, enuConnectionType.HttpListener, SystemID, UUID, GetNewTime_DBFormat, body, Result, Result_Message, "", "", "", "")

      '回覆資訊
      I_ProcessRequest_N(context, Response_Message)

    Catch ex As Exception
      '失敗了也要吐訊息回去，不然發送方會等到TIMEOUT
      I_Prepare_ResultMsg_NG(ex.ToString, Response_Message)
      I_ProcessRequest_N(context, Response_Message)

      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
#End Region

  Private Sub I_ProcessRequest_N(ByRef context As HttpListenerContext, ByVal Result_Message As String)
    Try
      Dim response As HttpListenerResponse = Nothing
      response = context.Response
      Dim responseString As String = Result_Message

      Dim buffer() As Byte = System.Text.Encoding.UTF8.GetBytes(responseString)
      response.ContentLength64 = buffer.Length
      Dim output As System.IO.Stream = response.OutputStream
      output.Write(buffer, 0, buffer.Length)
    Catch ex As Exception
      Console.WriteLine(ex.Message)
    Finally
      context.Response.Close()
    End Try
  End Sub

#Region "準備回傳的字串，如果需要用CLASS轉文字，請自行調整程式內容"
  Private Sub I_Prepare_ResultMsg_Success(ByVal Result_Message As String, ByRef Response_Message As String)
    Try
      Response_Message = "<HTML><BODY>OK" & "</BODY></HTML>"
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      Response_Message = "<HTML><BODY>NG:" & ex.ToString & "</BODY></HTML>"
    End Try
  End Sub


  Private Sub I_Prepare_ResultMsg_NG(ByVal Result_Message As String, ByRef Response_Message As String)
    Try
      Response_Message = "<HTML><BODY>NG: " & Result_Message & "</BODY></HTML>"
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      Response_Message = "<HTML><BODY>NG:" & ex.ToString & "</BODY></HTML>"
    End Try
  End Sub
#End Region
End Module






