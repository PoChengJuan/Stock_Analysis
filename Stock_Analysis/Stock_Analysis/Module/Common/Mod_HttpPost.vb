Imports System.ServiceModel
Imports WCFService_WMS
Imports System.ServiceModel.Description
Imports System.Reflection
Imports System
Imports System.IO
Imports System.Net
Imports System.Text

Module Mod_HttpPost
  Public Host As ServiceHost '-20180718 Jerry

  Function HpptHost() As Boolean
    Try

      ' 使用URL建立一個可以接收POST的請求
      Dim request As WebRequest = WebRequest.Create("http://www.contoso.com/PostAccepter.aspx ")
      ' 設定這個請求的方法屬性 (Method property) 為POST
      request.Method = "POST"

      ' 建立POST數據並將其轉換為 byte array（編碼為UTF8）
      Dim postData As String = "This is a test that posts this string to a Web server."
      Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)

      ' 為 WebRequest 設定內容類型 (ContentType)
      request.ContentType = "application/x-www-form-urlencoded"

      ' 為 WebRequest 設定內容長度 (ContentLength)
      request.ContentLength = byteArray.Length

      ' 取得請求串流
      Dim dataStream As Stream = request.GetRequestStream()

      ' 將資料寫入資料串流中
      dataStream.Write(byteArray, 0, byteArray.Length)

      ' 關閉串流物件
      dataStream.Close()


      '------------------------------------------------

      ' 取得回應.
      Dim response As WebResponse = request.GetResponse()

      ' 顯示狀態
      Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)

      ' 取得由server傳回的串流，及其內容
      dataStream = response.GetResponseStream()

      ' 使用StreamReader開啟串流，以便讀取內容
      Dim reader As New StreamReader(dataStream)

      ' 讀取內容
      Dim responseFromServer As String = reader.ReadToEnd()

      ' 顯示內容
      Console.WriteLine(responseFromServer)

      '--------------------------------------------------

      reader.Close()
      dataStream.Close()
      response.Close()

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  ''' <summary>
  ''' <para>發送REQUEST</para>
  ''' <para name="req">req : WebRequest物件</para>
  ''' <para name="uri">uri : 目標網址</para>
  ''' <para name="Request_Msg">Request_Msg : 發送的內容</para>
  ''' <para name="contentType">contentType : 發送的內容型態。範例:"application/json"，json可換成xml或text</para>
  ''' <para name="method">method : "POST"</para>
  ''' <para name="dicHeader">dicHeader : 用KEY，VALUE組好進來用FOR EACH直接展開</para>
  ''' <para name="timeout">timeout : 沒帶預設30秒</para>
  ''' </summary>
  ''' <returns></returns>
  Public Function SendHTTPRequest(ByRef req As WebRequest,
                                  ByVal uri As Uri,
                                  ByVal Request_Msg As String,
                                  ByVal contentType As String,
                                  ByVal method As String,
                                  ByVal dicHeader As Dictionary(Of String, String),
                                  Optional ByVal timeout As Integer = 30000) As Boolean
    Dim stream = Nothing
    Dim result As Boolean = False
    Try
      Dim DataBytes = Encoding.UTF8.GetBytes(Request_Msg)
      req = WebRequest.Create(uri)
      req.ContentType = contentType
      req.Method = method
      req.ContentLength = DataBytes.Length
      req.Timeout = timeout

      For Each objHeader In dicHeader
        req.Headers.Add(objHeader.Key, objHeader.Value)
      Next

      stream = req.GetRequestStream()
      stream.Write(DataBytes, 0, DataBytes.Length)


      result = True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    Finally
      If stream IsNot Nothing Then
        stream.Close()
      End If
    End Try

    Return result
  End Function

  Public Function GetHTTPResponse(ByRef req As WebRequest, ByRef response As String) As Boolean

    Dim stream As Stream = Nothing
    Dim reader As StreamReader = Nothing
    Dim result As Boolean = False

    Try

      stream = req.GetResponse().GetResponseStream()
      reader = New StreamReader(stream)

      response = reader.ReadToEnd()

      result = True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    Finally
      If stream IsNot Nothing Then
        stream.Close()
      End If
      If reader IsNot Nothing Then
        reader.Close()
      End If
    End Try

    Return result
  End Function
End Module
