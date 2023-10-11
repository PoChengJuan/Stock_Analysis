
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Module Tools

#Region "ToolClass"
    Public Class Reply_Info
        Public Property ReplyCode As Integer = 0
        Public Property ReplyMessage As String = "OK"
        Public Property ReplyData As Object
    End Class

    Public Class Page_Info
        Public Property PageStart As Integer
        Public Property PageEnd As Integer
        Public Property PageCounts As Integer
        Public Property PageIndex As Integer
        Public Property PageSize As Integer
        Public Property DataCounts As Integer
    End Class
#End Region

    ''' <summary> 條件式 WHERE IN 字串整理 List => String </summary>
    Public Function InListToString(_DataList As List(Of String)) As Reply_Info
        Dim retrun_info As New Reply_Info
        Try
            retrun_info.ReplyData = String.Empty
            If _DataList.Any Then
                For i As Integer = 0 To _DataList.Count - 1
                    retrun_info.ReplyData &= IIf(i <> _DataList.Count - 1, "'" & _DataList(i) & "',", "'" & _DataList(i) & "'")
                Next i
            Else
                retrun_info.ReplyData = "''"
            End If
        Catch ex As Exception
            retrun_info.ReplyCode = -1
            retrun_info.ReplyMessage = ex.Message
        End Try

        Return retrun_info
    End Function

    ''' <summary> 分頁控制 </summary>
    Public Function PageChange(_Page_Info As Page_Info) As Reply_Info
        Dim retrun_info As New Reply_Info
        Try
            retrun_info.ReplyData = _Page_Info
            retrun_info.ReplyData.PageCounts = retrun_info.ReplyData.DataCounts \ retrun_info.ReplyData.PageSize + IIf(CInt(retrun_info.ReplyData.DataCounts) Mod retrun_info.ReplyData.PageSize = 0, 0, 1)
            retrun_info.ReplyData.PageStart = (retrun_info.ReplyData.PageIndex - 1) * retrun_info.ReplyData.PageSize + 1
            retrun_info.ReplyData.PageEnd = IIf(retrun_info.ReplyData.PageIndex * retrun_info.ReplyData.PageSize < retrun_info.ReplyData.DataCounts, retrun_info.ReplyData.PageIndex * retrun_info.ReplyData.PageSize, retrun_info.ReplyData.DataCounts)

        Catch ex As Exception
            retrun_info.ReplyCode = -1
            retrun_info.ReplyMessage = ex.Message
        End Try

        Return retrun_info
    End Function

    ''' <summary> 獲得頁數Index </summary>
    Public Function GetPageIndex(_PageIndex As Integer, _PageCounts As Integer, _ObjectName As String) As Reply_Info
        Dim retrun_info As New Reply_Info
        Try
            retrun_info.ReplyData = _PageIndex

            Select Case _ObjectName
                Case "btn_First"
                    retrun_info.ReplyData = 1
                Case "btn_Previous"
                    retrun_info.ReplyData = IIf(_PageIndex - 1 > 0, _PageIndex - 1, 1)
                Case "btn_Next"
                    retrun_info.ReplyData = IIf(_PageIndex + 1 > _PageCounts, _PageIndex, _PageIndex + 1)
                Case "btn_Last"
                    retrun_info.ReplyData = _PageCounts
            End Select

        Catch ex As Exception
            retrun_info.ReplyCode = -1
            retrun_info.ReplyMessage = ex.Message
        End Try

        Return retrun_info
    End Function

    ''' <summary>
    ''' Dictionary擴充方法，讓Dictionary對另一個Dictionary直接相加
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <typeparam name="S"></typeparam>
    ''' <param name="_Source"></param>
    ''' <param name="_Collection"></param>
    <System.Runtime.CompilerServices.Extension>
    Public Sub AddRange(Of T, S)(_Source As Dictionary(Of T, S), _Collection As Dictionary(Of T, S))
        If _Collection Is Nothing Then
            Throw New ArgumentNullException("Collection is null")
        End If

        For Each item In _Collection
            If Not _Source.ContainsKey(item.Key) Then
                _Source.Add(item.Key, item.Value)
            End If
        Next
    End Sub

    Public Function ParseClassToXmlString(Of T)(_Class As T) As String
        Dim sw As StringWriter = New StringWriter()
        Dim ns As New XmlSerializerNamespaces()
        ns.Add("", "")
        Dim settings As New XmlWriterSettings()
        settings.OmitXmlDeclaration = True
        Dim x As New Xml.Serialization.XmlSerializer(_Class.GetType)
        x.Serialize(sw, _Class, ns)
        Return sw.ToString
    End Function

    Public Function ParseXmlStringToClass(Of T)(_XmlString As String) As T
        Try
            Dim ret_T = New XmlSerializer(GetType(T)).Deserialize(New StringReader(_XmlString))
            Return ret_T
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return Nothing
        End Try
    End Function

  Public Function GetRandString() As String
    Dim st1 As String = String.Empty

    Try
      Dim r1 As New Random
      Randomize()
      For i = 1 To 9
        st1 = st1 & r1.Next(0, 9)
      Next
    Catch ex As Exception
      'MosaWebService.g_LogManager.WriteLog(LogManager.EnumLog.Exception, "【GetRandString失敗】", ex.Message)
    End Try
    Return st1
  End Function

  ''' <summary>
  ''' 物件 轉 Json
  ''' </summary>
  ''' <param name="DATA"></param>
  ''' <returns></returns>
  Public Function ParseClassToJSONString(ByVal DATA As Object) As String
    Try
      'SendMessageToLog("Get JSON Start", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)


      Dim JSON_Str = JsonConvert.SerializeObject(DATA)
      Dim objJSON = JObject.Parse(JSON_Str)
      objJSON.Remove("gid")
      objJSON.Remove("DATA")
      JSON_Str = objJSON.ToString
      'SendMessageToLog(JSON_Str.ToString, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      'SendMessageToLog("Get JSON End", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Return JSON_Str
    Catch ex As Exception
      Dim ret_msg = ex.ToString
      SendMessageToLog(ret_msg, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function


  ''' <summary>
  ''' Json 轉 物件
  ''' </summary>
  ''' <param name="json"></param>
  ''' <returns></returns>
  Public Function ParseJSONStringToClass(Of T)(ByVal json As String) As T
    Try
      Dim ret_T = Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(json)
      Return ret_T
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  ''' <summary>yyyy/MM/dd HH:mm:ss </summary>
  Public Function GetDateTimeString() As String
        Return DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Function


    <System.Runtime.CompilerServices.Extension>
    Public Sub InsertToDbFromObject(Of T)(_Data As T)
    ' Dim sqlString = MosaWebService.g_DbManager.GetInsertSqlStringFromClass(_Data)
    Dim returnErrorString = String.Empty
    'If Not MosaWebService.g_DbManager.ExecuteSql(returnErrorString, sqlString, _Data) Then
    '    MosaWebService.g_LogManager.WriteLog(LogManager.EnumLog.Exception, "【InsertToDbFromObject】", returnErrorString)
    'End If
  End Sub

    <System.Runtime.CompilerServices.Extension>
    Public Sub UpdateToDbFromObject(Of T)(_Data As T, _WhereString As String)
    'Dim sqlString = MosaWebService.g_DbManager.GetUpdateSqlStringFromClass(_Data, _WhereString)
    Dim returnErrorString = String.Empty
    'If Not MosaWebService.g_DbManager.ExecuteSql(returnErrorString, sqlString) Then
    '    MosaWebService.g_LogManager.WriteLog(LogManager.EnumLog.Exception, "【UpdateToDbFromObject】", returnErrorString)
    'End If
  End Sub

    <System.Runtime.CompilerServices.Extension>
    Public Sub DeleteToDbFromObject(Of T)(_Data As T, _WhereString As String)
    'Dim sqlString = MosaWebService.g_DbManager.GetDeleteSqlStringFromClass(_Data, _WhereString)
    Dim returnErrorString = String.Empty
    'If Not MosaWebService.g_DbManager.ExecuteSql(returnErrorString, sqlString) Then
    '    MosaWebService.g_LogManager.WriteLog(LogManager.EnumLog.Exception, "【DeleteToDbFromObject】", returnErrorString)
    'End If
  End Sub








End Module
