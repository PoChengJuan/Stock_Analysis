Public Class clsHSToHostCommand
  Private ShareName As String = "HSToHostCommand"
  Private ShareKey As String = ""
  Private _gid As String
  Private _UUID As String '流水號

  Private _CONNECTION_TYPE As Long '連線類型遇到新的再增加Enum

  Private _SEND_SYSTEM As Long '發送的系統名稱遇到新的再增加Enum

  Private _FUNCTION_ID As String '交易編號

  Private _SEQ As Long '序列號

  Private _USER_ID As String '使用者編號

  Private _CLIENT_ID As String '操作站編號

  Private _IP As String '操作站IP

  Private _CREATE_TIME As String '時間

  Private _MESSAGE As String '交易內容

  Private _RESULT As String '回覆結果未回覆=成功=0失敗=1回覆超時=2

  Private _RESULT_MESSAGE As String '回覆失敗的原因

  Private _WAIT_UUID As String '等待其他系統回覆的UUID，如果需要其他系統回覆則要填入，發送給其他系統的UUID，以便後續進行回傳後的連結

  Private _HIST_TIME As String '寫入歷史記錄的時間

  Private _objWMS As clsHandlingObject
  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property UUID() As String
    Get
      Return _UUID
    End Get
    Set(ByVal value As String)
      _UUID = value
    End Set
  End Property
  Public Property CONNECTION_TYPE() As Long
    Get
      Return _CONNECTION_TYPE
    End Get
    Set(ByVal value As Long)
      _CONNECTION_TYPE = value
    End Set
  End Property
  Public Property SEND_SYSTEM() As Long
    Get
      Return _SEND_SYSTEM
    End Get
    Set(ByVal value As Long)
      _SEND_SYSTEM = value
    End Set
  End Property
  Public Property FUNCTION_ID() As String
    Get
      Return _FUNCTION_ID
    End Get
    Set(ByVal value As String)
      _FUNCTION_ID = value
    End Set
  End Property
  Public Property SEQ() As Long
    Get
      Return _SEQ
    End Get
    Set(ByVal value As Long)
      _SEQ = value
    End Set
  End Property
  Public Property USER_ID() As String
    Get
      Return _USER_ID
    End Get
    Set(ByVal value As String)
      _USER_ID = value
    End Set
  End Property
  Public Property CLIENT_ID() As String
    Get
      Return _CLIENT_ID
    End Get
    Set(ByVal value As String)
      _CLIENT_ID = value
    End Set
  End Property
  Public Property IP() As String
    Get
      Return _IP
    End Get
    Set(ByVal value As String)
      _IP = value
    End Set
  End Property
  Public Property CREATE_TIME() As String
    Get
      Return _CREATE_TIME
    End Get
    Set(ByVal value As String)
      _CREATE_TIME = value
    End Set
  End Property
  Public Property MESSAGE() As String
    Get
      Return _MESSAGE
    End Get
    Set(ByVal value As String)
      _MESSAGE = value
    End Set
  End Property
  Public Property RESULT() As String
    Get
      Return _RESULT
    End Get
    Set(ByVal value As String)
      _RESULT = value
    End Set
  End Property
  Public Property RESULT_MESSAGE() As String
    Get
      Return _RESULT_MESSAGE
    End Get
    Set(ByVal value As String)
      _RESULT_MESSAGE = value
    End Set
  End Property
  Public Property WAIT_UUID() As String
    Get
      Return _WAIT_UUID
    End Get
    Set(ByVal value As String)
      _WAIT_UUID = value
    End Set
  End Property
  Public Property HIST_TIME() As String
    Get
      Return _HIST_TIME
    End Get
    Set(ByVal value As String)
      _HIST_TIME = value
    End Set
  End Property
  Public Property objWMS() As clsHandlingObject
    Get
      Return _objWMS
    End Get
    Set(ByVal value As clsHandlingObject)
      _objWMS = value
    End Set
  End Property

  Public Sub New(ByVal UUID As String, ByVal CONNECTION_TYPE As Long, ByVal SEND_SYSTEM As Long, ByVal FUNCTION_ID As String, ByVal SEQ As Long, ByVal USER_ID As String, ByVal CLIENT_ID As String, ByVal IP As String, ByVal CREATE_TIME As String, ByVal MESSAGE As String, ByVal RESULT As String, ByVal RESULT_MESSAGE As String, ByVal WAIT_UUID As String, ByVal HIST_TIME As String)
    MyBase.New()
    Try
      'Dim key As String = Get_Combination_Key(UUID)
      _gid = UUID
      _UUID = UUID
      _CONNECTION_TYPE = CONNECTION_TYPE
      _SEND_SYSTEM = SEND_SYSTEM
      _FUNCTION_ID = FUNCTION_ID
      _SEQ = SEQ
      _USER_ID = USER_ID
      _CLIENT_ID = CLIENT_ID
      _IP = IP
      _CREATE_TIME = CREATE_TIME
      _MESSAGE = MESSAGE
      _RESULT = RESULT
      _RESULT_MESSAGE = RESULT_MESSAGE
      _WAIT_UUID = WAIT_UUID
      _HIST_TIME = HIST_TIME
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '物件結束時觸發的事件，用來清除物件的內容
  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
  Private Sub Class_Terminate_Renamed()
    '目的:結束物件
  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Function Clone() As clsHSToHostCommand
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Sub Add_Relationship(ByRef objWMS As clsHandlingObject)
    'Try
    '  '挷定WMS的關係                                                                        
    '  If objWMS IsNot Nothing Then
    '    _objWMS = objWMS
    '    objWMS.O_Add_!!!!!這邊就是你要改的東西啦(Me)
    '  End If
    'Catch ex As Exception
    '  SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    'End Try
  End Sub
  Public Sub Remove_Relationship()
    'Try
    '  If _objWMS IsNot Nothing Then
    '    _objWMS.O_Remove_!!!!!這也是你要改的東西
    '  End If
    'Catch ex As Exception
    '  SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    'End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = HS_H_HOST_COMMANDManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = HS_H_HOST_COMMANDManagement.GetUpdateSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Delete的SQL
  Public Function O_Add_Delete_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = HS_H_HOST_COMMANDManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objHS_H_HOST_COMMAND As clsHSToHostCommand) As Boolean
    Try
      Dim key As String = objHS_H_HOST_COMMAND._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _UUID = objHS_H_HOST_COMMAND.UUID
      _CONNECTION_TYPE = objHS_H_HOST_COMMAND.CONNECTION_TYPE
      _SEND_SYSTEM = objHS_H_HOST_COMMAND.SEND_SYSTEM
      _FUNCTION_ID = objHS_H_HOST_COMMAND.FUNCTION_ID
      _SEQ = objHS_H_HOST_COMMAND.SEQ
      _USER_ID = objHS_H_HOST_COMMAND.USER_ID
      _CLIENT_ID = objHS_H_HOST_COMMAND.CLIENT_ID
      _IP = objHS_H_HOST_COMMAND.IP
      _CREATE_TIME = objHS_H_HOST_COMMAND.CREATE_TIME
      _MESSAGE = objHS_H_HOST_COMMAND.MESSAGE
      _RESULT = objHS_H_HOST_COMMAND.RESULT
      _RESULT_MESSAGE = objHS_H_HOST_COMMAND.RESULT_MESSAGE
      _WAIT_UUID = objHS_H_HOST_COMMAND.WAIT_UUID
      _HIST_TIME = objHS_H_HOST_COMMAND.HIST_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
