Public Class clsGUI_M_Message_Send_DTL
  Private ShareName As String = "GUI_M_Message_Send_DTL"
  Private ShareKey As String = ""
  Private _gid As String
  Private _KEY_NO As String '對應主表的Key值

  Private _SEND_TYPE As Double '發送類型MAIL=1LINE=2WECHAT=3簡訊=4DB Message=5

  Private _SEND_USER_LIST As String '發送的使用者列表(多筆資訊是用,隔開)(儲存使用者編號)

  Private _SEND_GROUP_LIST As String '發送的群組列表(多筆資訊是用,隔開)(儲存群組編號)

  Private _SEND_ENABLE As Double '是否發送…0:不發送1:發送

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property KEY_NO() As String
    Get
      Return _KEY_NO
    End Get
    Set(ByVal value As String)
      _KEY_NO = value
    End Set
  End Property
  Public Property SEND_TYPE() As Double
    Get
      Return _SEND_TYPE
    End Get
    Set(ByVal value As Double)
      _SEND_TYPE = value
    End Set
  End Property
  Public Property SEND_USER_LIST() As String
    Get
      Return _SEND_USER_LIST
    End Get
    Set(ByVal value As String)
      _SEND_USER_LIST = value
    End Set
  End Property
  Public Property SEND_GROUP_LIST() As String
    Get
      Return _SEND_GROUP_LIST
    End Get
    Set(ByVal value As String)
      _SEND_GROUP_LIST = value
    End Set
  End Property
  Public Property SEND_ENABLE() As Double
    Get
      Return _SEND_ENABLE
    End Get
    Set(ByVal value As Double)
      _SEND_ENABLE = value
    End Set
  End Property

  Public Sub New(ByVal KEY_NO As String, ByVal SEND_TYPE As Double, ByVal SEND_USER_LIST As String, ByVal SEND_GROUP_LIST As String, ByVal SEND_ENABLE As Double)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(KEY_NO, SEND_TYPE)
      _gid = key
      _KEY_NO = KEY_NO
      _SEND_TYPE = SEND_TYPE
      _SEND_USER_LIST = SEND_USER_LIST
      _SEND_GROUP_LIST = SEND_GROUP_LIST
      _SEND_ENABLE = SEND_ENABLE
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
  Public Shared Function Get_Combination_Key(ByVal KEY_NO As String, ByVal SEND_TYPE As Double) As String
    Try
      Dim key As String = KEY_NO & LinkKey & SEND_TYPE
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsGUI_M_Message_Send_DTL
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = GUI_M_Message_Send_DTLManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = GUI_M_Message_Send_DTLManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = GUI_M_Message_Send_DTLManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objGUI_M_Message_Send_DTL As clsGUI_M_Message_Send_DTL) As Boolean
    Try
      Dim key As String = objGUI_M_Message_Send_DTL._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _KEY_NO = KEY_NO
      _SEND_TYPE = SEND_TYPE
      _SEND_USER_LIST = SEND_USER_LIST
      _SEND_GROUP_LIST = SEND_GROUP_LIST
      _SEND_ENABLE = SEND_ENABLE
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
