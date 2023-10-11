Public Class clsGUI_M_Message_Type
  Private ShareName As String = "GUI_M_Message_Type"
  Private ShareKey As String = ""
  Private _gid As String
  Private _MESSAGE_TYPE As Double '發送事件類型(程式自定義)

  Private _MESSAGE_TYPE_ALIS As String '發送事件說明

  Private _ENABLE As Double '是否啟用0:不啟用1:啟用

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property MESSAGE_TYPE() As Double
    Get
      Return _MESSAGE_TYPE
    End Get
    Set(ByVal value As Double)
      _MESSAGE_TYPE = value
    End Set
  End Property
  Public Property MESSAGE_TYPE_ALIS() As String
    Get
      Return _MESSAGE_TYPE_ALIS
    End Get
    Set(ByVal value As String)
      _MESSAGE_TYPE_ALIS = value
    End Set
  End Property
  Public Property ENABLE() As Double
    Get
      Return _ENABLE
    End Get
    Set(ByVal value As Double)
      _ENABLE = value
    End Set
  End Property

  Public Sub New(ByVal MESSAGE_TYPE As Double, ByVal MESSAGE_TYPE_DESC As String, ByVal ENABLE As Double)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(MESSAGE_TYPE)
      _gid = key
      _MESSAGE_TYPE = MESSAGE_TYPE
      _MESSAGE_TYPE_ALIS = MESSAGE_TYPE_DESC
      _ENABLE = ENABLE
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
  Public Shared Function Get_Combination_Key(ByVal MESSAGE_TYPE As Double) As String
    Try
      Dim key As String = MESSAGE_TYPE
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsGUI_M_Message_Type
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
      Dim strSQL As String = GUI_M_Message_TypeManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = GUI_M_Message_TypeManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = GUI_M_Message_TypeManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objGUI_M_Message_Type As clsGUI_M_Message_Type) As Boolean
    Try
      Dim key As String = objGUI_M_Message_Type._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _MESSAGE_TYPE = MESSAGE_TYPE
      _MESSAGE_TYPE_ALIS = MESSAGE_TYPE_ALIS
      _ENABLE = ENABLE
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
