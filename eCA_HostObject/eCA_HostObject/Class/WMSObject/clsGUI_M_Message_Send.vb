Public Class clsGUI_M_Message_Send
  Private ShareName As String = "GUI_M_Message_Send"
  Private ShareKey As String = ""
  Private _gid As String
  Private _KEY_NO As String '流水號

  Private _MESSAGE_TYPE As Double '發送事件類型(程式自定義)

  Private _MESSAGE_TYPE_DESC As String '發送事件說明

  Private _CONDITION1 As String '自定義條件(不同類型的情況，定義的條件不一樣)

  Private _CONDITION2 As String '自定義條件

  Private _CONDITION3 As String '自定義條件

  Private _CONDITION4 As String '自定義條件

  Private _CONDITION5 As String '自定義條件

  Private _ENABLE As Double '是否發送…0:不發送1:發送

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
  Public Property MESSAGE_TYPE() As Double
    Get
      Return _MESSAGE_TYPE
    End Get
    Set(ByVal value As Double)
      _MESSAGE_TYPE = value
    End Set
  End Property
  Public Property MESSAGE_TYPE_DESC() As String
    Get
      Return _MESSAGE_TYPE_DESC
    End Get
    Set(ByVal value As String)
      _MESSAGE_TYPE_DESC = value
    End Set
  End Property
  Public Property CONDITION1() As String
    Get
      Return _CONDITION1
    End Get
    Set(ByVal value As String)
      _CONDITION1 = value
    End Set
  End Property
  Public Property CONDITION2() As String
    Get
      Return _CONDITION2
    End Get
    Set(ByVal value As String)
      _CONDITION2 = value
    End Set
  End Property
  Public Property CONDITION3() As String
    Get
      Return _CONDITION3
    End Get
    Set(ByVal value As String)
      _CONDITION3 = value
    End Set
  End Property
  Public Property CONDITION4() As String
    Get
      Return _CONDITION4
    End Get
    Set(ByVal value As String)
      _CONDITION4 = value
    End Set
  End Property
  Public Property CONDITION5() As String
    Get
      Return _CONDITION5
    End Get
    Set(ByVal value As String)
      _CONDITION5 = value
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

  Public Sub New(ByVal KEY_NO As String, ByVal MESSAGE_TYPE As Double, ByVal MESSAGE_TYPE_DESC As String, ByVal CONDITION1 As String, ByVal CONDITION2 As String, ByVal CONDITION3 As String, ByVal CONDITION4 As String, ByVal CONDITION5 As String, ByVal ENABLE As Double)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(KEY_NO)
      _gid = key
      _KEY_NO = KEY_NO
      _MESSAGE_TYPE = MESSAGE_TYPE
      _MESSAGE_TYPE_DESC = MESSAGE_TYPE_DESC
      _CONDITION1 = CONDITION1
      _CONDITION2 = CONDITION2
      _CONDITION3 = CONDITION3
      _CONDITION4 = CONDITION4
      _CONDITION5 = CONDITION5
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
  Public Shared Function Get_Combination_Key(ByVal KEY_NO As String) As String
    Try
      Dim key As String = KEY_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsGUI_M_Message_Send
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
      Dim strSQL As String = GUI_M_Message_SendManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = GUI_M_Message_SendManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = GUI_M_Message_SendManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objGUI_M_Message_Send As clsGUI_M_Message_Send) As Boolean
    Try
      Dim key As String = objGUI_M_Message_Send._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _KEY_NO = KEY_NO
      _MESSAGE_TYPE = MESSAGE_TYPE
      _MESSAGE_TYPE_DESC = MESSAGE_TYPE_DESC
      _CONDITION1 = CONDITION1
      _CONDITION2 = CONDITION2
      _CONDITION3 = CONDITION3
      _CONDITION4 = CONDITION4
      _CONDITION5 = CONDITION5
      _ENABLE = ENABLE
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
