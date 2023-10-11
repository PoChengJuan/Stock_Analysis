Public Class clsGUI_M_User
  Private ShareName As String = "GUI_M_User"
  Private ShareKey As String = ""
  Private _gid As String
  Private _USER_ID As String '使用者帳號(登入帳號)

  Private _USER_LAST_NAME As String '使用者姓氏

  Private _USER_FIRST_NAME As String '使用者名字

  Private _USER_NICK_NAME As String '使用者暱稱

  Private _GROUP_ID As String '群組

  Private _ROLE_ID As String '角色編號

  Private _PASSWORD As String '密碼

  Private _FACTORY_LIST As String '可登入廠區

  Private _MAIL As String '郵箱

  Private _ENABLE As Double '是否啟用(1:啟用/0:禁用)

  Private _PASSWORD_UPDATE_TIME As String '密碼更新時間

  Private _LANGUAGE As String '使用者預設語言

  Private _CONTACT_LINE As String '聯絡_LINE

  Private _CONTACT_WECHAT As String '聯絡_唯信

  Private _CONTACT_PHONE As String '聯絡_手機

  Private _CONTACT_MAIL As String '聯絡_郵件

  Private _CONTACT_COMMON_1 As String '聯絡_預留

  Private _CONTACT_COMMON_2 As String '聯絡_預留

  Private _COMMENTS As String '備註

  Private _PASSWORD_TRY_COUNT As String '密碼嘗試次數

  Private _UNLOCK_TIME As String '解鎖時間

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
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
  Public Property USER_LAST_NAME() As String
    Get
      Return _USER_LAST_NAME
    End Get
    Set(ByVal value As String)
      _USER_LAST_NAME = value
    End Set
  End Property
  Public Property USER_FIRST_NAME() As String
    Get
      Return _USER_FIRST_NAME
    End Get
    Set(ByVal value As String)
      _USER_FIRST_NAME = value
    End Set
  End Property
  Public Property USER_NICK_NAME() As String
    Get
      Return _USER_NICK_NAME
    End Get
    Set(ByVal value As String)
      _USER_NICK_NAME = value
    End Set
  End Property
  Public Property GROUP_ID() As String
    Get
      Return _GROUP_ID
    End Get
    Set(ByVal value As String)
      _GROUP_ID = value
    End Set
  End Property
  Public Property ROLE_ID() As String
    Get
      Return _ROLE_ID
    End Get
    Set(ByVal value As String)
      _ROLE_ID = value
    End Set
  End Property
  Public Property PASSWORD() As String
    Get
      Return _PASSWORD
    End Get
    Set(ByVal value As String)
      _PASSWORD = value
    End Set
  End Property
  Public Property FACTORY_LIST() As String
    Get
      Return _FACTORY_LIST
    End Get
    Set(ByVal value As String)
      _FACTORY_LIST = value
    End Set
  End Property
  Public Property MAIL() As String
    Get
      Return _MAIL
    End Get
    Set(ByVal value As String)
      _MAIL = value
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
  Public Property PASSWORD_UPDATE_TIME() As String
    Get
      Return _PASSWORD_UPDATE_TIME
    End Get
    Set(ByVal value As String)
      _PASSWORD_UPDATE_TIME = value
    End Set
  End Property
  Public Property LANGUAGE() As String
    Get
      Return _LANGUAGE
    End Get
    Set(ByVal value As String)
      _LANGUAGE = value
    End Set
  End Property
  Public Property CONTACT_LINE() As String
    Get
      Return _CONTACT_LINE
    End Get
    Set(ByVal value As String)
      _CONTACT_LINE = value
    End Set
  End Property
  Public Property CONTACT_WECHAT() As String
    Get
      Return _CONTACT_WECHAT
    End Get
    Set(ByVal value As String)
      _CONTACT_WECHAT = value
    End Set
  End Property
  Public Property CONTACT_PHONE() As String
    Get
      Return _CONTACT_PHONE
    End Get
    Set(ByVal value As String)
      _CONTACT_PHONE = value
    End Set
  End Property
  Public Property CONTACT_MAIL() As String
    Get
      Return _CONTACT_MAIL
    End Get
    Set(ByVal value As String)
      _CONTACT_MAIL = value
    End Set
  End Property
  Public Property CONTACT_COMMON_1() As String
    Get
      Return _CONTACT_COMMON_1
    End Get
    Set(ByVal value As String)
      _CONTACT_COMMON_1 = value
    End Set
  End Property
  Public Property CONTACT_COMMON_2() As String
    Get
      Return _CONTACT_COMMON_2
    End Get
    Set(ByVal value As String)
      _CONTACT_COMMON_2 = value
    End Set
  End Property
  Public Property COMMENTS() As String
    Get
      Return _COMMENTS
    End Get
    Set(ByVal value As String)
      _COMMENTS = value
    End Set
  End Property
  Public Property PASSWORD_TRY_COUNT() As String
    Get
      Return _PASSWORD_TRY_COUNT
    End Get
    Set(ByVal value As String)
      _PASSWORD_TRY_COUNT = value
    End Set
  End Property
  Public Property UNLOCK_TIME() As String
    Get
      Return _UNLOCK_TIME
    End Get
    Set(ByVal value As String)
      _UNLOCK_TIME = value
    End Set
  End Property

  Public Sub New(ByVal USER_ID As String, ByVal USER_LAST_NAME As String, ByVal USER_FIRST_NAME As String, ByVal USER_NICK_NAME As String, ByVal GROUP_ID As String, ByVal ROLE_ID As String, ByVal PASSWORD As String, ByVal FACTORY_LIST As String, ByVal MAIL As String, ByVal ENABLE As Double, ByVal PASSWORD_UPDATE_TIME As String, ByVal LANGUAGE As String, ByVal CONTACT_LINE As String, ByVal CONTACT_WECHAT As String, ByVal CONTACT_PHONE As String, ByVal CONTACT_MAIL As String, ByVal CONTACT_COMMON_1 As String, ByVal CONTACT_COMMON_2 As String, ByVal COMMENTS As String, ByVal PASSWORD_TRY_COUNT As String, ByVal UNLOCK_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(USER_ID)
      _gid = key
      _USER_ID = USER_ID
      _USER_LAST_NAME = USER_LAST_NAME
      _USER_FIRST_NAME = USER_FIRST_NAME
      _USER_NICK_NAME = USER_NICK_NAME
      _GROUP_ID = GROUP_ID
      _ROLE_ID = ROLE_ID
      _PASSWORD = PASSWORD
      _FACTORY_LIST = FACTORY_LIST
      _MAIL = MAIL
      _ENABLE = ENABLE
      _PASSWORD_UPDATE_TIME = PASSWORD_UPDATE_TIME
      _LANGUAGE = LANGUAGE
      _CONTACT_LINE = CONTACT_LINE
      _CONTACT_WECHAT = CONTACT_WECHAT
      _CONTACT_PHONE = CONTACT_PHONE
      _CONTACT_MAIL = CONTACT_MAIL
      _CONTACT_COMMON_1 = CONTACT_COMMON_1
      _CONTACT_COMMON_2 = CONTACT_COMMON_2
      _COMMENTS = COMMENTS
      _PASSWORD_TRY_COUNT = PASSWORD_TRY_COUNT
      _UNLOCK_TIME = UNLOCK_TIME
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
  Public Shared Function Get_Combination_Key(ByVal USER_ID As String) As String
    Try
      Dim key As String = USER_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsGUI_M_User
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
      Dim strSQL As String = GUI_M_UserManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = GUI_M_UserManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = GUI_M_UserManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objGUI_M_User As clsGUI_M_User) As Boolean
    Try
      Dim key As String = objGUI_M_User._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _USER_ID = USER_ID
      _USER_LAST_NAME = USER_LAST_NAME
      _USER_FIRST_NAME = USER_FIRST_NAME
      _USER_NICK_NAME = USER_NICK_NAME
      _GROUP_ID = GROUP_ID
      _ROLE_ID = ROLE_ID
      _PASSWORD = PASSWORD
      _FACTORY_LIST = FACTORY_LIST
      _MAIL = MAIL
      _ENABLE = ENABLE
      _PASSWORD_UPDATE_TIME = PASSWORD_UPDATE_TIME
      _LANGUAGE = LANGUAGE
      _CONTACT_LINE = CONTACT_LINE
      _CONTACT_WECHAT = CONTACT_WECHAT
      _CONTACT_PHONE = CONTACT_PHONE
      _CONTACT_MAIL = CONTACT_MAIL
      _CONTACT_COMMON_1 = CONTACT_COMMON_1
      _CONTACT_COMMON_2 = CONTACT_COMMON_2
      _COMMENTS = COMMENTS
      _PASSWORD_TRY_COUNT = PASSWORD_TRY_COUNT
      _UNLOCK_TIME = UNLOCK_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
