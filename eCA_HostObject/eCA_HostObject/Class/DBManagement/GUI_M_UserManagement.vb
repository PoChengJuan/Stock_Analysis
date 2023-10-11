Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class GUI_M_UserManagement
  Public Shared TableName As String = "GUI_M_User"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsGUI_M_User)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    USER_ID
    USER_LAST_NAME
    USER_FIRST_NAME
    USER_NICK_NAME
    GROUP_ID
    ROLE_ID
    PASSWORD
    FACTORY_LIST
    MAIL
    ENABLE
    PASSWORD_UPDATE_TIME
    LANGUAGE
    CONTACT_LINE
    CONTACT_WECHAT
    CONTACT_PHONE
    CONTACT_MAIL
    CONTACT_COMMON_1
    CONTACT_COMMON_2
    COMMENTS
    PASSWORD_TRY_COUNT
    UNLOCK_TIME
  End Enum

  Public Enum UpdateOption As Integer
    UpdateDic = 0
    UpdateDB = 1
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsGUI_M_User) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}',{21},'{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}')",
      strSQL,
      TableName,
      IdxColumnName.USER_ID.ToString, CI.USER_ID,
      IdxColumnName.USER_LAST_NAME.ToString, CI.USER_LAST_NAME,
      IdxColumnName.USER_FIRST_NAME.ToString, CI.USER_FIRST_NAME,
      IdxColumnName.USER_NICK_NAME.ToString, CI.USER_NICK_NAME,
      IdxColumnName.GROUP_ID.ToString, CI.GROUP_ID,
      IdxColumnName.ROLE_ID.ToString, CI.ROLE_ID,
      IdxColumnName.PASSWORD.ToString, CI.PASSWORD,
      IdxColumnName.FACTORY_LIST.ToString, CI.FACTORY_LIST,
      IdxColumnName.MAIL.ToString, CI.MAIL,
      IdxColumnName.ENABLE.ToString, CI.ENABLE,
      IdxColumnName.PASSWORD_UPDATE_TIME.ToString, CI.PASSWORD_UPDATE_TIME,
      IdxColumnName.LANGUAGE.ToString, CI.LANGUAGE,
      IdxColumnName.CONTACT_LINE.ToString, CI.CONTACT_LINE,
      IdxColumnName.CONTACT_WECHAT.ToString, CI.CONTACT_WECHAT,
      IdxColumnName.CONTACT_PHONE.ToString, CI.CONTACT_PHONE,
      IdxColumnName.CONTACT_MAIL.ToString, CI.CONTACT_MAIL,
      IdxColumnName.CONTACT_COMMON_1.ToString, CI.CONTACT_COMMON_1,
      IdxColumnName.CONTACT_COMMON_2.ToString, CI.CONTACT_COMMON_2,
      IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
      IdxColumnName.PASSWORD_TRY_COUNT.ToString, CI.PASSWORD_TRY_COUNT,
      IdxColumnName.UNLOCK_TIME.ToString, CI.UNLOCK_TIME
      )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef CI As clsGUI_M_User) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.USER_ID.ToString, CI.USER_ID,
      IdxColumnName.USER_LAST_NAME.ToString, CI.USER_LAST_NAME,
      IdxColumnName.USER_FIRST_NAME.ToString, CI.USER_FIRST_NAME,
      IdxColumnName.USER_NICK_NAME.ToString, CI.USER_NICK_NAME,
      IdxColumnName.GROUP_ID.ToString, CI.GROUP_ID,
      IdxColumnName.ROLE_ID.ToString, CI.ROLE_ID,
      IdxColumnName.PASSWORD.ToString, CI.PASSWORD,
      IdxColumnName.FACTORY_LIST.ToString, CI.FACTORY_LIST,
      IdxColumnName.MAIL.ToString, CI.MAIL,
      IdxColumnName.ENABLE.ToString, CI.ENABLE,
      IdxColumnName.PASSWORD_UPDATE_TIME.ToString, CI.PASSWORD_UPDATE_TIME,
      IdxColumnName.LANGUAGE.ToString, CI.LANGUAGE,
      IdxColumnName.CONTACT_LINE.ToString, CI.CONTACT_LINE,
      IdxColumnName.CONTACT_WECHAT.ToString, CI.CONTACT_WECHAT,
      IdxColumnName.CONTACT_PHONE.ToString, CI.CONTACT_PHONE,
      IdxColumnName.CONTACT_MAIL.ToString, CI.CONTACT_MAIL,
      IdxColumnName.CONTACT_COMMON_1.ToString, CI.CONTACT_COMMON_1,
      IdxColumnName.CONTACT_COMMON_2.ToString, CI.CONTACT_COMMON_2,
      IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
      IdxColumnName.PASSWORD_TRY_COUNT.ToString, CI.PASSWORD_TRY_COUNT,
      IdxColumnName.UNLOCK_TIME.ToString, CI.UNLOCK_TIME
      )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef CI As clsGUI_M_User) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}={21},{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.USER_ID.ToString, CI.USER_ID,
      IdxColumnName.USER_LAST_NAME.ToString, CI.USER_LAST_NAME,
      IdxColumnName.USER_FIRST_NAME.ToString, CI.USER_FIRST_NAME,
      IdxColumnName.USER_NICK_NAME.ToString, CI.USER_NICK_NAME,
      IdxColumnName.GROUP_ID.ToString, CI.GROUP_ID,
      IdxColumnName.ROLE_ID.ToString, CI.ROLE_ID,
      IdxColumnName.PASSWORD.ToString, CI.PASSWORD,
      IdxColumnName.FACTORY_LIST.ToString, CI.FACTORY_LIST,
      IdxColumnName.MAIL.ToString, CI.MAIL,
      IdxColumnName.ENABLE.ToString, CI.ENABLE,
      IdxColumnName.PASSWORD_UPDATE_TIME.ToString, CI.PASSWORD_UPDATE_TIME,
      IdxColumnName.LANGUAGE.ToString, CI.LANGUAGE,
      IdxColumnName.CONTACT_LINE.ToString, CI.CONTACT_LINE,
      IdxColumnName.CONTACT_WECHAT.ToString, CI.CONTACT_WECHAT,
      IdxColumnName.CONTACT_PHONE.ToString, CI.CONTACT_PHONE,
      IdxColumnName.CONTACT_MAIL.ToString, CI.CONTACT_MAIL,
      IdxColumnName.CONTACT_COMMON_1.ToString, CI.CONTACT_COMMON_1,
      IdxColumnName.CONTACT_COMMON_2.ToString, CI.CONTACT_COMMON_2,
      IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
      IdxColumnName.PASSWORD_TRY_COUNT.ToString, CI.PASSWORD_TRY_COUNT,
      IdxColumnName.UNLOCK_TIME.ToString, CI.UNLOCK_TIME
      )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetGUI_M_UserDataListByKey_USER_ID(ByVal user_id As String) As Dictionary(Of String, clsGUI_M_User)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsGUI_M_User)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
            strSQL,
            TableName,
            IdxColumnName.USER_ID.ToString, user_id
            )
            Dim DatasetMessage As New DataSet
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsGUI_M_User = Nothing
                SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                If _lstReturn.ContainsKey(Info.gid) = False Then
                  _lstReturn.Add(Info.gid, Info)
                End If
              Next
            End If
          End If
        End If
        Return _lstReturn
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function
  Public Shared Function GetGUI_M_UserDataListByALL() As Dictionary(Of String, clsGUI_M_User)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsGUI_M_User)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty


            strSQL = String.Format("Select * from {1} ",
            strSQL,
            TableName
            )
            Dim DatasetMessage As New DataSet
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsGUI_M_User = Nothing
                SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                If _lstReturn.ContainsKey(Info.gid) = False Then
                  _lstReturn.Add(Info.gid, Info)
                End If
              Next
            End If
          End If
        End If
        Return _lstReturn
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function

  Private Shared Function SetInfoFromDB(ByRef Info As clsGUI_M_User, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim USER_ID = "" & RowData.Item(IdxColumnName.USER_ID.ToString)
        Dim USER_LAST_NAME = "" & RowData.Item(IdxColumnName.USER_LAST_NAME.ToString)
        Dim USER_FIRST_NAME = "" & RowData.Item(IdxColumnName.USER_FIRST_NAME.ToString)
        Dim USER_NICK_NAME = "" & RowData.Item(IdxColumnName.USER_NICK_NAME.ToString)
        Dim GROUP_ID = "" & RowData.Item(IdxColumnName.GROUP_ID.ToString)
        Dim ROLE_ID = "" & RowData.Item(IdxColumnName.ROLE_ID.ToString)
        Dim PASSWORD = "" & RowData.Item(IdxColumnName.PASSWORD.ToString)
        Dim FACTORY_LIST = "" & RowData.Item(IdxColumnName.FACTORY_LIST.ToString)
        Dim MAIL = "" & RowData.Item(IdxColumnName.MAIL.ToString)
        Dim ENABLE = 0 & RowData.Item(IdxColumnName.ENABLE.ToString)
        Dim PASSWORD_UPDATE_TIME = "" & RowData.Item(IdxColumnName.PASSWORD_UPDATE_TIME.ToString)
        Dim LANGUAGE = "" & RowData.Item(IdxColumnName.LANGUAGE.ToString)
        Dim CONTACT_LINE = "" & RowData.Item(IdxColumnName.CONTACT_LINE.ToString)
        Dim CONTACT_WECHAT = "" & RowData.Item(IdxColumnName.CONTACT_WECHAT.ToString)
        Dim CONTACT_PHONE = "" & RowData.Item(IdxColumnName.CONTACT_PHONE.ToString)
        Dim CONTACT_MAIL = "" & RowData.Item(IdxColumnName.CONTACT_MAIL.ToString)
        Dim CONTACT_COMMON_1 = "" & RowData.Item(IdxColumnName.CONTACT_COMMON_1.ToString)
        Dim CONTACT_COMMON_2 = "" & RowData.Item(IdxColumnName.CONTACT_COMMON_2.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim PASSWORD_TRY_COUNT = "" & RowData.Item(IdxColumnName.PASSWORD_TRY_COUNT.ToString)
        Dim UNLOCK_TIME = "" & RowData.Item(IdxColumnName.UNLOCK_TIME.ToString)
        Info = New clsGUI_M_User(USER_ID, USER_LAST_NAME, USER_FIRST_NAME, USER_NICK_NAME, GROUP_ID, ROLE_ID, PASSWORD, FACTORY_LIST, MAIL, ENABLE, PASSWORD_UPDATE_TIME, LANGUAGE, CONTACT_LINE, CONTACT_WECHAT, CONTACT_PHONE, CONTACT_MAIL, CONTACT_COMMON_1, CONTACT_COMMON_2, COMMENTS, PASSWORD_TRY_COUNT, UNLOCK_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function SendSQLToDB(ByRef lstSQL As List(Of String)) As Boolean
    Try
      If lstSQL Is Nothing Then Return False
      If lstSQL.Count = 0 Then Return True
      For i = 0 To lstSQL.Count - 1
        SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next
      If fUseBatchUpdate_DynamicConnection = 0 Then
        For i = 0 To lstSQL.Count - 1
          DBTool.O_AddSQLQueue(TableName, lstSQL(i))
        Next
      Else
        Dim rtnMsg As String = DBTool.BatchUpdate_DynamicConnection(lstSQL)
        If rtnMsg.StartsWith("OK") Then
          SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        End If
      End If
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
End Class
