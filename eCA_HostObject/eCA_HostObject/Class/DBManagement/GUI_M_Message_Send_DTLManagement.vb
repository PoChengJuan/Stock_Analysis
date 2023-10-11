Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class GUI_M_Message_Send_DTLManagement
  Public Shared TableName As String = "GUI_M_Message_Send_DTL"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsGUI_M_Message_Send_DTL)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    KEY_NO
    SEND_TYPE
    SEND_USER_LIST
    SEND_GROUP_LIST
    SEND_ENABLE
  End Enum

  Public Enum UpdateOption As Integer
    UpdateDic = 0
    UpdateDB = 1
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsGUI_M_Message_Send_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10}) values ('{3}',{5},'{7}','{9}',{11})",
      strSQL,
      TableName,
      IdxColumnName.KEY_NO.ToString, CI.KEY_NO,
      IdxColumnName.SEND_TYPE.ToString, CI.SEND_TYPE,
      IdxColumnName.SEND_USER_LIST.ToString, CI.SEND_USER_LIST,
      IdxColumnName.SEND_GROUP_LIST.ToString, CI.SEND_GROUP_LIST,
      IdxColumnName.SEND_ENABLE.ToString, CI.SEND_ENABLE
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
  Public Shared Function GetDeleteSQL(ByRef CI As clsGUI_M_Message_Send_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}={5} ",
      strSQL,
      TableName,
      IdxColumnName.KEY_NO.ToString, CI.KEY_NO,
      IdxColumnName.SEND_TYPE.ToString, CI.SEND_TYPE,
      IdxColumnName.SEND_USER_LIST.ToString, CI.SEND_USER_LIST,
      IdxColumnName.SEND_GROUP_LIST.ToString, CI.SEND_GROUP_LIST,
      IdxColumnName.SEND_ENABLE.ToString, CI.SEND_ENABLE
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
  Public Shared Function GetUpdateSQL(ByRef CI As clsGUI_M_Message_Send_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}',{8}='{9}',{10}={11} WHERE {2}='{3}' And {4}={5}",
      strSQL,
      TableName,
      IdxColumnName.KEY_NO.ToString, CI.KEY_NO,
      IdxColumnName.SEND_TYPE.ToString, CI.SEND_TYPE,
      IdxColumnName.SEND_USER_LIST.ToString, CI.SEND_USER_LIST,
      IdxColumnName.SEND_GROUP_LIST.ToString, CI.SEND_GROUP_LIST,
      IdxColumnName.SEND_ENABLE.ToString, CI.SEND_ENABLE
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

  Public Shared Function GetGUI_M_Message_Send_DTLDataListByKey_KEY_NO_SEND_TYPE(ByVal key_no As String, send_type As Double) As Dictionary(Of String, clsGUI_M_Message_Send_DTL)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsGUI_M_Message_Send_DTL)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4} = {5} ",
            strSQL,
            TableName,
            IdxColumnName.KEY_NO.ToString, key_no,
            IdxColumnName.SEND_TYPE.ToString, send_type
            )
            Dim DatasetMessage As New DataSet
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsGUI_M_Message_Send_DTL = Nothing
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

  Public Shared Function GetGUI_M_Message_Send_DTLDataListBylstKEY_NO(ByVal lstkey_no As List(Of String)) As Dictionary(Of String, clsGUI_M_Message_Send_DTL)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsGUI_M_Message_Send_DTL)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim str_KeyNo As String = "'"
            Dim count = 0
            For i = count To lstkey_no.Count - 1
              str_KeyNo += lstkey_no(i) & "',"

            Next
            str_KeyNo = str_KeyNo.TrimEnd(",")
            strSQL = String.Format("Select * from {1} WHERE  {2} in ({3}) ",
            strSQL,
            TableName,
            IdxColumnName.KEY_NO.ToString, str_KeyNo
            )
            Dim DatasetMessage As New DataSet
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsGUI_M_Message_Send_DTL = Nothing
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
  Private Shared Function UpdateInfo(ByRef Key As String, ByRef Info As clsGUI_M_Message_Send_DTL, ByRef objNewTC As clsGUI_M_Message_Send_DTL) As clsGUI_M_Message_Send_DTL
    Try
      If Key = Info.gid Then
        Info.Update_To_Memory(objNewTC)
      Else
        SendMessageToLog("Dictionary has the different key", eCALogTool.ILogTool.enuTrcLevel.lvError)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
    Return Info
  End Function
  Private Shared Function SetInfoFromDB(ByRef Info As clsGUI_M_Message_Send_DTL, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim KEY_NO = "" & RowData.Item(IdxColumnName.KEY_NO.ToString)
        Dim SEND_TYPE = 0 & RowData.Item(IdxColumnName.SEND_TYPE.ToString)
        Dim SEND_USER_LIST = "" & RowData.Item(IdxColumnName.SEND_USER_LIST.ToString)
        Dim SEND_GROUP_LIST = "" & RowData.Item(IdxColumnName.SEND_GROUP_LIST.ToString)
        Dim SEND_ENABLE = 0 & RowData.Item(IdxColumnName.SEND_ENABLE.ToString)
        Info = New clsGUI_M_Message_Send_DTL(KEY_NO, SEND_TYPE, SEND_USER_LIST, SEND_GROUP_LIST, SEND_ENABLE)

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
