Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_M_UUIDManagement
  Public Shared TableName As String = "WMS_M_UUID"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsUUID)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    UUID_NO
    UUID_SEQ
    IDLENGTH
    APPEND
    COMMENTS
    RESETABLE
    UPDATE_DATE
  End Enum

  '- GetSQL
  '-請將 clsUUID 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsUUID) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14}) values ('{3}',{5},{7},'{9}','{11}','{13}','{15}')",
      strSQL,
      TableName,
      IdxColumnName.UUID_NO.ToString, Info.get_UUID_NO,
      IdxColumnName.UUID_SEQ.ToString, Info.get_UUID_SEQ,
      IdxColumnName.IDLENGTH.ToString, Info.get_IDLENGTH,
      IdxColumnName.APPEND.ToString, Info.get_APPEND,
      IdxColumnName.COMMENTS.ToString, Info.get_COMMENTS,
      IdxColumnName.RESETABLE.ToString, Info.get_RESETABLE,
      IdxColumnName.UPDATE_DATE.ToString, Info.get_UPDATE_DATE
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsUUID) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.UUID_NO.ToString, Info.get_UUID_NO,
      IdxColumnName.UUID_SEQ.ToString, Info.get_UUID_SEQ,
      IdxColumnName.IDLENGTH.ToString, Info.get_IDLENGTH,
      IdxColumnName.APPEND.ToString, Info.get_APPEND,
      IdxColumnName.COMMENTS.ToString, Info.get_COMMENTS,
      IdxColumnName.RESETABLE.ToString, Info.get_RESETABLE,
      IdxColumnName.UPDATE_DATE.ToString, Info.get_UPDATE_DATE
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsUUID) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}={5},{6}={7},{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.UUID_NO.ToString, Info.get_UUID_NO,
      IdxColumnName.UUID_SEQ.ToString, Info.get_UUID_SEQ,
      IdxColumnName.IDLENGTH.ToString, Info.get_IDLENGTH,
      IdxColumnName.APPEND.ToString, Info.get_APPEND,
      IdxColumnName.COMMENTS.ToString, Info.get_COMMENTS,
      IdxColumnName.RESETABLE.ToString, Info.get_RESETABLE,
      IdxColumnName.UPDATE_DATE.ToString, Info.get_UPDATE_DATE
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

  '- Add & Insert
  Public Shared Function AddWMS_M_UUIDData(ByVal Info As clsUUID, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If AddlstWMS_M_UUIDData(New List(Of clsUUID)({Info}), SendToDB) = True Then
          Return True
        End If '-載不載入記憶體都是呼叫同一個function
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function AddlstWMS_M_UUIDData(ByVal Info As List(Of clsUUID), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True

        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).get_gid()
            If dicData.ContainsKey(key) = True Then
              SendMessageToLog("Add the same key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next

          If SendToDB Then
            If InsertWMS_M_UUIDDataToDB(Info) Then
              If AddOrUpdateWMS_M_UUIDDataToDictionary(Info) Then
                SendMessageToLog("InsertDic WMS_M_UUIDData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("InsertDic WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("InsertDB WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If AddOrUpdateWMS_M_UUIDDataToDictionary(Info) Then
              SendMessageToLog("InsertDic WMS_M_UUIDData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("InsertDic WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If InsertWMS_M_UUIDDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("InsertDic WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function

  '- Update
  Public Shared Function UpdateWMS_M_UUIDData(ByVal Info As clsUUID, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If UpdatelstWMS_M_UUIDData(New List(Of clsUUID)({Info}), SendToDB) = True Then
          Return True
        End If
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function UpdatelstWMS_M_UUIDData(ByVal Info As List(Of clsUUID), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True

        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).get_gid()
            If dicData.ContainsKey(key) = False Then
              SendMessageToLog("There is no key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next

          If SendToDB Then
            If UpdateWMS_M_UUIDDataToDB(Info) Then
              If AddOrUpdateWMS_M_UUIDDataToDictionary(Info) Then
                SendMessageToLog("UpdateDic WMS_M_UUIDData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("UpdateDic WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("UpdateDB WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If AddOrUpdateWMS_M_UUIDDataToDictionary(Info) Then
              SendMessageToLog("UpdateDic WMS_M_UUIDData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("UpdateDic WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If UpdateWMS_M_UUIDDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("UpdateDB WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function

  '- Delete
  Public Shared Function DeleteWMS_M_UUIDData(ByVal Info As clsUUID, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If DeletelstWMS_M_UUIDData(New List(Of clsUUID)({Info}), SendToDB) = True Then
          Return True
        End If '-載不載入記憶體都是呼叫同一個function
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function DeletelstWMS_M_UUIDData(ByVal Info As List(Of clsUUID), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True

        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).get_gid()
            If dicData.ContainsKey(key) = False Then
              SendMessageToLog("There is no key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next

          If SendToDB Then
            If DeleteWMS_M_UUIDDataToDB(Info) Then
              If DeleteWMS_M_UUIDDataToDictionary(Info) Then
                SendMessageToLog("DeleteDic WMS_M_UUIDData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("DeleteDic WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("DeleteDB WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If DeleteWMS_M_UUIDDataToDB(Info) Then
              SendMessageToLog("DeleteDic WMS_M_UUIDData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("DeleteDB WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
          Return True
        Else
          If SendToDB Then
            If DeleteWMS_M_UUIDDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("DeleteDB WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function

  '- GET
  Public Shared Function GetWMS_M_UUIDDataListByALL() As List(Of clsUUID)
    SyncLock objLock
      Try
        Dim _lstReturn As New List(Of clsUUID)
        If DictionaryNeeded = 1 Then '-載入記憶體
          Dim LinqFind As IEnumerable(Of clsUUID) = From TC In dicData Select TC.Value
          '- From TC In dicData Where TC.Value.xxx = xxx AND TC.Value.xxx = xxx AND TC.Value.xxx = xxx Select TC.Value '-範例
          For Each objTC As clsUUID In LinqFind
            _lstReturn.Add(objTC)
          Next

          Return _lstReturn
        Else
          If DBTool IsNot Nothing Then
            If DBTool.isConnection(DBTool.m_CN) = True Then
              Dim strSQL As String = String.Empty
              Dim DatasetMessage As New DataSet

              strSQL = String.Format("Select * from {0}", TableName)
              SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
              DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)


              If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
                For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                  Dim Info As clsUUID = Nothing
                  SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                  _lstReturn.Add(Info)
                Next
              End If
            End If
          End If
          Return _lstReturn
        End If
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function
  Public Shared Function GetclsUUIDListByUUID_NO(ByVal uuid_no As String) As Dictionary(Of String, clsUUID)
    SyncLock objLock
      Try
        Dim ret_dic As New Dictionary(Of String, clsUUID)

        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
            strSQL,
            TableName,
            IdxColumnName.UUID_NO.ToString, uuid_no
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)


            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsUUID = Nothing
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) Then
                  If Info IsNot Nothing Then
                    If ret_dic.ContainsKey(Info.get_gid) = False Then
                      ret_dic.Add(Info.get_gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsUUID Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsUUID Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Next
            End If
          End If
          Return ret_dic
        End If
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function

  '-Function


  '-以下為內部私人用
  Private Shared Function InsertWMS_M_UUIDDataToDB(ByRef Info As List(Of clsUUID)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14}) values ('{3}',{5},{7},'{9}','{11}','{13}','{15}')",
        strSQL,
        TableName,
        IdxColumnName.UUID_NO.ToString, CI.get_UUID_NO,
        IdxColumnName.UUID_SEQ.ToString, CI.get_UUID_SEQ,
        IdxColumnName.IDLENGTH.ToString, CI.get_IDLENGTH,
        IdxColumnName.APPEND.ToString, CI.get_APPEND,
        IdxColumnName.COMMENTS.ToString, CI.get_COMMENTS,
        IdxColumnName.RESETABLE.ToString, CI.get_RESETABLE,
        IdxColumnName.UPDATE_DATE.ToString, CI.get_UPDATE_DATE
        )
        lstSql.Add(strSQL)
      Next
      Dim NewSQL As New List(Of String)
      If SQLCorrect(lstSql, NewSQL) = False Then
        Return Nothing
      End If
      If SendSQLToDB(NewSQL) = True Then
        Return True
      Else
        SendMessageToLog("Insert to WMS_M_UUIDData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateWMS_M_UUIDDataToDB(ByRef Info As List(Of clsUUID)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Update {1} SET {4}={5},{6}={7},{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}' WHERE {2}='{3}'",
        strSQL,
        TableName,
        IdxColumnName.UUID_NO.ToString, CI.get_UUID_NO,
        IdxColumnName.UUID_SEQ.ToString, CI.get_UUID_SEQ,
        IdxColumnName.IDLENGTH.ToString, CI.get_IDLENGTH,
        IdxColumnName.APPEND.ToString, CI.get_APPEND,
        IdxColumnName.COMMENTS.ToString, CI.get_COMMENTS,
        IdxColumnName.RESETABLE.ToString, CI.get_RESETABLE,
        IdxColumnName.UPDATE_DATE.ToString, CI.get_UPDATE_DATE
        )
        lstSql.Add(strSQL)
      Next
      Dim NewSQL As New List(Of String)
      If SQLCorrect(lstSql, NewSQL) = False Then
        Return Nothing
      End If
      If SendSQLToDB(NewSQL) = True Then
        Return True
      Else
        SendMessageToLog("Update to WMS_M_UUIDData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function DeleteWMS_M_UUIDDataToDB(ByRef Info As List(Of clsUUID)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
        strSQL,
        TableName,
        IdxColumnName.UUID_NO.ToString, CI.get_UUID_NO,
        IdxColumnName.UUID_SEQ.ToString, CI.get_UUID_SEQ,
        IdxColumnName.IDLENGTH.ToString, CI.get_IDLENGTH,
        IdxColumnName.APPEND.ToString, CI.get_APPEND,
        IdxColumnName.COMMENTS.ToString, CI.get_COMMENTS,
        IdxColumnName.RESETABLE.ToString, CI.get_RESETABLE,
        IdxColumnName.UPDATE_DATE.ToString, CI.get_UPDATE_DATE
        )
        lstSql.Add(strSQL)
      Next

      Dim NewSQL As New List(Of String)
      If SQLCorrect(lstSql, NewSQL) = False Then
        Return Nothing
      End If
      If SendSQLToDB(NewSQL) = True Then
        Return True
      Else
        SendMessageToLog("Delete WMS_M_UUIDData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '-內部記憶體增刪修
  Private Shared Function AddOrUpdateWMS_M_UUIDDataToDictionary(ByRef Info As List(Of clsUUID)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      For Each CI In Info
        Dim _Data As clsUUID = CI
        Dim key As String = _Data.get_gid()
        dicData.AddOrUpdate(key,
        _Data,
        Function(dicKey, ExistVal)
          UpdateInfo(dicKey, ExistVal, _Data)
          Return ExistVal
        End Function)
      Next

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  Private Shared Function DeleteWMS_M_UUIDDataToDictionary(ByRef Info As List(Of clsUUID)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      For i = 0 To Info.Count - 1
        Dim key As String = Info(i).get_gid()
        If dicData.TryRemove(key, Nothing) = False Then

          SendMessageToLog("dicData.TryRemove Failed -WMS_M_UUIDData", eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
      Next

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsUUID, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim UUID_NO = "" & RowData.Item(IdxColumnName.UUID_NO.ToString)
        Dim UUID_SEQ = 0 & RowData.Item(IdxColumnName.UUID_SEQ.ToString)
        Dim IDLENGTH = 0 & RowData.Item(IdxColumnName.IDLENGTH.ToString)
        Dim APPEND = "" & RowData.Item(IdxColumnName.APPEND.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim RESETABLE = "" & RowData.Item(IdxColumnName.RESETABLE.ToString)
        Dim UPDATE_DATE = "" & RowData.Item(IdxColumnName.UPDATE_DATE.ToString)
        Info = New clsUUID(UUID_NO, UUID_SEQ, IDLENGTH, APPEND, COMMENTS, RESETABLE, UPDATE_DATE)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateInfo(ByRef Key As String, ByRef Info As clsUUID, ByRef objNewTC As clsUUID) As clsUUID
    Try
      If Key = Info.get_gid() Then
        Info.Update_UUID(objNewTC)

      Else
        SendMessageToLog("Dictionary has the different key", eCALogTool.ILogTool.enuTrcLevel.lvError)
      End If

    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
    Return Info
  End Function
  Private Shared Function SendSQLToDB(ByRef lstSQL As List(Of String)) As Boolean
    Try
      If lstSQL Is Nothing Then Return False
      If lstSQL.Count = 0 Then Return True
      For i = 0 To lstSQL.Count - 1
        SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next
      If fUseBatchUpdate = 0 Then
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
