Public Class HS_T_HOST_COMMANDManagement
  Public Shared TableName As String = "HS_T_HOST_COMMAND"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
  Private Shared fUseBatchUpdate As Integer = 1

  Enum IdxColumnName As Integer
    UUID
    CONNECTION_TYPE
    SEND_SYSTEM
    FUNCTION_ID
    SEQ
    USER_ID
    CLIENT_ID
    IP
    CREATE_TIME
    MESSAGE
    RESULT
    RESULT_MESSAGE
    WAIT_UUID
  End Enum

  Public Enum UpdateOption As Integer
    UpdateDic = 0
    UpdateDB = 1
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsHSToHostCommand) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26}) values ('{3}',{5},{7},'{9}',{11},'{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}')",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, CI.UUID,
      IdxColumnName.CONNECTION_TYPE.ToString, CInt(CI.CONNECTION_TYPE),
      IdxColumnName.SEND_SYSTEM.ToString, CInt(CI.SEND_SYSTEM),
      IdxColumnName.FUNCTION_ID.ToString, CI.FUNCTION_ID,
      IdxColumnName.SEQ.ToString, CI.SEQ,
      IdxColumnName.USER_ID.ToString, CI.USER_ID,
      IdxColumnName.CLIENT_ID.ToString, CI.CLIENT_ID,
      IdxColumnName.IP.ToString, CI.IP,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
      IdxColumnName.MESSAGE.ToString, CI.MESSAGE,
      IdxColumnName.RESULT.ToString, CI.RESULT,
      IdxColumnName.RESULT_MESSAGE.ToString, CI.RESULT_MESSAGE,
      IdxColumnName.WAIT_UUID.ToString, CI.WAIT_UUID
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
  Public Shared Function GetDeleteSQL(ByRef CI As clsHSToHostCommand) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' and {8}='{9}' and {10}='{11}'",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, CI.UUID,
      IdxColumnName.CONNECTION_TYPE.ToString, CInt(CI.CONNECTION_TYPE),
      IdxColumnName.SEND_SYSTEM.ToString, CInt(CI.SEND_SYSTEM),
      IdxColumnName.FUNCTION_ID.ToString, CI.FUNCTION_ID,
      IdxColumnName.SEQ.ToString, CI.SEQ
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
  Public Shared Function GetUpdateSQL(ByRef CI As clsHSToHostCommand) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}={7},{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}' WHERE {2}='{3}' and {8}='{9}' and {10}='{11}'",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, CI.UUID,
      IdxColumnName.CONNECTION_TYPE.ToString, CInt(CI.CONNECTION_TYPE),
      IdxColumnName.SEND_SYSTEM.ToString, CInt(CI.SEND_SYSTEM),
      IdxColumnName.FUNCTION_ID.ToString, CI.FUNCTION_ID,
      IdxColumnName.SEQ.ToString, CI.SEQ,
      IdxColumnName.USER_ID.ToString, CI.USER_ID,
      IdxColumnName.CLIENT_ID.ToString, CI.CLIENT_ID,
      IdxColumnName.IP.ToString, CI.IP,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
      IdxColumnName.MESSAGE.ToString, CI.MESSAGE,
      IdxColumnName.RESULT.ToString, CInt(CI.RESULT),
      IdxColumnName.RESULT_MESSAGE.ToString, CI.RESULT_MESSAGE,
      IdxColumnName.WAIT_UUID.ToString, CI.WAIT_UUID
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
  Public Shared Function UpdateWaitUUID_ByUUID(ByVal UUID As String, ByVal WaitUUID As String) As String
    Try
      Dim strSQL As String = ""
      Dim lstSql As New List(Of String)
      strSQL = String.Format("Update {1} SET {4}='{5}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, UUID,
      IdxColumnName.WAIT_UUID.ToString, WaitUUID
      )
      lstSql.Add(strSQL)
      Dim NewSQL As New List(Of String)
      If SQLCorrect(lstSql, NewSQL) Then
        If SendSQLToDB(NewSQL) = True Then
          Return True
        Else
          SendMessageToLog("Update to WMS_M_UUIDData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        End If
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function UpdateResult_ResultMessage_ByUUID(ByVal UUID As String, ByVal Result As String, ByVal RESULT_MESSAGE As String) As String
    Try
      If RESULT_MESSAGE.Length > 3000 Then RESULT_MESSAGE = RESULT_MESSAGE.Substring(0, 3000)
      Dim strSQL As String = ""
      Dim lstSql As New List(Of String)
      strSQL = String.Format("Update {1} SET {4}='{5}', {6}='{7}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, UUID,
      IdxColumnName.RESULT.ToString, Result,
      IdxColumnName.RESULT_MESSAGE.ToString, RESULT_MESSAGE
      )
      lstSql.Add(strSQL)
      Dim NewSQL As New List(Of String)
      If SQLCorrect(lstSql, NewSQL) Then
        If SendSQLToDB(NewSQL) = True Then
          Return True
        Else
          SendMessageToLog("Update to WMS_M_UUIDData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        End If
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Private Shared Function SetInfoFromDB(ByRef Info As clsHSToHostCommand, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim UUID = "" & RowData.Item(IdxColumnName.UUID.ToString)
        Dim CONNECTION_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.CONNECTION_TYPE.ToString)), RowData.Item(IdxColumnName.CONNECTION_TYPE.ToString), 0 & RowData.Item(IdxColumnName.CONNECTION_TYPE.ToString))
        Dim SEND_SYSTEM = If(IsNumeric(RowData.Item(IdxColumnName.SEND_SYSTEM.ToString)), RowData.Item(IdxColumnName.SEND_SYSTEM.ToString), 0 & RowData.Item(IdxColumnName.SEND_SYSTEM.ToString))
        Dim FUNCTION_ID = "" & RowData.Item(IdxColumnName.FUNCTION_ID.ToString)
        Dim SEQ = If(IsNumeric(RowData.Item(IdxColumnName.SEQ.ToString)), RowData.Item(IdxColumnName.SEQ.ToString), 0 & RowData.Item(IdxColumnName.SEQ.ToString))
        Dim USER_ID = "" & RowData.Item(IdxColumnName.USER_ID.ToString)
        Dim CLIENT_ID = "" & RowData.Item(IdxColumnName.CLIENT_ID.ToString)
        Dim IP = "" & RowData.Item(IdxColumnName.IP.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim MESSAGE = "" & RowData.Item(IdxColumnName.MESSAGE.ToString)
        Dim RESULT = "" & RowData.Item(IdxColumnName.RESULT.ToString)
        Dim RESULT_MESSAGE = "" & RowData.Item(IdxColumnName.RESULT_MESSAGE.ToString)
        Dim WAIT_UUID = "" & RowData.Item(IdxColumnName.WAIT_UUID.ToString)

        Info = New clsHSToHostCommand(UUID, CONNECTION_TYPE, SEND_SYSTEM, FUNCTION_ID, SEQ, USER_ID, CLIENT_ID, IP, CREATE_TIME, MESSAGE, RESULT, RESULT_MESSAGE, WAIT_UUID, GetNewTime_DBFormat)
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

  Public Shared Function GetHS_T_CommandByUUID(ByVal UUID As String) As Dictionary(Of String, clsHSToHostCommand)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsHSToHostCommand)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE {2}='{3}'",
        strSQL,
        TableName,
        IdxColumnName.UUID.ToString, UUID
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsHSToHostCommand = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetHS_T_CommandResultIsNotNull(ByVal SEND_SYSTEM As enuSystemType) As Dictionary(Of String, clsHSToHostCommand)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsHSToHostCommand)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE {2} is not null AND {3} = '{4}'",
        strSQL,
        TableName,
        IdxColumnName.RESULT.ToString,
        IdxColumnName.SEND_SYSTEM, CInt(SEND_SYSTEM)
        )
        'SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsHSToHostCommand = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetHS_T_Command_Timeout(ByVal SEND_SYSTEM As enuSystemType) As Dictionary(Of String, clsHSToHostCommand)
    Try
      '取得命令建立時間大於兩小時的
      Dim _lstReturn As New Dictionary(Of String, clsHSToHostCommand)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE {2} < '{3}' AND ({4} = '' or {4} is null) AND {5} = '{6}'",
        strSQL,
        TableName,
        IdxColumnName.CREATE_TIME.ToString, AddTractTime_Hour(GetNewTime_DBFormat(), -2),
        IdxColumnName.RESULT,
        IdxColumnName.SEND_SYSTEM, CInt(SEND_SYSTEM)
        )
        'SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsHSToHostCommand = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function



  '從資料庫抓取Command，使用Wait_UUID當條件
  Public Shared Function GetCommandDictionaryByReceiveSystem_WaitUUID(ByVal Wait_UUID As String) As Dictionary(Of String, clsHSToHostCommand)

    Try
      Dim ret_dic As New Dictionary(Of String, clsHSToHostCommand)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * FROM {0} WHERE ({1} IS NULL OR {1} = '') AND {2}='{3}'",
                                   TableName,
                                   IdxColumnName.RESULT.ToString,
                                   IdxColumnName.WAIT_UUID.ToString, Wait_UUID)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsHSToHostCommand = Nothing
            If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
              If Info IsNot Nothing Then
                If ret_dic.ContainsKey(Info.gid) = False Then
                  ret_dic.Add(Info.gid, Info)
                End If
              Else
                SendMessageToLog("Get clsHSToHostCommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Else
              SendMessageToLog("Get clsHSToHostCommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            End If
          Next
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
