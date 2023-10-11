Partial Class HOST_H_HS_COMMANDManagement
  Public Shared TableName As String = "HOST_H_HS_COMMAND"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    UUID
    CONNECTION_TYPE
    RECEIVE_SYSTEM
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
    HIST_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsHostToHSCommand) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28}) values ('{3}',{5},{7},'{9}',{11},'{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}')",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, Info.UUID,
      IdxColumnName.CONNECTION_TYPE.ToString, CInt(Info.CONNECTION_TYPE),
      IdxColumnName.RECEIVE_SYSTEM.ToString, CInt(Info.RECEIVE_SYSTEM),
      IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
      IdxColumnName.SEQ.ToString, Info.SEQ,
      IdxColumnName.USER_ID.ToString, Info.USER_ID,
      IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
      IdxColumnName.IP.ToString, Info.IP,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.MESSAGE.ToString, Info.MESSAGE,
      IdxColumnName.RESULT.ToString, CInt(Info.RESULT),
      IdxColumnName.RESULT_MESSAGE.ToString, Info.RESULT_MESSAGE,
      IdxColumnName.WAIT_UUID.ToString, Info.WAIT_UUID,
      IdxColumnName.HIST_TIME.ToString, Info.HIST_TIME
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsHostToHSCommand) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {2}='{3}',{4}={5},{6}={7},{8}='{9}',{10}={11},{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}' WHERE ",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, Info.UUID,
      IdxColumnName.CONNECTION_TYPE.ToString, CInt(Info.CONNECTION_TYPE),
      IdxColumnName.RECEIVE_SYSTEM.ToString, CInt(Info.RECEIVE_SYSTEM),
      IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
      IdxColumnName.SEQ.ToString, Info.SEQ,
      IdxColumnName.USER_ID.ToString, Info.USER_ID,
      IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
      IdxColumnName.IP.ToString, Info.IP,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.MESSAGE.ToString, Info.MESSAGE,
      IdxColumnName.RESULT.ToString, CInt(Info.RESULT),
      IdxColumnName.RESULT_MESSAGE.ToString, Info.RESULT_MESSAGE,
      IdxColumnName.WAIT_UUID.ToString, Info.WAIT_UUID,
      IdxColumnName.HIST_TIME.ToString, Info.HIST_TIME
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsHostToHSCommand) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WH",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, Info.UUID,
      IdxColumnName.CONNECTION_TYPE.ToString, CInt(Info.CONNECTION_TYPE),
      IdxColumnName.RECEIVE_SYSTEM.ToString, CInt(Info.RECEIVE_SYSTEM),
      IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
      IdxColumnName.SEQ.ToString, Info.SEQ,
      IdxColumnName.USER_ID.ToString, Info.USER_ID,
      IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
      IdxColumnName.IP.ToString, Info.IP,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.MESSAGE.ToString, Info.MESSAGE,
      IdxColumnName.RESULT.ToString, CInt(Info.RESULT),
      IdxColumnName.RESULT_MESSAGE.ToString, Info.RESULT_MESSAGE,
      IdxColumnName.WAIT_UUID.ToString, Info.WAIT_UUID,
      IdxColumnName.HIST_TIME.ToString, Info.HIST_TIME
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsHostToHSCommand, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim UUID = "" & RowData.Item(IdxColumnName.UUID.ToString)
        Dim CONNECTION_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.CONNECTION_TYPE.ToString)), RowData.Item(IdxColumnName.CONNECTION_TYPE.ToString), 0 & RowData.Item(IdxColumnName.CONNECTION_TYPE.ToString))
        Dim RECEIVE_SYSTEM = If(IsNumeric(RowData.Item(IdxColumnName.RECEIVE_SYSTEM.ToString)), RowData.Item(IdxColumnName.RECEIVE_SYSTEM.ToString), 0 & RowData.Item(IdxColumnName.RECEIVE_SYSTEM.ToString))
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
        Dim HIST_TIME = "" & RowData.Item(IdxColumnName.HIST_TIME.ToString)
        Info = New clsHostToHSCommand(UUID, CONNECTION_TYPE, RECEIVE_SYSTEM, FUNCTION_ID, SEQ, USER_ID, CLIENT_ID, IP, CREATE_TIME, MESSAGE, RESULT, RESULT_MESSAGE, WAIT_UUID, HIST_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetHOST_H_HS_COMMANDListByALL() As List(Of clsHostToHSCommand)
    Try
      Dim _lstReturn As New List(Of clsHostToHSCommand)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} ",
        strSQL,
        TableName
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsHostToHSCommand = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info)
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
