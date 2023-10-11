
Public Class HOST_H_Command_HistManagement
  Public Shared TableName As String = "HOST_H_Command_Hist"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    UUID
    SEND_SYSTEM
    RECEIVE_SYSTEM
    FUNCTION_ID
    SEQ
    USER_ID
    CREATE_TIME
    MESSAGE
    RESULT
    RESULT_MESSAGE
    WAIT_UUID
    HIST_TIME
  End Enum
  Public Shared Function GetInsertSQL(ByRef Info As clsFromHostCommandHist) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}')",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, Info.UUID,
      IdxColumnName.SEND_SYSTEM.ToString, CInt(Info.Send_System),
      IdxColumnName.RECEIVE_SYSTEM.ToString, CInt(Info.Receive_System),
      IdxColumnName.FUNCTION_ID.ToString, Info.Function_ID,
      IdxColumnName.SEQ.ToString, Info.SEQ,
      IdxColumnName.USER_ID.ToString, Info.User_ID,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.MESSAGE.ToString, Info.Message,
      IdxColumnName.RESULT.ToString, Info.Result,
      IdxColumnName.RESULT_MESSAGE.ToString, Info.Result_Message,
      IdxColumnName.WAIT_UUID.ToString, Info.Wait_UUID,
      IdxColumnName.HIST_TIME, Info.Hist_Time
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
