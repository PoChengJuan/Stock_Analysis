Public Class WMS_CH_LINE_STATUS_HISTManagement
  Public Shared TableName As String = "WMS_CH_LINE_STATUS_HIST"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    FROM_STATUS
    TO_STATUS
    HIST_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsLine_Status_Hist) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
      IdxColumnName.FROM_STATUS.ToString, CInt(Info.From_Status),
      IdxColumnName.TO_STATUS.ToString, CInt(Info.To_Status),
      IdxColumnName.HIST_TIME.ToString, Info.Hist_Time
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
End Class
