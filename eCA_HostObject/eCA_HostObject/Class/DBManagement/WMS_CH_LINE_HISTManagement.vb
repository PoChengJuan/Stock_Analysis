Public Class WMS_CH_LINE_HISTManagement
  Public Shared TableName As String = "WMS_CH_LINE_HIST"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    OCCUR_TIME
    MAINTENANCE_MESSAGE
    REMOVE_USER
		HIST_TIME
		MAINTENANCE_ID
		FUCTION_ID
		OPERATOR_USER
		COMMENTS

	End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsLineInfo_Hist) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
      IdxColumnName.OCCUR_TIME.ToString, Info.Occur_Time,
      IdxColumnName.MAINTENANCE_MESSAGE.ToString, Info.Maintenance_Message,
      IdxColumnName.REMOVE_USER.ToString, Info.Remove_User,
      IdxColumnName.HIST_TIME.ToString, Info.Hist_Time,
      IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
      IdxColumnName.FUCTION_ID.ToString, Info.FUCTION_ID,
      IdxColumnName.OPERATOR_USER.ToString, Info.OPERATOR_USER,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS
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
