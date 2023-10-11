Partial Class WMS_CH_CLASS_ATTENDANCE_HISTManagement
Public Shared TableName As String = "WMS_CH_CLASS_ATTENDANCE_HIST"
Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
 
Enum IdxColumnName As Integer
CLASS_NO
ATTENDANCE_COUNT
UPDATE_USER
HIST_TIME
End Enum
	'- GetSQL
	Public Shared Function GetInsertSQL(ByRef Info As clsCLASS_ATTENDANCE_HIST) As String
		Try

			Dim strSQL As String = ""
			strSQL = String.Format("Insert into {1} ({2},{4},{6},{8}) values ('{3}',{5},'{7}','{9}')",
 strSQL,
 TableName,
 IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
 IdxColumnName.ATTENDANCE_COUNT.ToString, Info.ATTENDANCE_COUNT,
 IdxColumnName.UPDATE_USER.ToString, Info.UPDATE_USER,
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
End Class
