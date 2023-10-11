Partial Class WMS_CH_CLASS_ASSIGNATION_HISTManagement
Public Shared TableName As String = "WMS_CH_CLASS_ASSIGNATION_HIST"
Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
 
Enum IdxColumnName As Integer
FACTORY_NO
AREA_NO
CLASS_NO
ASSIGNATION_RATE
UPDATE_USER
HIST_TIME
End Enum
	'- GetSQL
	Public Shared Function GetInsertSQL(ByRef Info As clsCLASS_ASSIGNATION_HIST) As String
		Try

			Dim strSQL As String = ""
			strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12}) values ('{3}','{5}','{7}',{9},'{11}','{13}')",
 strSQL,
 TableName,
 IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
 IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
 IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
 IdxColumnName.ASSIGNATION_RATE.ToString, Info.ASSIGNATION_RATE,
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
