Partial Class WMS_CH_COUNT_MODIFY_HISTManagement
	Public Shared TableName As String = "WMS_CH_COUNT_MODIFY_HIST"
	Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

	Enum IdxColumnName As Integer
		FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    USER_ID
		QTY_MODIFY
		QTY_NG
		HIST_TIME
	End Enum
	'- GetSQL
	Public Shared Function GetInsertSQL(ByRef Info As clsCOUNT_MODIFY_HIST) As String
		Try

			Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}')",
       strSQL,
       TableName,
       IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
       IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
       IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
       IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
       IdxColumnName.USER_ID.ToString, Info.USER_ID,
       IdxColumnName.QTY_MODIFY.ToString, Info.QTY_MODIFY,
       IdxColumnName.QTY_NG.ToString, Info.QTY_NG,
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
