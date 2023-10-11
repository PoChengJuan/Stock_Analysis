
Partial Class WMS_H_ALARM_HISTManagement
  Public Shared TableName As String = "WMS_H_ALARM_HIST"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    OCCUR_TIME
    ALARM_CODE
    ALARM_TYPE
    CMD_ID
    SEND_STATUS
    CLEAR_TIME
    HIST_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsALARM_HIST) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22}) values ('{3}','{5}','{7}','{9}','{11}','{13}',{15},'{17}',{19},'{21}','{23}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
      IdxColumnName.DEVICE_NO.ToString, CI.DEVICE_NO,
      IdxColumnName.UNIT_ID.ToString, CI.UNIT_ID,
      IdxColumnName.OCCUR_TIME.ToString, CI.OCCUR_TIME,
      IdxColumnName.ALARM_CODE.ToString, CI.ALARM_CODE,
      IdxColumnName.ALARM_TYPE.ToString, CInt(CI.ALARM_TYPE),
      IdxColumnName.CMD_ID.ToString, CI.CMD_ID,
      IdxColumnName.SEND_STATUS.ToString, CInt(CI.SEND_STATUS),
      IdxColumnName.CLEAR_TIME.ToString, CI.CLEAR_TIME,
      IdxColumnName.HIST_TIME.ToString, CI.HIST_TIME
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
