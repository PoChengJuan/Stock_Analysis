Partial Class WMS_M_DATA_REPORT_SETManagement
	Public Shared TableName As String = "WMS_M_DATA_REPORT_SET"
	Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

	Enum IdxColumnName As Integer
		ROLE_ID
		ROLE_TYPE
		FUNCTION_ID
		FUNCTION_NAME
		DEVICE_NO
		AREA_NO
		UNIT_ID
		HIGH_WATER_VALUE
		LOW_WATER_VALUE
		STANDARD_VALUE
		VALUE_RANGE
		NOTICE_TYPE
		CONTINUE_SEND
		SEND_INTERVAL
		ENABLE
	End Enum
	'- GetSQL
	Public Shared Function GetInsertSQL(ByRef Info As clsDATA_REPORT_SET) As String
		Try

			Dim strSQL As String = ""
			strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30}) values ('{3}',{5},'{7}','{9}','{11}','{13}','{15}',{17},{19},{21},{23},{25},{27},{29},{31})",
 strSQL,
 TableName,
 IdxColumnName.ROLE_ID.ToString, Info.ROLE_ID,
 IdxColumnName.ROLE_TYPE.ToString, Info.ROLE_TYPE,
 IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
 IdxColumnName.FUNCTION_NAME.ToString, Info.FUNCTION_NAME,
 IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
 IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
 IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
 IdxColumnName.HIGH_WATER_VALUE.ToString, Info.HIGH_WATER_VALUE,
 IdxColumnName.LOW_WATER_VALUE.ToString, Info.LOW_WATER_VALUE,
 IdxColumnName.STANDARD_VALUE.ToString, Info.STANDARD_VALUE,
 IdxColumnName.VALUE_RANGE.ToString, Info.VALUE_RANGE,
 IdxColumnName.NOTICE_TYPE.ToString, Info.NOTICE_TYPE,
 IdxColumnName.CONTINUE_SEND.ToString, Info.CONTINUE_SEND,
 IdxColumnName.SEND_INTERVAL.ToString, Info.SEND_INTERVAL,
 IdxColumnName.ENABLE.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.ENABLE)
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
	Public Shared Function GetUpdateSQL(ByRef Info As clsDATA_REPORT_SET) As String
		Try
			Dim strSQL As String = ""
			strSQL = String.Format("Update {1} SET {4}={5},{8}='{9}',{16}={17},{18}={19},{20}={21},{22}={23},{24}={25},{26}={27},{28}={29},{30}={31} WHERE {2}='{3}' And {6}='{7}' And {10}='{11}' And {12}='{13}' And {14}='{15}'",
			strSQL,
			TableName,
			IdxColumnName.ROLE_ID.ToString, Info.ROLE_ID,
			IdxColumnName.ROLE_TYPE.ToString, Info.ROLE_TYPE,
			IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
			IdxColumnName.FUNCTION_NAME.ToString, Info.FUNCTION_NAME,
			IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
			IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
			IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
			IdxColumnName.HIGH_WATER_VALUE.ToString, Info.HIGH_WATER_VALUE,
			IdxColumnName.LOW_WATER_VALUE.ToString, Info.LOW_WATER_VALUE,
			IdxColumnName.STANDARD_VALUE.ToString, Info.STANDARD_VALUE,
			IdxColumnName.VALUE_RANGE.ToString, Info.VALUE_RANGE,
			IdxColumnName.NOTICE_TYPE.ToString, Info.NOTICE_TYPE,
			IdxColumnName.CONTINUE_SEND.ToString, Info.CONTINUE_SEND,
			IdxColumnName.SEND_INTERVAL.ToString, Info.SEND_INTERVAL,
			IdxColumnName.ENABLE.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.ENABLE)
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
	Public Shared Function GetDeleteSQL(ByRef Info As clsDATA_REPORT_SET) As Integer
		Try
			Dim strSQL As String = ""
			strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {6}='{7}' AND {10}='{11}' AND {12}='{13}' AND {14}='{15}' ",
			strSQL,
			TableName,
			IdxColumnName.ROLE_ID.ToString, Info.ROLE_ID,
			IdxColumnName.ROLE_TYPE.ToString, Info.ROLE_TYPE,
			IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
			IdxColumnName.FUNCTION_NAME.ToString, Info.FUNCTION_NAME,
			IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
			IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
			IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
			IdxColumnName.HIGH_WATER_VALUE.ToString, Info.HIGH_WATER_VALUE,
			IdxColumnName.LOW_WATER_VALUE.ToString, Info.LOW_WATER_VALUE,
			IdxColumnName.STANDARD_VALUE.ToString, Info.STANDARD_VALUE,
			IdxColumnName.VALUE_RANGE.ToString, Info.VALUE_RANGE,
			IdxColumnName.NOTICE_TYPE.ToString, Info.NOTICE_TYPE,
			IdxColumnName.CONTINUE_SEND.ToString, Info.CONTINUE_SEND,
			IdxColumnName.SEND_INTERVAL.ToString, Info.SEND_INTERVAL,
			IdxColumnName.ENABLE.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.ENABLE)
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


	'get
	Public Shared Function SelectWMS_M_DATA_REPORT_SETDataByROLEID_FUNCTIONID_DEVICENO_AREANO_UNITID(ByVal role_id As String,
								ByVal function_id As String, ByVal device_no As String, ByVal area_no As String, ByVal unit_id As String) As Dictionary(Of String, clsDATA_REPORT_SET)
		Try
			Dim dicWMS_M_DATA_REPORT_SET As New Dictionary(Of String, clsDATA_REPORT_SET)
			If DBTool IsNot Nothing Then
				If DBTool.isConnection(DBTool.m_CN) = True Then
					Dim strSQL As String = String.Empty
					Dim strSQLWhere As String = ""
          If role_id <> "" Then
            If strSQLWhere = "" Then
              strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.ROLE_ID.ToString, role_id)
            Else
              strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.ROLE_ID.ToString, role_id)
            End If
          End If
          If function_id <> "" Then
            If strSQLWhere = "" Then
              strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.FUNCTION_ID.ToString, function_id)
            Else
              strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.FUNCTION_ID.ToString, function_id)
            End If
          End If
          If device_no <> "" Then
            If strSQLWhere = "" Then
              strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.DEVICE_NO.ToString, device_no)
            Else
              strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.DEVICE_NO.ToString, device_no)
            End If
          End If
          If area_no <> "" Then
            If strSQLWhere = "" Then
              strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.AREA_NO.ToString, area_no)
            Else
              strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.AREA_NO.ToString, area_no)
            End If
          End If
          If unit_id <> "" Then
            If strSQLWhere = "" Then
              strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.UNIT_ID.ToString, unit_id)
            Else
              strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.UNIT_ID.ToString, unit_id)
            End If
          End If
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("SELECT * FROM {1} {2} ",
           strSQL,
           TableName,
           strSQLWhere
           )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
					DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
          Dim _lstReturn As New List(Of clsDATA_REPORT_SET)
          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsDATA_REPORT_SET = Nothing
              SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
              If Info IsNot Nothing Then
                If dicWMS_M_DATA_REPORT_SET.ContainsKey(Info.gid) = False Then
                  dicWMS_M_DATA_REPORT_SET.Add(Info.gid, Info)
                End If
              Else
                SendMessageToLog(" Select clsWMS_M_DATA_REPORT_SET Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Next
          End If
        End If
			End If
			Return dicWMS_M_DATA_REPORT_SET
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function



	Private Shared Function SetInfoFromDB(ByRef Info As clsDATA_REPORT_SET, ByRef RowData As DataRow) As Boolean
		Try
			If RowData IsNot Nothing Then
				Dim ROLE_ID = "" & RowData.Item(IdxColumnName.ROLE_ID.ToString)
				Dim ROLE_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.ROLE_TYPE.ToString)), RowData.Item(IdxColumnName.ROLE_TYPE.ToString), 0 & RowData.Item(IdxColumnName.ROLE_TYPE.ToString))
				Dim FUNCTION_ID = "" & RowData.Item(IdxColumnName.FUNCTION_ID.ToString)
				Dim FUNCTION_NAME = "" & RowData.Item(IdxColumnName.FUNCTION_NAME.ToString)
				Dim DEVICE_NO = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
				Dim AREA_NO = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
				Dim UNIT_NO = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
				Dim HIGH_WATER_VALUE = If(IsNumeric(RowData.Item(IdxColumnName.HIGH_WATER_VALUE.ToString)), RowData.Item(IdxColumnName.HIGH_WATER_VALUE.ToString), 0 & RowData.Item(IdxColumnName.HIGH_WATER_VALUE.ToString))
        Dim LOW_WATER_VALUE = If(IsNumeric(RowData.Item(IdxColumnName.LOW_WATER_VALUE.ToString)), RowData.Item(IdxColumnName.LOW_WATER_VALUE.ToString), 0 & RowData.Item(IdxColumnName.LOW_WATER_VALUE.ToString))
        Dim STANDARD_VALUE = If(IsNumeric(RowData.Item(IdxColumnName.STANDARD_VALUE.ToString)), RowData.Item(IdxColumnName.STANDARD_VALUE.ToString), 0 & RowData.Item(IdxColumnName.STANDARD_VALUE.ToString))
        Dim VALUE_RANGE = If(IsNumeric(RowData.Item(IdxColumnName.VALUE_RANGE.ToString)), RowData.Item(IdxColumnName.VALUE_RANGE.ToString), 0 & RowData.Item(IdxColumnName.VALUE_RANGE.ToString))
        Dim NOTICE_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.NOTICE_TYPE.ToString)), RowData.Item(IdxColumnName.NOTICE_TYPE.ToString), 0 & RowData.Item(IdxColumnName.NOTICE_TYPE.ToString))
        Dim CONTINUE_SEND = If(IsNumeric(RowData.Item(IdxColumnName.CONTINUE_SEND.ToString)), RowData.Item(IdxColumnName.CONTINUE_SEND.ToString), 0 & RowData.Item(IdxColumnName.CONTINUE_SEND.ToString))
        Dim SEND_INTERVAL = If(IsNumeric(RowData.Item(IdxColumnName.SEND_INTERVAL.ToString)), RowData.Item(IdxColumnName.SEND_INTERVAL.ToString), 0 & RowData.Item(IdxColumnName.SEND_INTERVAL.ToString))
        Dim ENABLE = IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.ENABLE.ToString))
        Info = New clsDATA_REPORT_SET(ROLE_ID, ROLE_TYPE, FUNCTION_ID, FUNCTION_NAME, DEVICE_NO, AREA_NO, UNIT_NO, HIGH_WATER_VALUE, LOW_WATER_VALUE, STANDARD_VALUE, VALUE_RANGE, NOTICE_TYPE, CONTINUE_SEND, SEND_INTERVAL, ENABLE)

      End If
      Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
