Partial Class WMS_M_MAINTENANCE_DTLManagement
	Public Shared TableName As String = "WMS_M_MAINTENANCE_DTL"
	Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

	Enum IdxColumnName As Integer
		FACTORY_NO
		DEVICE_NO
		AREA_NO
    UNIT_ID
    MAINTENANCE_ID
    FUNCTION_ID
    VALUE_TYPE
    NOTICE_TYPE
    HIGH_WATER_VALUE
    LOW_WATER_VALUE
    STANDARD_VALUE
    VALUE_RANGE
    MAINTENANCE_MESSAGE
    VALUE_SOURCE
    VALUE_UPDATE_TYPE
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMAINTENANCE_DTL) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30}) values ('{3}','{5}','{7}','{9}','{11}','{13}',{15},{17},'{19}','{21}','{23}','{25}','{27}',{29},{31})",
 strSQL,
 TableName,
 IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
 IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
 IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
 IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
 IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
 IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
 IdxColumnName.VALUE_TYPE.ToString, Info.VALUE_TYPE,
 IdxColumnName.NOTICE_TYPE.ToString, Info.NOTICE_TYPE,
 IdxColumnName.HIGH_WATER_VALUE.ToString, Info.HIGH_WATER_VALUE,
 IdxColumnName.LOW_WATER_VALUE.ToString, Info.LOW_WATER_VALUE,
 IdxColumnName.STANDARD_VALUE.ToString, Info.STANDARD_VALUE,
 IdxColumnName.VALUE_RANGE.ToString, Info.VALUE_RANGE,
 IdxColumnName.MAINTENANCE_MESSAGE.ToString, Info.MAINTENANCE_MESSAGE,
 IdxColumnName.VALUE_SOURCE.ToString, Info.VALUE_SOURCE,
 IdxColumnName.VALUE_UPDATE_TYPE.ToString, Info.VALUE_UPDATE_TYPE
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMAINTENANCE_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {14}={15},{16}={17},{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}={29},{30}={31} WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}' And {10}='{11}' And {12}='{13}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
      IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
      IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
      IdxColumnName.VALUE_TYPE.ToString, Info.VALUE_TYPE,
      IdxColumnName.NOTICE_TYPE.ToString, Info.NOTICE_TYPE,
      IdxColumnName.HIGH_WATER_VALUE.ToString, Info.HIGH_WATER_VALUE,
      IdxColumnName.LOW_WATER_VALUE.ToString, Info.LOW_WATER_VALUE,
      IdxColumnName.STANDARD_VALUE.ToString, Info.STANDARD_VALUE,
      IdxColumnName.VALUE_RANGE.ToString, Info.VALUE_RANGE,
      IdxColumnName.MAINTENANCE_MESSAGE.ToString, Info.MAINTENANCE_MESSAGE,
      IdxColumnName.VALUE_SOURCE.ToString, Info.VALUE_SOURCE,
      IdxColumnName.VALUE_UPDATE_TYPE.ToString, Info.VALUE_UPDATE_TYPE
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMAINTENANCE_DTL) As Integer
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' AND {10}='{11}' AND {12}='{13}' ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
      IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
      IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
      IdxColumnName.VALUE_TYPE.ToString, Info.VALUE_TYPE,
      IdxColumnName.NOTICE_TYPE.ToString, Info.NOTICE_TYPE,
      IdxColumnName.HIGH_WATER_VALUE.ToString, Info.HIGH_WATER_VALUE,
      IdxColumnName.LOW_WATER_VALUE.ToString, Info.LOW_WATER_VALUE,
      IdxColumnName.STANDARD_VALUE.ToString, Info.STANDARD_VALUE,
      IdxColumnName.VALUE_RANGE.ToString, Info.VALUE_RANGE,
      IdxColumnName.MAINTENANCE_MESSAGE.ToString, Info.MAINTENANCE_MESSAGE,
      IdxColumnName.VALUE_SOURCE.ToString, Info.VALUE_SOURCE,
      IdxColumnName.VALUE_UPDATE_TYPE.ToString, Info.VALUE_UPDATE_TYPE
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

	'- GET
	Public Shared Function GetWMS_M_MaintenanceDTLDataListByALL() As List(Of clsMAINTENANCE_DTL)
		Try
			Dim _lstReturn As New List(Of clsMAINTENANCE_DTL)
			If DBTool IsNot Nothing Then
				'If DBTool.isConnection(DBTool.m_CN) = True Then
				Dim strSQL As String = String.Empty
				Dim rs As DataSet = Nothing
				Dim DatasetMessage As New DataSet
				strSQL = String.Format("Select * from {0}", TableName)
				SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
				DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

				'Dim OLEDBAdapter As New OleDbDataAdapter
				'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

				If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
					For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
						Dim Info As clsMAINTENANCE_DTL = Nothing
						SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
						_lstReturn.Add(Info)
					Next
				End If
				'End If
			End If
			Return _lstReturn
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	Public Shared Function SelectWMS_M_MAINTENANCE_DTLDataByFACTORY_NO_DEVICE_NO_AREA_NO_UNIT_ID_MAINTENANCE_ID_FUNCTION_ID(ByVal factory_no As String,
																																									ByVal device_no As String, ByVal area_no As String,
																																									ByVal unit_id As String, ByVal maintenance_id As String,
																																												ByVal function_id As String) As Dictionary(Of String, clsMAINTENANCE_DTL)
		Try
			Dim dicWMS_M_MAINTENANCE_DTL As New Dictionary(Of String, clsMAINTENANCE_DTL)
			If DBTool IsNot Nothing Then
				If DBTool.isConnection(DBTool.m_CN) = True Then
					Dim strSQL As String = String.Empty
					Dim strSQLWhere As String = ""
					If factory_no <> "" Then
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.FACTORY_NO.ToString, factory_no)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.FACTORY_NO.ToString, factory_no)
						End If
					Else
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0} is null ", IdxColumnName.FACTORY_NO.ToString)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0} is null ", IdxColumnName.FACTORY_NO.ToString)
						End If
					End If

					If device_no <> "" Then
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.DEVICE_NO.ToString, device_no)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.DEVICE_NO.ToString, device_no)
						End If
					Else
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0} is null ", IdxColumnName.DEVICE_NO.ToString)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0} is null ", IdxColumnName.DEVICE_NO.ToString)
						End If
					End If

					If area_no <> "" Then
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.AREA_NO.ToString, area_no)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.AREA_NO.ToString, area_no)
						End If
					Else
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0} is null ", IdxColumnName.AREA_NO.ToString)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0} is null ", IdxColumnName.AREA_NO.ToString)
						End If
					End If

					If unit_id <> "" Then
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.UNIT_ID.ToString, unit_id)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.UNIT_ID.ToString, unit_id)
						End If
					Else
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0} is null ", IdxColumnName.UNIT_ID.ToString)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0} is null ", IdxColumnName.UNIT_ID.ToString)
						End If
					End If

					If maintenance_id <> "" Then
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.MAINTENANCE_ID.ToString, maintenance_id)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.MAINTENANCE_ID.ToString, maintenance_id)
						End If
					Else
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0} is null ", IdxColumnName.MAINTENANCE_ID.ToString)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0} is null ", IdxColumnName.MAINTENANCE_ID.ToString)
						End If
					End If

					If function_id <> "" Then
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0}='{1}' ", IdxColumnName.FUNCTION_ID.ToString, function_id)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.FUNCTION_ID.ToString, function_id)
						End If
					Else
						If strSQLWhere = "" Then
							strSQLWhere = String.Format("  WHERE {0} is null ", IdxColumnName.FUNCTION_ID.ToString)
						Else
							strSQLWhere = strSQLWhere & String.Format(" AND {0} is null ", IdxColumnName.FUNCTION_ID.ToString)
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
					Dim _lstReturn As New List(Of clsMAINTENANCE_DTL)
					If DatasetMessage.Tables(TableName).Rows.Count > 0 Then
						For RowIndex = 0 To DatasetMessage.Tables(TableName).Rows.Count - 1
							Dim Info As clsMAINTENANCE_DTL = Nothing
							If SetInfoFromDB(Info, DatasetMessage.Tables(TableName).Rows(RowIndex)) = True Then
								If Info IsNot Nothing Then
									If dicWMS_M_MAINTENANCE_DTL.ContainsKey(Info.gid) = False Then
										dicWMS_M_MAINTENANCE_DTL.Add(Info.gid, Info)
									End If
								Else
									SendMessageToLog(" Select clsWMS_M_MAINTENANCE_DTL Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
								End If
							Else
								SendMessageToLog(" Select clsWMS_M_MAINTENANCE_DTL Info Failed ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
							End If
						Next
					End If
				End If
			End If
			Return dicWMS_M_MAINTENANCE_DTL
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function

	Private Shared Function SetInfoFromDB(ByRef Info As clsMAINTENANCE_DTL, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim FACTORY_NO = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim DEVICE_NO = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim AREA_NO = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim UNIT_ID = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
        Dim MAINTENANCE_ID = "" & RowData.Item(IdxColumnName.MAINTENANCE_ID.ToString)
        Dim FUNCTION_ID = "" & RowData.Item(IdxColumnName.FUNCTION_ID.ToString)
        Dim VALUE_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.VALUE_TYPE.ToString)), RowData.Item(IdxColumnName.VALUE_TYPE.ToString), 0 & RowData.Item(IdxColumnName.VALUE_TYPE.ToString))
        Dim NOTICE_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.NOTICE_TYPE.ToString)), RowData.Item(IdxColumnName.NOTICE_TYPE.ToString), 0 & RowData.Item(IdxColumnName.NOTICE_TYPE.ToString))
        Dim HIGH_WATER_VALUE = "" & RowData.Item(IdxColumnName.HIGH_WATER_VALUE.ToString)
        Dim LOW_WATER_VALUE = "" & RowData.Item(IdxColumnName.LOW_WATER_VALUE.ToString)
        Dim STANDARD_VALUE = "" & RowData.Item(IdxColumnName.STANDARD_VALUE.ToString)
        Dim VALUE_RANGE = "" & RowData.Item(IdxColumnName.VALUE_RANGE.ToString)
        Dim MAINTENANCE_MESSAGE = "" & RowData.Item(IdxColumnName.MAINTENANCE_MESSAGE.ToString)
        Dim VALUE_SOURCE = If(IsNumeric(RowData.Item(IdxColumnName.VALUE_SOURCE.ToString)), RowData.Item(IdxColumnName.VALUE_SOURCE.ToString), 0 & RowData.Item(IdxColumnName.VALUE_SOURCE.ToString))
        Dim VALUE_UPDATE_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.VALUE_UPDATE_TYPE.ToString)), RowData.Item(IdxColumnName.VALUE_UPDATE_TYPE.ToString), 0 & RowData.Item(IdxColumnName.VALUE_UPDATE_TYPE.ToString))
        Info = New clsMAINTENANCE_DTL(FACTORY_NO, DEVICE_NO, AREA_NO, UNIT_ID, MAINTENANCE_ID, FUNCTION_ID, VALUE_TYPE, NOTICE_TYPE, HIGH_WATER_VALUE, LOW_WATER_VALUE, STANDARD_VALUE, VALUE_RANGE, MAINTENANCE_MESSAGE, VALUE_SOURCE, VALUE_UPDATE_TYPE)

      End If
      Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
