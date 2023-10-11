Partial Class WMS_M_MAINTENANCEManagement
  Public Shared TableName As String = "WMS_M_MAINTENANCE"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    DEVICE_NO
    AREA_NO
    UNIT_ID
    MAINTENANCE_ID
    MAINTENANCE_NAME
    CONTINUE_SEND
    SEND_INTERVAL
    SEND_TYPE
    UPDATE_TIME
    ENABLE
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMAINTENANCE) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22}) values ('{3}','{5}','{7}','{9}','{11}','{13}',{15},{17},{19},'{21}',{23})",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
      IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
      IdxColumnName.MAINTENANCE_NAME.ToString, Info.MAINTENANCE_NAME,
      IdxColumnName.CONTINUE_SEND.ToString, Info.CONTINUE_SEND,
      IdxColumnName.SEND_INTERVAL.ToString, Info.SEND_INTERVAL,
      IdxColumnName.SEND_TYPE.ToString, Info.SEND_TYPE,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.ENABLE.ToString, BooleanConvertToInteger(Info.ENABLE)
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMAINTENANCE) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {12}='{13}',{14}={15},{16}={17},{18}={19},{20}='{21}',{22}={23} WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}' And {10}='{11}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
      IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
      IdxColumnName.MAINTENANCE_NAME.ToString, Info.MAINTENANCE_NAME,
      IdxColumnName.CONTINUE_SEND.ToString, Info.CONTINUE_SEND,
      IdxColumnName.SEND_INTERVAL.ToString, Info.SEND_INTERVAL,
      IdxColumnName.SEND_TYPE.ToString, Info.SEND_TYPE,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.ENABLE.ToString, BooleanConvertToInteger(Info.ENABLE)
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMAINTENANCE) As Integer
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' AND {10}='{11}' ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
      IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
      IdxColumnName.MAINTENANCE_NAME.ToString, Info.MAINTENANCE_NAME,
      IdxColumnName.CONTINUE_SEND.ToString, Info.CONTINUE_SEND,
      IdxColumnName.SEND_INTERVAL.ToString, Info.SEND_INTERVAL,
      IdxColumnName.SEND_TYPE.ToString, Info.SEND_TYPE,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.ENABLE.ToString, BooleanConvertToInteger(Info.ENABLE)
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
	Public Shared Function GetWMS_M_MaintenanceDataListByALL() As List(Of clsMAINTENANCE)
		Try
			Dim _lstReturn As New List(Of clsMAINTENANCE)
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
						Dim Info As clsMAINTENANCE = Nothing
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
	Public Shared Function SelectWMS_M_MAINTENANCEDataByFACTORY_NO_DEVICE_NO_AREA_NO_UNIT_ID_MAINTENANCE_ID(ByVal factory_no As String,
																																																					ByVal device_no As String,
																																																					ByVal area_no As String,
																																																					ByVal unit_id As String,
																																																					ByVal maintenance_id As String) As Dictionary(Of String, clsMAINTENANCE)
		Try
			Dim dicWMS_M_MAINTENANCE As New Dictionary(Of String, clsMAINTENANCE)
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

					'Dim rs As ADODB.Recordset = Nothing
					Dim DatasetMessage As New DataSet
					strSQL = String.Format("SELECT * FROM {1} {2} ",
 strSQL,
 TableName,
 strSQLWhere
 )
					SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
					DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
					'Dim DatasetMessage As New DataSet
					'Dim OLEDBAdapter As New OleDbDataAdapter
					'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)
					Dim _lstReturn As New List(Of clsMAINTENANCE)
					If DatasetMessage.Tables(TableName).Rows.Count > 0 Then
						For RowIndex = 0 To DatasetMessage.Tables(TableName).Rows.Count - 1
							Dim Info As clsMAINTENANCE = Nothing
							If SetInfoFromDB(Info, DatasetMessage.Tables(TableName).Rows(RowIndex)) = True Then
								If Info IsNot Nothing Then
									If dicWMS_M_MAINTENANCE.ContainsKey(Info.gid) = False Then
										dicWMS_M_MAINTENANCE.Add(Info.gid, Info)
									End If
								Else
									SendMessageToLog(" Select clsWMS_M_MAINTENANCE Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
								End If
							Else
								SendMessageToLog(" Select clsWMS_M_MAINTENANCE Info Failed ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
							End If
						Next
					End If
				End If
			End If
			Return dicWMS_M_MAINTENANCE
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function



	Private Shared Function SetInfoFromDB(ByRef Info As clsMAINTENANCE, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim FACTORY_NO = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim DEVICE_NO = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim AREA_NO = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim UNIT_ID = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
        Dim MAINTENANCE_ID = "" & RowData.Item(IdxColumnName.MAINTENANCE_ID.ToString)
        Dim MAINTENANCE_NAME = "" & RowData.Item(IdxColumnName.MAINTENANCE_NAME.ToString)
        Dim CONTINUE_SEND = If(IsNumeric(RowData.Item(IdxColumnName.CONTINUE_SEND.ToString)), RowData.Item(IdxColumnName.CONTINUE_SEND.ToString), 0 & RowData.Item(IdxColumnName.CONTINUE_SEND.ToString))
        Dim SEND_INTERVAL = If(IsNumeric(RowData.Item(IdxColumnName.SEND_INTERVAL.ToString)), RowData.Item(IdxColumnName.SEND_INTERVAL.ToString), 0 & RowData.Item(IdxColumnName.SEND_INTERVAL.ToString))
        Dim SEND_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.SEND_TYPE.ToString)), RowData.Item(IdxColumnName.SEND_TYPE.ToString), 0 & RowData.Item(IdxColumnName.SEND_TYPE.ToString))
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim ENABLE = IntegerConvertToBoolean(RowData.Item(IdxColumnName.ENABLE.ToString))
        Info = New clsMAINTENANCE(FACTORY_NO, DEVICE_NO, AREA_NO, UNIT_ID, MAINTENANCE_ID, MAINTENANCE_NAME, CONTINUE_SEND, SEND_INTERVAL, SEND_TYPE, UPDATE_TIME, ENABLE)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
