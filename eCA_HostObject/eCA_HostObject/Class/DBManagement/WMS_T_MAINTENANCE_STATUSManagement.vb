Partial Class WMS_T_MAINTENANCE_STATUSManagement
	Public Shared TableName As String = "WMS_T_MAINTENANCE_STATUS"
	Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

	Enum IdxColumnName As Integer
		FACTORY_NO
		DEVICE_NO
		AREA_NO
    UNIT_ID
    MAINTENANCE_ID
    FUNCTION_ID
    VALUE
    UPDATE_TIME
    MAINTENANCE_SET
    MAINTENANCE_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMAINTENANCE_STATUS) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}',{19},'{21}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
      IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
      IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
      IdxColumnName.VALUE.ToString, Info.VALUE,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.MAINTENANCE_SET.ToString, BooleanConvertToInteger(Info.MAINTENANCE_SET),
      IdxColumnName.MAINTENANCE_TIME.ToString, Info.MAINTENANCE_TIME
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMAINTENANCE_STATUS) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {14}='{15}',{16}='{17}',{18}={19},{20}='{21}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}' And {10}='{11}' And {12}='{13}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.DEVICE_NO.ToString, Info.DEVICE_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.UNIT_ID.ToString, Info.UNIT_ID,
      IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
      IdxColumnName.FUNCTION_ID.ToString, Info.FUNCTION_ID,
      IdxColumnName.VALUE.ToString, Info.VALUE,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.MAINTENANCE_SET.ToString, BooleanConvertToInteger(Info.MAINTENANCE_SET),
      IdxColumnName.MAINTENANCE_TIME.ToString, Info.MAINTENANCE_TIME
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMAINTENANCE_STATUS) As Integer
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
      IdxColumnName.VALUE.ToString, Info.VALUE,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.MAINTENANCE_SET.ToString, BooleanConvertToInteger(Info.MAINTENANCE_SET),
      IdxColumnName.MAINTENANCE_TIME.ToString, Info.MAINTENANCE_TIME
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
  Public Shared Function GetWMS_T_MaintenanceStatusDataListByALL() As List(Of clsMAINTENANCE_STATUS)
    Try
      Dim _lstReturn As New List(Of clsMAINTENANCE_STATUS)
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
            Dim Info As clsMAINTENANCE_STATUS = Nothing
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsMAINTENANCE_STATUS, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim FACTORY_NO = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim DEVICE_NO = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim AREA_NO = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim UNIT_ID = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
        Dim MAINTENANCE_ID = "" & RowData.Item(IdxColumnName.MAINTENANCE_ID.ToString)
        Dim FUNCTION_ID = "" & RowData.Item(IdxColumnName.FUNCTION_ID.ToString)
        Dim VALUE = "" & RowData.Item(IdxColumnName.VALUE.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim MAINTENANCE_SET = IntegerConvertToBoolean(RowData.Item(IdxColumnName.MAINTENANCE_SET.ToString))
        Dim MAINTENANCE_TIME = "" & RowData.Item(IdxColumnName.MAINTENANCE_TIME.ToString)
        Info = New clsMAINTENANCE_STATUS(FACTORY_NO, DEVICE_NO, AREA_NO, UNIT_ID, MAINTENANCE_ID, FUNCTION_ID, VALUE,
                                  UPDATE_TIME, MAINTENANCE_SET, MAINTENANCE_TIME)

      End If
      Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
