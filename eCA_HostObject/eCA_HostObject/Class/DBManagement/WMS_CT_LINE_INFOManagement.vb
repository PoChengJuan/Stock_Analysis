Public Class WMS_CT_LINE_INFOManagement
  Public Shared TableName As String = "WMS_CT_LINE_INFO"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    OCCUR_TIME
		MAINTENANCE_MESSAGE
		MAINTENANCE_ID
		FUCTION_ID
	End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsLineInfo) As String
    Try

      Dim strSQL As String = ""
			strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}')",
			strSQL,
			TableName,
			IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
			IdxColumnName.AREA_NO.ToString, Info.Area_No,
			IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
			IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
			IdxColumnName.OCCUR_TIME.ToString, Info.Occur_Time,
			IdxColumnName.MAINTENANCE_MESSAGE.ToString, Info.Maintenance_Message,
														 IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
														 IdxColumnName.FUCTION_ID.ToString, Info.FUCTION_ID
		 )
			Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsLineInfo) As String
    Try

      Dim strSQL As String = ""
			strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' AND {14}='{15}' AND {16}='{17}' ",
			strSQL,
			TableName,
			IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
			IdxColumnName.AREA_NO.ToString, Info.Area_No,
			IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
			IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
			IdxColumnName.OCCUR_TIME.ToString, Info.Occur_Time,
			IdxColumnName.MAINTENANCE_MESSAGE.ToString, Info.Maintenance_Message,
															IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
														 IdxColumnName.FUCTION_ID.ToString, Info.FUCTION_ID
			)
			Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsLineInfo) As String
    Try
      Dim strSQL As String = ""
			strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' AND {14}='{15}' AND {16}='{17}'",
			strSQL,
			TableName,
			IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
			IdxColumnName.AREA_NO.ToString, Info.Area_No,
			IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
			IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
			IdxColumnName.OCCUR_TIME.ToString, Info.Occur_Time,
			IdxColumnName.MAINTENANCE_MESSAGE.ToString, Info.Maintenance_Message,
														 	IdxColumnName.MAINTENANCE_ID.ToString, Info.MAINTENANCE_ID,
														 IdxColumnName.FUCTION_ID.ToString, Info.FUCTION_ID
			)
			Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- GET
  Public Shared Function GetWMS_CT_LINE_INFODataListByALL() As List(Of clsLineInfo)
    Try
      Dim _lstReturn As New List(Of clsLineInfo)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsLineInfo = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info)
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsLineInfo, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim Factory_No = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim Area_No = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim Device_No = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim Unit_ID = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
        Dim Occur_Time = "" & RowData.Item(IdxColumnName.OCCUR_TIME.ToString)
        Dim Maintenance_Message = "" & RowData.Item(IdxColumnName.MAINTENANCE_MESSAGE.ToString)
        Dim MAINTENANCE_ID = "" & RowData.Item(IdxColumnName.MAINTENANCE_ID.ToString)
				Dim FUCTION_ID = "" & RowData.Item(IdxColumnName.FUCTION_ID.ToString)

				Info = New clsLineInfo(Factory_No, Area_No, Device_No, Unit_ID, Occur_Time, Maintenance_Message, MAINTENANCE_ID,
															 FUCTION_ID)
			End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
