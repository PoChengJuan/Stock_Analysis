Public Class WMS_CT_LINE_PRODUCTION_INFOManagement
  Public Shared TableName As String = "WMS_CT_LINE_PRODUCTION_INFO"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    QTY_PROCESS
    PREVIOUS_QTY_PROCESS
    RESET_QTY_PROCESS
    QTY_MODIFY
    PREVIOUS_QTY_MODIFY
    RESET_QTY_MODIFY
    QTY_NG
    PREVIOUS_QTY_NG
    RESET_QTY_NG
		UPDATE_TIME
		QTY_TOTAL
	End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsLineProduction_Info) As String
    Try

      Dim strSQL As String = ""
			strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}',{31})",
			strSQL,
			TableName,
			IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
			IdxColumnName.AREA_NO.ToString, Info.Area_No,
			IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
			IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
			IdxColumnName.QTY_PROCESS.ToString, Info.Qty_Process,
			IdxColumnName.PREVIOUS_QTY_PROCESS.ToString, Info.Previous_Qty_Process,
			IdxColumnName.RESET_QTY_PROCESS.ToString, Info.Reset_Qty_Process,
			IdxColumnName.QTY_MODIFY.ToString, Info.Qty_Modify,
			IdxColumnName.PREVIOUS_QTY_MODIFY.ToString, Info.Previous_Qty_Modify,
			IdxColumnName.RESET_QTY_MODIFY.ToString, Info.Reset_Qty_Modify,
			IdxColumnName.QTY_NG.ToString, Info.Qty_NG,
			IdxColumnName.PREVIOUS_QTY_NG.ToString, Info.Previous_Qty_NG,
			IdxColumnName.RESET_QTY_NG.ToString, Info.Reset_Qty_NG,
			IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time,
IdxColumnName.QTY_TOTAL.ToString, Info.QTY_TOTAL
		 )
			Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsLineProduction_Info) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.UNIT_ID.ToString, Info.Unit_ID
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsLineProduction_Info) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}={31} WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
      IdxColumnName.QTY_PROCESS.ToString, Info.Qty_Process,
      IdxColumnName.PREVIOUS_QTY_PROCESS.ToString, Info.Previous_Qty_Process,
      IdxColumnName.RESET_QTY_PROCESS.ToString, Info.Reset_Qty_Process,
      IdxColumnName.QTY_MODIFY.ToString, Info.Qty_Modify,
      IdxColumnName.PREVIOUS_QTY_MODIFY.ToString, Info.Previous_Qty_Modify,
      IdxColumnName.RESET_QTY_MODIFY.ToString, Info.Reset_Qty_Modify,
      IdxColumnName.QTY_NG.ToString, Info.Qty_NG,
      IdxColumnName.PREVIOUS_QTY_NG.ToString, Info.Previous_Qty_NG,
      IdxColumnName.RESET_QTY_NG.ToString, Info.Reset_Qty_NG,
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time,
      IdxColumnName.QTY_TOTAL.ToString, Info.QTY_TOTAL
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- GET
  Public Shared Function GetWMS_CT_LINE_PRODUCTION_INFODataListByALL() As List(Of clsLineProduction_Info)
    Try
      Dim _lstReturn As New List(Of clsLineProduction_Info)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsLineProduction_Info = Nothing
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsLineProduction_Info, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim Factory_No = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim Area_No = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim Device_No = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim Unit_ID = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
        Dim Qty_Process = If(IsNumeric(RowData.Item(IdxColumnName.QTY_PROCESS.ToString)), RowData.Item(IdxColumnName.QTY_PROCESS.ToString), 0 & RowData.Item(IdxColumnName.QTY_PROCESS.ToString))
        Dim Previous_Qty_Process = If(IsNumeric(RowData.Item(IdxColumnName.PREVIOUS_QTY_PROCESS.ToString)), RowData.Item(IdxColumnName.PREVIOUS_QTY_PROCESS.ToString), 0 & RowData.Item(IdxColumnName.PREVIOUS_QTY_PROCESS.ToString))
        Dim Reset_Qty_Process = If(IsNumeric(RowData.Item(IdxColumnName.RESET_QTY_PROCESS.ToString)), RowData.Item(IdxColumnName.RESET_QTY_PROCESS.ToString), 0 & RowData.Item(IdxColumnName.RESET_QTY_PROCESS.ToString))
        Dim Qty_Modify = If(IsNumeric(RowData.Item(IdxColumnName.QTY_MODIFY.ToString)), RowData.Item(IdxColumnName.QTY_MODIFY.ToString), 0 & RowData.Item(IdxColumnName.QTY_MODIFY.ToString))
        Dim Previous_Qty_Modify = If(IsNumeric(RowData.Item(IdxColumnName.PREVIOUS_QTY_MODIFY.ToString)), RowData.Item(IdxColumnName.PREVIOUS_QTY_MODIFY.ToString), 0 & RowData.Item(IdxColumnName.PREVIOUS_QTY_MODIFY.ToString))
        Dim Reset_Qty_Modify = If(IsNumeric(RowData.Item(IdxColumnName.RESET_QTY_MODIFY.ToString)), RowData.Item(IdxColumnName.RESET_QTY_MODIFY.ToString), 0 & RowData.Item(IdxColumnName.RESET_QTY_MODIFY.ToString))
        Dim Qty_NG = If(IsNumeric(RowData.Item(IdxColumnName.QTY_NG.ToString)), RowData.Item(IdxColumnName.QTY_NG.ToString), 0 & RowData.Item(IdxColumnName.QTY_NG.ToString))
        Dim Previous_Qty_NG = If(IsNumeric(RowData.Item(IdxColumnName.PREVIOUS_QTY_NG.ToString)), RowData.Item(IdxColumnName.PREVIOUS_QTY_NG.ToString), 0 & RowData.Item(IdxColumnName.PREVIOUS_QTY_NG.ToString))
        Dim Reset_Qty_NG = If(IsNumeric(RowData.Item(IdxColumnName.RESET_QTY_NG.ToString)), RowData.Item(IdxColumnName.RESET_QTY_NG.ToString), 0 & RowData.Item(IdxColumnName.RESET_QTY_NG.ToString))
        Dim Update_Time = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim QTY_TOTAL = If(IsNumeric(RowData.Item(IdxColumnName.QTY_TOTAL.ToString)), RowData.Item(IdxColumnName.QTY_TOTAL.ToString), 0 & RowData.Item(IdxColumnName.QTY_TOTAL.ToString))
        Info = New clsLineProduction_Info(Factory_No, Area_No, Device_No, Unit_ID, Qty_Process, Previous_Qty_Process,
																					Reset_Qty_Process, Qty_Modify, Previous_Qty_Modify, Reset_Qty_Modify, Qty_NG, Previous_Qty_NG,
																					Reset_Qty_NG, Update_Time, QTY_TOTAL)
			End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
