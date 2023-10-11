Public Class WMS_CH_PRODUCE_HISTManagement
  Public Shared TableName As String = "WMS_CH_PRODUCE_HIST"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

	Enum IdxColumnName As Integer
		FACTORY_NO
		AREA_NO
		PO_ID
		SKU_NO
		STATUS
		QTY
		QTY_PROCESS
		QTY_NG
		PREVIOUS_AREA_NO
		CREATE_TIME
		START_TIME
		UPDATE_TIME
		FINISH_TIME
		HIST_TIME
		PREVIOUS_QTY_PROCESS
    PREVIOUS_QTY_NG
    PO_Info1
    PO_Info2
    PO_Info3
    PO_Info4
    PO_Info5
    PO_Info6
    PO_Info7
    PO_Info8
    PO_Info9
    PO_Info10
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsProduce_Hist) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.STATUS.ToString, CInt(Info.Status),
      IdxColumnName.QTY.ToString, Info.Qty,
      IdxColumnName.QTY_PROCESS.ToString, Info.Qty_Process,
      IdxColumnName.QTY_NG.ToString, Info.Qty_NG,
      IdxColumnName.PREVIOUS_AREA_NO.ToString, Info.Previous_Area_No,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.START_TIME.ToString, Info.Start_Time,
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.Finish_Time,
      IdxColumnName.HIST_TIME.ToString, Info.Hist_Time,
      IdxColumnName.PREVIOUS_QTY_PROCESS.ToString, Info.PREVIOUS_QTY_PROCESS,
      IdxColumnName.PREVIOUS_QTY_NG.ToString, Info.PREVIOUS_QTY_NG,
      IdxColumnName.PO_Info1.ToString, Info.PO_Info1,
      IdxColumnName.PO_Info2.ToString, Info.PO_Info2,
      IdxColumnName.PO_Info3.ToString, Info.PO_Info3,
      IdxColumnName.PO_Info4.ToString, Info.PO_Info4,
      IdxColumnName.PO_Info5.ToString, Info.PO_Info5,
      IdxColumnName.PO_Info6.ToString, Info.PO_Info6,
      IdxColumnName.PO_Info7.ToString, Info.PO_Info7,
      IdxColumnName.PO_Info8.ToString, Info.PO_Info8,
      IdxColumnName.PO_Info9.ToString, Info.PO_Info9,
      IdxColumnName.PO_Info10.ToString, Info.PO_Info10
     )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
