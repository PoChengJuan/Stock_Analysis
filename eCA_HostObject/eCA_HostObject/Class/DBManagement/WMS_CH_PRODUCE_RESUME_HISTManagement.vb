Partial Class WMS_CH_PRODUCE_RESUME_HISTManagement
	Public Shared TableName As String = "WMS_CH_PRODUCE_RESUME_HIST"
	Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

	Enum IdxColumnName As Integer
		WO_ID
		WO_TYPE
		PO_ID
		PO_TYPE1
		PO_TYPE2
		PO_TYPE3
		RELATED_WO
		CARRIER_ID
		PACKAGE_ID
		SKU_NO
		LOT_NO
		ITEM_COMMON1
		ITEM_COMMON2
		ITEM_COMMON3
		ITEM_COMMON4
		ITEM_COMMON5
		ITEM_COMMON6
		ITEM_COMMON7
		ITEM_COMMON8
		ITEM_COMMON9
		ITEM_COMMON10
		RECEIPT_DATE
		MANUFACETURE_DATE
		EXPIRED_DATE
		EFFECTIVE_DATE
		CREATE_TIME
		QTY
		HIST_TIME
	End Enum
	'- GetSQL
	Public Shared Function GetInsertSQL(ByRef Info As clsPRODUCE_RESUME_HIST) As String
		Try

			Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56}) values ('{3}',{5},'{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}',{55},'{57}')",
       strSQL,
       TableName,
       IdxColumnName.WO_ID.ToString, Info.WO_ID,
       IdxColumnName.WO_TYPE.ToString, Info.WO_TYPE,
       IdxColumnName.PO_ID.ToString, Info.PO_ID,
       IdxColumnName.PO_TYPE1.ToString, Info.PO_TYPE1,
       IdxColumnName.PO_TYPE2.ToString, Info.PO_TYPE2,
       IdxColumnName.PO_TYPE3.ToString, Info.PO_TYPE3,
       IdxColumnName.RELATED_WO.ToString, Info.RELATED_WO,
       IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
       IdxColumnName.PACKAGE_ID.ToString, Info.PACKAGE_ID,
       IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
       IdxColumnName.LOT_NO.ToString, Info.LOT_NO,
       IdxColumnName.ITEM_COMMON1.ToString, Info.ITEM_COMMON1,
       IdxColumnName.ITEM_COMMON2.ToString, Info.ITEM_COMMON2,
       IdxColumnName.ITEM_COMMON3.ToString, Info.ITEM_COMMON3,
       IdxColumnName.ITEM_COMMON4.ToString, Info.ITEM_COMMON4,
       IdxColumnName.ITEM_COMMON5.ToString, Info.ITEM_COMMON5,
       IdxColumnName.ITEM_COMMON6.ToString, Info.ITEM_COMMON6,
       IdxColumnName.ITEM_COMMON7.ToString, Info.ITEM_COMMON7,
       IdxColumnName.ITEM_COMMON8.ToString, Info.ITEM_COMMON8,
       IdxColumnName.ITEM_COMMON9.ToString, Info.ITEM_COMMON9,
       IdxColumnName.ITEM_COMMON10.ToString, Info.ITEM_COMMON10,
       IdxColumnName.RECEIPT_DATE.ToString, Info.RECEIPT_DATE,
       IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
       IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
       IdxColumnName.EFFECTIVE_DATE.ToString, Info.EFFECTIVE_DATE,
       IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
       IdxColumnName.QTY.ToString, Info.QTY,
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
