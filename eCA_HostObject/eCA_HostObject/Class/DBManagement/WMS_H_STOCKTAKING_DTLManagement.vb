Partial Class WMS_H_STOCKTAKING_DTLManagement
Public Shared TableName As String = "WMS_H_STOCKTAKING_DTL"
Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
 
Enum IdxColumnName As Integer
STOCKTAKING_ID
STOCKTAKING_SERIAL_NO
AREA_NO
BLOCK_NO
SKU_NO
OWNER_NO
SUB_OWNER_NO
SL_NO
STORAGE_TYPE
BND
CARRIER_ID
PERCENTAGE
CARRIER_QTY
CARRIER_QTY_CHECKED
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
SORT_ITEM_COMMON1
SORT_ITEM_COMMON2
SORT_ITEM_COMMON3
SORT_ITEM_COMMON4
SORT_ITEM_COMMON5
SUPPLIER_NO
CUSTOMER_NO
RECEIPT_DATE
MANUFACETURE_DATE
EXPIRED_DATE
ERP_QTY
HIST_TIME
End Enum
'- GetSQL
 Public Shared Function GetInsertSQL(ByRef Info As  clsHSTOCKTAKINGDTL) As String
	Try
 
 Dim strSQL As String = ""
 strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}',{25},{27},{29},'{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}',{73},'{75}')",
 strSQL,
 TableName,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_SERIAL_NO.ToString, Info.STOCKTAKING_SERIAL_NO,
 IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
 IdxColumnName.BLOCK_NO.ToString, Info.BLOCK_NO,
 IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
 IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
 IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
 IdxColumnName.SL_NO.ToString, Info.SL_NO,
 IdxColumnName.STORAGE_TYPE.ToString, Info.STORAGE_TYPE,
 IdxColumnName.BND.ToString, Info.BND,
 IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
 IdxColumnName.PERCENTAGE.ToString, Info.PERCENTAGE,
 IdxColumnName.CARRIER_QTY.ToString, Info.CARRIER_QTY,
 IdxColumnName.CARRIER_QTY_CHECKED.ToString, Info.CARRIER_QTY_CHECKED,
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
 IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
 IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
 IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
 IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
 IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
 IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
 IdxColumnName.CUSTOMER_NO.ToString, Info.CUSTOMER_NO,
 IdxColumnName.RECEIPT_DATE.ToString, Info.RECEIPT_DATE,
 IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
 IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
 IdxColumnName.ERP_QTY.ToString, CINT(Info.ERP_QTY),
 IdxColumnName.HIST_TIME.ToString, Info.HIST_TIME
)
 Dim NewSQL As String = ""
 If SQLCorrect(strSQL, NewSQL) Then
 Return NewSQL
 End If
 Return Nothing
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return nothing
 End Try
 End Function
 Public Shared Function GetUpdateSQL(ByRef Info As clsHSTOCKTAKINGDTL) As String
	Try
 Dim strSQL As String = ""
 strSQL = String.Format("Update {1} SET {6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}={25},{26}={27},{28}={29},{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}={73},{74}='{75}' WHERE {2}='{3}' And {4}='{5}'",
 strSQL,
 TableName,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_SERIAL_NO.ToString, Info.STOCKTAKING_SERIAL_NO,
 IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
 IdxColumnName.BLOCK_NO.ToString, Info.BLOCK_NO,
 IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
 IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
 IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
 IdxColumnName.SL_NO.ToString, Info.SL_NO,
 IdxColumnName.STORAGE_TYPE.ToString, Info.STORAGE_TYPE,
 IdxColumnName.BND.ToString, Info.BND,
 IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
 IdxColumnName.PERCENTAGE.ToString, Info.PERCENTAGE,
 IdxColumnName.CARRIER_QTY.ToString, Info.CARRIER_QTY,
 IdxColumnName.CARRIER_QTY_CHECKED.ToString, Info.CARRIER_QTY_CHECKED,
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
 IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
 IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
 IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
 IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
 IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
 IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
 IdxColumnName.CUSTOMER_NO.ToString, Info.CUSTOMER_NO,
 IdxColumnName.RECEIPT_DATE.ToString, Info.RECEIPT_DATE,
 IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
 IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
 IdxColumnName.ERP_QTY.ToString, CINT(Info.ERP_QTY),
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
 Public Shared Function GetDeleteSQL(ByRef Info As clsHSTOCKTAKINGDTL) As String
	Try
 Dim strSQL As String = ""
 strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' ",
 strSQL,
 TableName,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_SERIAL_NO.ToString, Info.STOCKTAKING_SERIAL_NO,
 IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
 IdxColumnName.BLOCK_NO.ToString, Info.BLOCK_NO,
 IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
 IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
 IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
 IdxColumnName.SL_NO.ToString, Info.SL_NO,
 IdxColumnName.STORAGE_TYPE.ToString, Info.STORAGE_TYPE,
 IdxColumnName.BND.ToString, Info.BND,
 IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
 IdxColumnName.PERCENTAGE.ToString, Info.PERCENTAGE,
 IdxColumnName.CARRIER_QTY.ToString, Info.CARRIER_QTY,
 IdxColumnName.CARRIER_QTY_CHECKED.ToString, Info.CARRIER_QTY_CHECKED,
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
 IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
 IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
 IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
 IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
 IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
 IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
 IdxColumnName.CUSTOMER_NO.ToString, Info.CUSTOMER_NO,
 IdxColumnName.RECEIPT_DATE.ToString, Info.RECEIPT_DATE,
 IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
 IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
 IdxColumnName.ERP_QTY.ToString, CINT(Info.ERP_QTY),
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
Private Shared Function SetInfoFromDB(ByRef Info As clsHSTOCKTAKINGDTL , ByRef RowData As DataRow) As Boolean
 Try
If RowData IsNot Nothing Then
Dim STOCKTAKING_ID=""&RowData.Item(IdxColumnName.STOCKTAKING_ID.ToString)
Dim STOCKTAKING_SERIAL_NO=""&RowData.Item(IdxColumnName.STOCKTAKING_SERIAL_NO.ToString)
Dim AREA_NO=""&RowData.Item(IdxColumnName.AREA_NO.ToString)
Dim BLOCK_NO=""&RowData.Item(IdxColumnName.BLOCK_NO.ToString)
Dim SKU_NO=""&RowData.Item(IdxColumnName.SKU_NO.ToString)
Dim OWNER_NO=""&RowData.Item(IdxColumnName.OWNER_NO.ToString)
Dim SUB_OWNER_NO=""&RowData.Item(IdxColumnName.SUB_OWNER_NO.ToString)
Dim SL_NO=""&RowData.Item(IdxColumnName.SL_NO.ToString)
Dim STORAGE_TYPE=""&RowData.Item(IdxColumnName.STORAGE_TYPE.ToString)
Dim BND=""&RowData.Item(IdxColumnName.BND.ToString)
Dim CARRIER_ID=""&RowData.Item(IdxColumnName.CARRIER_ID.ToString)
Dim PERCENTAGE=If(IsNumeric(RowData.Item(IdxColumnName.PERCENTAGE.ToString))  ,RowData.Item(IdxColumnName.PERCENTAGE.ToString), 0 & RowData.Item(IdxColumnName.PERCENTAGE.ToString))
Dim CARRIER_QTY=If(IsNumeric(RowData.Item(IdxColumnName.CARRIER_QTY.ToString))  ,RowData.Item(IdxColumnName.CARRIER_QTY.ToString), 0 & RowData.Item(IdxColumnName.CARRIER_QTY.ToString))
Dim CARRIER_QTY_CHECKED=If(IsNumeric(RowData.Item(IdxColumnName.CARRIER_QTY_CHECKED.ToString))  ,RowData.Item(IdxColumnName.CARRIER_QTY_CHECKED.ToString), 0 & RowData.Item(IdxColumnName.CARRIER_QTY_CHECKED.ToString))
Dim LOT_NO=""&RowData.Item(IdxColumnName.LOT_NO.ToString)
Dim ITEM_COMMON1=""&RowData.Item(IdxColumnName.ITEM_COMMON1.ToString)
Dim ITEM_COMMON2=""&RowData.Item(IdxColumnName.ITEM_COMMON2.ToString)
Dim ITEM_COMMON3=""&RowData.Item(IdxColumnName.ITEM_COMMON3.ToString)
Dim ITEM_COMMON4=""&RowData.Item(IdxColumnName.ITEM_COMMON4.ToString)
Dim ITEM_COMMON5=""&RowData.Item(IdxColumnName.ITEM_COMMON5.ToString)
Dim ITEM_COMMON6=""&RowData.Item(IdxColumnName.ITEM_COMMON6.ToString)
Dim ITEM_COMMON7=""&RowData.Item(IdxColumnName.ITEM_COMMON7.ToString)
Dim ITEM_COMMON8=""&RowData.Item(IdxColumnName.ITEM_COMMON8.ToString)
Dim ITEM_COMMON9=""&RowData.Item(IdxColumnName.ITEM_COMMON9.ToString)
Dim ITEM_COMMON10=""&RowData.Item(IdxColumnName.ITEM_COMMON10.ToString)
Dim SORT_ITEM_COMMON1=""&RowData.Item(IdxColumnName.SORT_ITEM_COMMON1.ToString)
Dim SORT_ITEM_COMMON2=""&RowData.Item(IdxColumnName.SORT_ITEM_COMMON2.ToString)
Dim SORT_ITEM_COMMON3=""&RowData.Item(IdxColumnName.SORT_ITEM_COMMON3.ToString)
Dim SORT_ITEM_COMMON4=""&RowData.Item(IdxColumnName.SORT_ITEM_COMMON4.ToString)
Dim SORT_ITEM_COMMON5=""&RowData.Item(IdxColumnName.SORT_ITEM_COMMON5.ToString)
Dim SUPPLIER_NO=""&RowData.Item(IdxColumnName.SUPPLIER_NO.ToString)
Dim CUSTOMER_NO=""&RowData.Item(IdxColumnName.CUSTOMER_NO.ToString)
Dim RECEIPT_DATE=""&RowData.Item(IdxColumnName.RECEIPT_DATE.ToString)
Dim MANUFACETURE_DATE=""&RowData.Item(IdxColumnName.MANUFACETURE_DATE.ToString)
Dim EXPIRED_DATE=""&RowData.Item(IdxColumnName.EXPIRED_DATE.ToString)
Dim ERP_QTY=If(IsNumeric(RowData.Item(IdxColumnName.ERP_QTY.ToString))  ,RowData.Item(IdxColumnName.ERP_QTY.ToString), 0 & RowData.Item(IdxColumnName.ERP_QTY.ToString))
Dim HIST_TIME=""&RowData.Item(IdxColumnName.HIST_TIME.ToString)
 Info = New clsHSTOCKTAKINGDTL(STOCKTAKING_ID,STOCKTAKING_SERIAL_NO,AREA_NO,BLOCK_NO,SKU_NO,OWNER_NO,SUB_OWNER_NO,SL_NO,STORAGE_TYPE,BND,CARRIER_ID,PERCENTAGE,CARRIER_QTY,CARRIER_QTY_CHECKED,LOT_NO,ITEM_COMMON1,ITEM_COMMON2,ITEM_COMMON3,ITEM_COMMON4,ITEM_COMMON5,ITEM_COMMON6,ITEM_COMMON7,ITEM_COMMON8,ITEM_COMMON9,ITEM_COMMON10,SORT_ITEM_COMMON1,SORT_ITEM_COMMON2,SORT_ITEM_COMMON3,SORT_ITEM_COMMON4,SORT_ITEM_COMMON5,SUPPLIER_NO,CUSTOMER_NO,RECEIPT_DATE,MANUFACETURE_DATE,EXPIRED_DATE,ERP_QTY,HIST_TIME)
 
End If
Return true
Catch ex As Exception
SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
Return false
End Try
End Function
  Public Shared Function GetWMS_H_STOCKTAKING_DTLdicByStocktaking_ID(ByVal Stocktaking_id As String) As Dictionary(Of String, clsHSTOCKTAKINGDTL)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsHSTOCKTAKINGDTL)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} where {2}='{3}' ",
 strSQL,
 TableName,
 IdxColumnName.STOCKTAKING_ID.ToString,
 Stocktaking_id
 )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsHSTOCKTAKINGDTL = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info.gid, Info)
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
