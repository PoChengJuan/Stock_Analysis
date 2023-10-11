Partial Class WMS_H_STOCKTAKING_CARRIERManagement
Public Shared TableName As String = "WMS_H_STOCKTAKING_CARRIER"
Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
 
Enum IdxColumnName As Integer
KEY_NO
STOCKTAKING_ID
STOCKTAKING_SERIAL_NO
CARRIER_ID
QTY
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
SORT_ITEM_COMMON1
SORT_ITEM_COMMON2
SORT_ITEM_COMMON3
SORT_ITEM_COMMON4
SORT_ITEM_COMMON5
OWNER_NO
SUB_OWNER_NO
SL_NO
STORAGE_TYPE
BND
QC_STATUS
MANUFACETURE_DATE
EXPIRED_DATE
REPORT_QTY
REPORT_PACKAGE_ID
REPORT_SKU_NO
REPORT_LOT_NO
REPORT_ITEM_COMMON1
REPORT_ITEM_COMMON2
REPORT_ITEM_COMMON3
REPORT_ITEM_COMMON4
REPORT_ITEM_COMMON5
REPORT_ITEM_COMMON6
REPORT_ITEM_COMMON7
REPORT_ITEM_COMMON8
REPORT_ITEM_COMMON9
REPORT_ITEM_COMMON10
REPORT_SORT_ITEM_COMMON1
REPORT_SORT_ITEM_COMMON2
REPORT_SORT_ITEM_COMMON3
REPORT_SORT_ITEM_COMMON4
REPORT_SORT_ITEM_COMMON5
REPORT_OWNER_NO
REPORT_SUB_OWNER_NO
REPORT_SL_NO
REPORT_STORAGE_TYPE
REPORT_BND
REPORT_QC_STATUS
REPORT_MANUFACETURE_DATE
REPORT_EXPIRED_DATE
STOCKTAKING_COUNT
STOCKTAKING_STATUS
PROFITP_TYPE
REPORT_USER
DEST_LOCATION_NO
ACTUAL_AREA_NO
ACTUAL_LOCATION_NO
ACTUAL_SUBLOCATION_X
ACTUAL_SUBLOCATION_Y
ACTUAL_SUBLOCATION_Z
MANUAL_KEY
HIST_TIME
End Enum
'- GetSQL
 Public Shared Function GetInsertSQL(ByRef Info As  clsHSTOCKTAKINGCARRIER) As String
	Try
 
 Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92},{94},{96},{98},{100},{102},{104},{106},{108},{110},{112},{114},{116},{118},{120},{122},{124},{126},{128},{130},{132},{134},{136},{138},{140}) values ('{3}','{5}','{7}','{9}',{11},'{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}',{55},{57},{59},'{61}','{63}',{65},'{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}','{95}','{97}','{99}','{101}','{103}','{105}','{107}',{109},{111},{113},'{115}','{117}',{119},{121},{123},'{125}','{127}','{129}','{131}','{133}','{135}','{137}','{139}','{141}')",
 strSQL,
 TableName,
 IdxColumnName.KEY_NO.ToString, Info.KEY_NO,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_SERIAL_NO.ToString, Info.STOCKTAKING_SERIAL_NO,
 IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
 IdxColumnName.QTY.ToString, Info.QTY,
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
 IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
 IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
 IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
 IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
 IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
 IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
 IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
 IdxColumnName.SL_NO.ToString, Info.SL_NO,
 IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.STORAGE_TYPE),
 IdxColumnName.BND.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.BND),
 IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
 IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
 IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
 IdxColumnName.REPORT_QTY.ToString, Info.REPORT_QTY,
 IdxColumnName.REPORT_PACKAGE_ID.ToString, Info.REPORT_PACKAGE_ID,
 IdxColumnName.REPORT_SKU_NO.ToString, Info.REPORT_SKU_NO,
 IdxColumnName.REPORT_LOT_NO.ToString, Info.REPORT_LOT_NO,
 IdxColumnName.REPORT_ITEM_COMMON1.ToString, Info.REPORT_ITEM_COMMON1,
 IdxColumnName.REPORT_ITEM_COMMON2.ToString, Info.REPORT_ITEM_COMMON2,
 IdxColumnName.REPORT_ITEM_COMMON3.ToString, Info.REPORT_ITEM_COMMON3,
 IdxColumnName.REPORT_ITEM_COMMON4.ToString, Info.REPORT_ITEM_COMMON4,
 IdxColumnName.REPORT_ITEM_COMMON5.ToString, Info.REPORT_ITEM_COMMON5,
 IdxColumnName.REPORT_ITEM_COMMON6.ToString, Info.REPORT_ITEM_COMMON6,
 IdxColumnName.REPORT_ITEM_COMMON7.ToString, Info.REPORT_ITEM_COMMON7,
 IdxColumnName.REPORT_ITEM_COMMON8.ToString, Info.REPORT_ITEM_COMMON8,
 IdxColumnName.REPORT_ITEM_COMMON9.ToString, Info.REPORT_ITEM_COMMON9,
 IdxColumnName.REPORT_ITEM_COMMON10.ToString, Info.REPORT_ITEM_COMMON10,
 IdxColumnName.REPORT_SORT_ITEM_COMMON1.ToString, Info.REPORT_SORT_ITEM_COMMON1,
 IdxColumnName.REPORT_SORT_ITEM_COMMON2.ToString, Info.REPORT_SORT_ITEM_COMMON2,
 IdxColumnName.REPORT_SORT_ITEM_COMMON3.ToString, Info.REPORT_SORT_ITEM_COMMON3,
 IdxColumnName.REPORT_SORT_ITEM_COMMON4.ToString, Info.REPORT_SORT_ITEM_COMMON4,
 IdxColumnName.REPORT_SORT_ITEM_COMMON5.ToString, Info.REPORT_SORT_ITEM_COMMON5,
 IdxColumnName.REPORT_OWNER_NO.ToString, Info.REPORT_OWNER_NO,
 IdxColumnName.REPORT_SUB_OWNER_NO.ToString, Info.REPORT_SUB_OWNER_NO,
 IdxColumnName.REPORT_SL_NO.ToString, Info.REPORT_SL_NO,
 IdxColumnName.REPORT_STORAGE_TYPE.ToString, CInt(Info.REPORT_STORAGE_TYPE),
 IdxColumnName.REPORT_BND.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.REPORT_BND),
 IdxColumnName.REPORT_QC_STATUS.ToString, CInt(Info.REPORT_QC_STATUS),
 IdxColumnName.REPORT_MANUFACETURE_DATE.ToString, Info.REPORT_MANUFACETURE_DATE,
 IdxColumnName.REPORT_EXPIRED_DATE.ToString, Info.REPORT_EXPIRED_DATE,
 IdxColumnName.STOCKTAKING_COUNT.ToString, Info.STOCKTAKING_COUNT,
 IdxColumnName.STOCKTAKING_STATUS.ToString, Info.STOCKTAKING_STATUS,
 IdxColumnName.PROFITP_TYPE.ToString, CInt(Info.PROFITP_TYPE),
 IdxColumnName.REPORT_USER.ToString, Info.REPORT_USER,
 IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_NO,
 IdxColumnName.ACTUAL_AREA_NO.ToString, Info.ACTUAL_AREA_NO,
 IdxColumnName.ACTUAL_LOCATION_NO.ToString, Info.ACTUAL_LOCATION_NO,
 IdxColumnName.ACTUAL_SUBLOCATION_X.ToString, Info.ACTUAL_SUBLOCATION_X,
 IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString, Info.ACTUAL_SUBLOCATION_Y,
 IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString, Info.ACTUAL_SUBLOCATION_Z,
 IdxColumnName.MANUAL_KEY.ToString, Info.MANUAL_KEY,
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
 Public Shared Function GetUpdateSQL(ByRef Info As clsHSTOCKTAKINGCARRIER) As String
	Try
 Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}={11},{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}={55},{56}={57},{58}={59},{60}='{61}',{62}='{63}',{64}={65},{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}',{94}='{95}',{96}='{97}',{98}='{99}',{100}='{101}',{102}='{103}',{104}='{105}',{106}='{107}',{108}={109},{110}={111},{112}={113},{114}='{115}',{116}='{117}',{118}={119},{120}={121},{122}={123},{124}='{125}',{126}='{127}',{128}='{129}',{130}='{131}',{132}='{133}',{134}='{135}',{136}='{137}',{138}='{139}',{140}='{141}' WHERE {2}='{3}'",
 strSQL,
 TableName,
 IdxColumnName.KEY_NO.ToString, Info.KEY_NO,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_SERIAL_NO.ToString, Info.STOCKTAKING_SERIAL_NO,
 IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
 IdxColumnName.QTY.ToString, Info.QTY,
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
 IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
 IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
 IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
 IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
 IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
 IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
 IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
 IdxColumnName.SL_NO.ToString, Info.SL_NO,
 IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.STORAGE_TYPE),
 IdxColumnName.BND.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.BND),
 IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
 IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
 IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
 IdxColumnName.REPORT_QTY.ToString, Info.REPORT_QTY,
 IdxColumnName.REPORT_PACKAGE_ID.ToString, Info.REPORT_PACKAGE_ID,
 IdxColumnName.REPORT_SKU_NO.ToString, Info.REPORT_SKU_NO,
 IdxColumnName.REPORT_LOT_NO.ToString, Info.REPORT_LOT_NO,
 IdxColumnName.REPORT_ITEM_COMMON1.ToString, Info.REPORT_ITEM_COMMON1,
 IdxColumnName.REPORT_ITEM_COMMON2.ToString, Info.REPORT_ITEM_COMMON2,
 IdxColumnName.REPORT_ITEM_COMMON3.ToString, Info.REPORT_ITEM_COMMON3,
 IdxColumnName.REPORT_ITEM_COMMON4.ToString, Info.REPORT_ITEM_COMMON4,
 IdxColumnName.REPORT_ITEM_COMMON5.ToString, Info.REPORT_ITEM_COMMON5,
 IdxColumnName.REPORT_ITEM_COMMON6.ToString, Info.REPORT_ITEM_COMMON6,
 IdxColumnName.REPORT_ITEM_COMMON7.ToString, Info.REPORT_ITEM_COMMON7,
 IdxColumnName.REPORT_ITEM_COMMON8.ToString, Info.REPORT_ITEM_COMMON8,
 IdxColumnName.REPORT_ITEM_COMMON9.ToString, Info.REPORT_ITEM_COMMON9,
 IdxColumnName.REPORT_ITEM_COMMON10.ToString, Info.REPORT_ITEM_COMMON10,
 IdxColumnName.REPORT_SORT_ITEM_COMMON1.ToString, Info.REPORT_SORT_ITEM_COMMON1,
 IdxColumnName.REPORT_SORT_ITEM_COMMON2.ToString, Info.REPORT_SORT_ITEM_COMMON2,
 IdxColumnName.REPORT_SORT_ITEM_COMMON3.ToString, Info.REPORT_SORT_ITEM_COMMON3,
 IdxColumnName.REPORT_SORT_ITEM_COMMON4.ToString, Info.REPORT_SORT_ITEM_COMMON4,
 IdxColumnName.REPORT_SORT_ITEM_COMMON5.ToString, Info.REPORT_SORT_ITEM_COMMON5,
 IdxColumnName.REPORT_OWNER_NO.ToString, Info.REPORT_OWNER_NO,
 IdxColumnName.REPORT_SUB_OWNER_NO.ToString, Info.REPORT_SUB_OWNER_NO,
 IdxColumnName.REPORT_SL_NO.ToString, Info.REPORT_SL_NO,
 IdxColumnName.REPORT_STORAGE_TYPE.ToString, CInt(Info.REPORT_STORAGE_TYPE),
 IdxColumnName.REPORT_BND.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.REPORT_BND),
 IdxColumnName.REPORT_QC_STATUS.ToString, CInt(Info.REPORT_QC_STATUS),
 IdxColumnName.REPORT_MANUFACETURE_DATE.ToString, Info.REPORT_MANUFACETURE_DATE,
 IdxColumnName.REPORT_EXPIRED_DATE.ToString, Info.REPORT_EXPIRED_DATE,
 IdxColumnName.STOCKTAKING_COUNT.ToString, Info.STOCKTAKING_COUNT,
 IdxColumnName.STOCKTAKING_STATUS.ToString, Info.STOCKTAKING_STATUS,
 IdxColumnName.PROFITP_TYPE.ToString, CInt(Info.PROFITP_TYPE),
 IdxColumnName.REPORT_USER.ToString, Info.REPORT_USER,
 IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_NO,
 IdxColumnName.ACTUAL_AREA_NO.ToString, Info.ACTUAL_AREA_NO,
 IdxColumnName.ACTUAL_LOCATION_NO.ToString, Info.ACTUAL_LOCATION_NO,
 IdxColumnName.ACTUAL_SUBLOCATION_X.ToString, Info.ACTUAL_SUBLOCATION_X,
 IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString, Info.ACTUAL_SUBLOCATION_Y,
 IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString, Info.ACTUAL_SUBLOCATION_Z,
 IdxColumnName.MANUAL_KEY.ToString, Info.MANUAL_KEY,
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsHSTOCKTAKINGCARRIER) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
 strSQL,
 TableName,
 IdxColumnName.KEY_NO.ToString, Info.KEY_NO,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_SERIAL_NO.ToString, Info.STOCKTAKING_SERIAL_NO,
 IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
 IdxColumnName.QTY.ToString, Info.QTY,
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
 IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
 IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
 IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
 IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
 IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
 IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
 IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
 IdxColumnName.SL_NO.ToString, Info.SL_NO,
 IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.STORAGE_TYPE),
 IdxColumnName.BND.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.BND),
 IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
 IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
 IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
 IdxColumnName.REPORT_QTY.ToString, Info.REPORT_QTY,
 IdxColumnName.REPORT_PACKAGE_ID.ToString, Info.REPORT_PACKAGE_ID,
 IdxColumnName.REPORT_SKU_NO.ToString, Info.REPORT_SKU_NO,
 IdxColumnName.REPORT_LOT_NO.ToString, Info.REPORT_LOT_NO,
 IdxColumnName.REPORT_ITEM_COMMON1.ToString, Info.REPORT_ITEM_COMMON1,
 IdxColumnName.REPORT_ITEM_COMMON2.ToString, Info.REPORT_ITEM_COMMON2,
 IdxColumnName.REPORT_ITEM_COMMON3.ToString, Info.REPORT_ITEM_COMMON3,
 IdxColumnName.REPORT_ITEM_COMMON4.ToString, Info.REPORT_ITEM_COMMON4,
 IdxColumnName.REPORT_ITEM_COMMON5.ToString, Info.REPORT_ITEM_COMMON5,
 IdxColumnName.REPORT_ITEM_COMMON6.ToString, Info.REPORT_ITEM_COMMON6,
 IdxColumnName.REPORT_ITEM_COMMON7.ToString, Info.REPORT_ITEM_COMMON7,
 IdxColumnName.REPORT_ITEM_COMMON8.ToString, Info.REPORT_ITEM_COMMON8,
 IdxColumnName.REPORT_ITEM_COMMON9.ToString, Info.REPORT_ITEM_COMMON9,
 IdxColumnName.REPORT_ITEM_COMMON10.ToString, Info.REPORT_ITEM_COMMON10,
 IdxColumnName.REPORT_SORT_ITEM_COMMON1.ToString, Info.REPORT_SORT_ITEM_COMMON1,
 IdxColumnName.REPORT_SORT_ITEM_COMMON2.ToString, Info.REPORT_SORT_ITEM_COMMON2,
 IdxColumnName.REPORT_SORT_ITEM_COMMON3.ToString, Info.REPORT_SORT_ITEM_COMMON3,
 IdxColumnName.REPORT_SORT_ITEM_COMMON4.ToString, Info.REPORT_SORT_ITEM_COMMON4,
 IdxColumnName.REPORT_SORT_ITEM_COMMON5.ToString, Info.REPORT_SORT_ITEM_COMMON5,
 IdxColumnName.REPORT_OWNER_NO.ToString, Info.REPORT_OWNER_NO,
 IdxColumnName.REPORT_SUB_OWNER_NO.ToString, Info.REPORT_SUB_OWNER_NO,
 IdxColumnName.REPORT_SL_NO.ToString, Info.REPORT_SL_NO,
 IdxColumnName.REPORT_STORAGE_TYPE.ToString, CInt(Info.REPORT_STORAGE_TYPE),
 IdxColumnName.REPORT_BND.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.REPORT_BND),
 IdxColumnName.REPORT_QC_STATUS.ToString, CInt(Info.REPORT_QC_STATUS),
 IdxColumnName.REPORT_MANUFACETURE_DATE.ToString, Info.REPORT_MANUFACETURE_DATE,
 IdxColumnName.REPORT_EXPIRED_DATE.ToString, Info.REPORT_EXPIRED_DATE,
 IdxColumnName.STOCKTAKING_COUNT.ToString, Info.STOCKTAKING_COUNT,
 IdxColumnName.STOCKTAKING_STATUS.ToString, Info.STOCKTAKING_STATUS,
 IdxColumnName.PROFITP_TYPE.ToString, CInt(Info.PROFITP_TYPE),
 IdxColumnName.REPORT_USER.ToString, Info.REPORT_USER,
 IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_NO,
 IdxColumnName.ACTUAL_AREA_NO.ToString, Info.ACTUAL_AREA_NO,
 IdxColumnName.ACTUAL_LOCATION_NO.ToString, Info.ACTUAL_LOCATION_NO,
 IdxColumnName.ACTUAL_SUBLOCATION_X.ToString, Info.ACTUAL_SUBLOCATION_X,
 IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString, Info.ACTUAL_SUBLOCATION_Y,
 IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString, Info.ACTUAL_SUBLOCATION_Z,
 IdxColumnName.MANUAL_KEY.ToString, Info.MANUAL_KEY,
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
Private Shared Function SetInfoFromDB(ByRef Info As clsHSTOCKTAKINGCARRIER , ByRef RowData As DataRow) As Boolean
 Try
If RowData IsNot Nothing Then
Dim KEY_NO=""&RowData.Item(IdxColumnName.KEY_NO.ToString)
Dim STOCKTAKING_ID=""&RowData.Item(IdxColumnName.STOCKTAKING_ID.ToString)
Dim STOCKTAKING_SERIAL_NO=""&RowData.Item(IdxColumnName.STOCKTAKING_SERIAL_NO.ToString)
Dim CARRIER_ID=""&RowData.Item(IdxColumnName.CARRIER_ID.ToString)
Dim QTY=If(IsNumeric(RowData.Item(IdxColumnName.QTY.ToString))  ,RowData.Item(IdxColumnName.QTY.ToString), 0 & RowData.Item(IdxColumnName.QTY.ToString))
Dim PACKAGE_ID=""&RowData.Item(IdxColumnName.PACKAGE_ID.ToString)
Dim SKU_NO=""&RowData.Item(IdxColumnName.SKU_NO.ToString)
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
Dim OWNER_NO=""&RowData.Item(IdxColumnName.OWNER_NO.ToString)
Dim SUB_OWNER_NO=""&RowData.Item(IdxColumnName.SUB_OWNER_NO.ToString)
Dim SL_NO=""&RowData.Item(IdxColumnName.SL_NO.ToString)
Dim STORAGE_TYPE=If(IsNumeric(RowData.Item(IdxColumnName.STORAGE_TYPE.ToString))  ,RowData.Item(IdxColumnName.STORAGE_TYPE.ToString), 0 & RowData.Item(IdxColumnName.STORAGE_TYPE.ToString))
Dim BND=IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.BND.ToString))
Dim QC_STATUS=If(IsNumeric(RowData.Item(IdxColumnName.QC_STATUS.ToString))  ,RowData.Item(IdxColumnName.QC_STATUS.ToString), 0 & RowData.Item(IdxColumnName.QC_STATUS.ToString))
Dim MANUFACETURE_DATE=""&RowData.Item(IdxColumnName.MANUFACETURE_DATE.ToString)
Dim EXPIRED_DATE=""&RowData.Item(IdxColumnName.EXPIRED_DATE.ToString)
Dim REPORT_QTY=If(IsNumeric(RowData.Item(IdxColumnName.REPORT_QTY.ToString))  ,RowData.Item(IdxColumnName.REPORT_QTY.ToString), 0 & RowData.Item(IdxColumnName.REPORT_QTY.ToString))
Dim REPORT_PACKAGE_ID=""&RowData.Item(IdxColumnName.REPORT_PACKAGE_ID.ToString)
Dim REPORT_SKU_NO=""&RowData.Item(IdxColumnName.REPORT_SKU_NO.ToString)
Dim REPORT_LOT_NO=""&RowData.Item(IdxColumnName.REPORT_LOT_NO.ToString)
Dim REPORT_ITEM_COMMON1=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON1.ToString)
Dim REPORT_ITEM_COMMON2=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON2.ToString)
Dim REPORT_ITEM_COMMON3=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON3.ToString)
Dim REPORT_ITEM_COMMON4=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON4.ToString)
Dim REPORT_ITEM_COMMON5=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON5.ToString)
Dim REPORT_ITEM_COMMON6=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON6.ToString)
Dim REPORT_ITEM_COMMON7=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON7.ToString)
Dim REPORT_ITEM_COMMON8=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON8.ToString)
Dim REPORT_ITEM_COMMON9=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON9.ToString)
Dim REPORT_ITEM_COMMON10=""&RowData.Item(IdxColumnName.REPORT_ITEM_COMMON10.ToString)
Dim REPORT_SORT_ITEM_COMMON1=""&RowData.Item(IdxColumnName.REPORT_SORT_ITEM_COMMON1.ToString)
Dim REPORT_SORT_ITEM_COMMON2=""&RowData.Item(IdxColumnName.REPORT_SORT_ITEM_COMMON2.ToString)
Dim REPORT_SORT_ITEM_COMMON3=""&RowData.Item(IdxColumnName.REPORT_SORT_ITEM_COMMON3.ToString)
Dim REPORT_SORT_ITEM_COMMON4=""&RowData.Item(IdxColumnName.REPORT_SORT_ITEM_COMMON4.ToString)
Dim REPORT_SORT_ITEM_COMMON5=""&RowData.Item(IdxColumnName.REPORT_SORT_ITEM_COMMON5.ToString)
Dim REPORT_OWNER_NO=""&RowData.Item(IdxColumnName.REPORT_OWNER_NO.ToString)
Dim REPORT_SUB_OWNER_NO=""&RowData.Item(IdxColumnName.REPORT_SUB_OWNER_NO.ToString)
Dim REPORT_SL_NO=""&RowData.Item(IdxColumnName.REPORT_SL_NO.ToString)
Dim REPORT_STORAGE_TYPE=If(IsNumeric(RowData.Item(IdxColumnName.REPORT_STORAGE_TYPE.ToString))  ,RowData.Item(IdxColumnName.REPORT_STORAGE_TYPE.ToString), 0 & RowData.Item(IdxColumnName.REPORT_STORAGE_TYPE.ToString))
Dim REPORT_BND=IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.REPORT_BND.ToString))
Dim REPORT_QC_STATUS=If(IsNumeric(RowData.Item(IdxColumnName.REPORT_QC_STATUS.ToString))  ,RowData.Item(IdxColumnName.REPORT_QC_STATUS.ToString), 0 & RowData.Item(IdxColumnName.REPORT_QC_STATUS.ToString))
Dim REPORT_MANUFACETURE_DATE=""&RowData.Item(IdxColumnName.REPORT_MANUFACETURE_DATE.ToString)
Dim REPORT_EXPIRED_DATE=""&RowData.Item(IdxColumnName.REPORT_EXPIRED_DATE.ToString)
Dim STOCKTAKING_COUNT=If(IsNumeric(RowData.Item(IdxColumnName.STOCKTAKING_COUNT.ToString))  ,RowData.Item(IdxColumnName.STOCKTAKING_COUNT.ToString), 0 & RowData.Item(IdxColumnName.STOCKTAKING_COUNT.ToString))
Dim STOCKTAKING_STATUS=If(IsNumeric(RowData.Item(IdxColumnName.STOCKTAKING_STATUS.ToString))  ,RowData.Item(IdxColumnName.STOCKTAKING_STATUS.ToString), 0 & RowData.Item(IdxColumnName.STOCKTAKING_STATUS.ToString))
Dim PROFITP_TYPE=If(IsNumeric(RowData.Item(IdxColumnName.PROFITP_TYPE.ToString))  ,RowData.Item(IdxColumnName.PROFITP_TYPE.ToString), 0 & RowData.Item(IdxColumnName.PROFITP_TYPE.ToString))
Dim REPORT_USER=""&RowData.Item(IdxColumnName.REPORT_USER.ToString)
Dim DEST_LOCATION_NO=""&RowData.Item(IdxColumnName.DEST_LOCATION_NO.ToString)
Dim ACTUAL_AREA_NO=""&RowData.Item(IdxColumnName.ACTUAL_AREA_NO.ToString)
Dim ACTUAL_LOCATION_NO=""&RowData.Item(IdxColumnName.ACTUAL_LOCATION_NO.ToString)
Dim ACTUAL_SUBLOCATION_X=""&RowData.Item(IdxColumnName.ACTUAL_SUBLOCATION_X.ToString)
Dim ACTUAL_SUBLOCATION_Y=""&RowData.Item(IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString)
Dim ACTUAL_SUBLOCATION_Z=""&RowData.Item(IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString)
Dim MANUAL_KEY=""&RowData.Item(IdxColumnName.MANUAL_KEY.ToString)
Dim HIST_TIME=""&RowData.Item(IdxColumnName.HIST_TIME.ToString)
 Info = New clsHSTOCKTAKINGCARRIER(KEY_NO,STOCKTAKING_ID,STOCKTAKING_SERIAL_NO,CARRIER_ID,QTY,PACKAGE_ID,SKU_NO,LOT_NO,ITEM_COMMON1,ITEM_COMMON2,ITEM_COMMON3,ITEM_COMMON4,ITEM_COMMON5,ITEM_COMMON6,ITEM_COMMON7,ITEM_COMMON8,ITEM_COMMON9,ITEM_COMMON10,SORT_ITEM_COMMON1,SORT_ITEM_COMMON2,SORT_ITEM_COMMON3,SORT_ITEM_COMMON4,SORT_ITEM_COMMON5,OWNER_NO,SUB_OWNER_NO,SL_NO,STORAGE_TYPE,BND,QC_STATUS,MANUFACETURE_DATE,EXPIRED_DATE,REPORT_QTY,REPORT_PACKAGE_ID,REPORT_SKU_NO,REPORT_LOT_NO,REPORT_ITEM_COMMON1,REPORT_ITEM_COMMON2,REPORT_ITEM_COMMON3,REPORT_ITEM_COMMON4,REPORT_ITEM_COMMON5,REPORT_ITEM_COMMON6,REPORT_ITEM_COMMON7,REPORT_ITEM_COMMON8,REPORT_ITEM_COMMON9,REPORT_ITEM_COMMON10,REPORT_SORT_ITEM_COMMON1,REPORT_SORT_ITEM_COMMON2,REPORT_SORT_ITEM_COMMON3,REPORT_SORT_ITEM_COMMON4,REPORT_SORT_ITEM_COMMON5,REPORT_OWNER_NO,REPORT_SUB_OWNER_NO,REPORT_SL_NO,REPORT_STORAGE_TYPE,REPORT_BND,REPORT_QC_STATUS,REPORT_MANUFACETURE_DATE,REPORT_EXPIRED_DATE,STOCKTAKING_COUNT,STOCKTAKING_STATUS,PROFITP_TYPE,REPORT_USER,DEST_LOCATION_NO,ACTUAL_AREA_NO,ACTUAL_LOCATION_NO,ACTUAL_SUBLOCATION_X,ACTUAL_SUBLOCATION_Y,ACTUAL_SUBLOCATION_Z,MANUAL_KEY,HIST_TIME)
 
End If
Return true
Catch ex As Exception
SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
Return false
End Try
End Function
  Public Shared Function GetWMS_H_STOCKTAKING_CARRIERdicByStocktaking_ID(ByVal Stocktaking_ID As String) As Dictionary(Of String, clsHSTOCKTAKINGCARRIER)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsHSTOCKTAKINGCARRIER)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE {2}='{3}'",
 strSQL,
 TableName,
 IdxColumnName.STOCKTAKING_ID.ToString,
 Stocktaking_ID
 )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsHSTOCKTAKINGCARRIER = Nothing
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
