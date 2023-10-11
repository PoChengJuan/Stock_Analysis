Imports System.Collections.Concurrent
Partial Class WMS_T_INBOUND_DTLManagement
  Public Shared TableName As String = "WMS_T_INBOUND_DTL"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    KEY_NO = 0
    WO_ID = 1
    WO_SERIAL_NO = 2
    CARRIER_ID = 3
    SKU_NO = 4
    QTY_INBOUND = 5
    ITEM_KEY_NO = 6
    COMMENTS = 7
    PACKAGE_ID = 8
    LOT_NO = 9
    ITEM_COMMON1 = 10
    ITEM_COMMON2 = 11
    ITEM_COMMON3 = 12
    ITEM_COMMON4 = 13
    ITEM_COMMON5 = 14
    ITEM_COMMON6 = 15
    ITEM_COMMON7 = 16
    ITEM_COMMON8 = 17
    ITEM_COMMON9 = 18
    ITEM_COMMON10 = 19
    SORT_ITEM_COMMON1 = 20
    SORT_ITEM_COMMON2 = 21
    SORT_ITEM_COMMON3 = 22
    SORT_ITEM_COMMON4 = 23
    SORT_ITEM_COMMON5 = 24
    SUPPLIER_NO = 25
    CUSTOMER_NO = 26
    OWNER_NO = 27
    SUB_OWNER_NO = 28
    LENGTH = 29
    WIDTH = 30
    HEIGHT = 31
    WEIGHT = 32
    ITEM_VALUE = 33
    CONTRACT_NO = 34
    CONTRACT_SERIAL_NO = 35
    PO_ID = 36
    PO_SERIAL_NO = 37
    STORAGE_TYPE = 38
    BND = 39
    INBOUND_STATUS = 40
    QC_STATUS = 41
    QC_TIME = 42
    INBOUND_TIME = 43
    RECEIPT_DATE = 44
    MANUFACETURE_DATE = 45
    EXPIRED_DATE = 46
    EFFECTIVE_DATE = 47
    LOCATION_NO = 48
    DEST_FACTORY_NO = 49
    DEST_AREA_NO = 50
    DEST_BLOCK_NO = 51
    DEST_LOCATION_NO = 52
    ACTUAL_AREA_NO = 53
    ACTUAL_LOCATION_NO = 54
    ACTUAL_SUBLOCATION_X = 55
    ACTUAL_SUBLOCATION_Y = 56
    ACTUAL_SUBLOCATION_Z = 57
    USER_ID = 58
    CLIENT_ID = 59
    COMMAND_ID = 60
    CREATE_TIME = 61
    CREATE_CMD_TIME = 62
    COMPLETED_TIME = 63
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsINBOUND_DTL) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92},{94},{96},{98},{100},{102},{104},{106},{108},{110},{112},{114},{116},{118},{120},{122},{124},{126},{128}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}','{95}','{97}','{99}','{101}','{103}','{105}','{107}','{109}','{111}','{113}','{115}','{117}','{119}','{121}','{123}','{125}','{127}','{129}')",
      strSQL,
      TableName,
      IdxColumnName.KEY_NO.ToString, Info.KEY_NO,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_SERIAL_NO,
      IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.QTY_INBOUND.ToString, Info.QTY_INBOUND,
      IdxColumnName.ITEM_KEY_NO.ToString, Info.ITEM_KEY_NO,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.PACKAGE_ID,
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
      IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
      IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
      IdxColumnName.LENGTH.ToString, Info.LENGTH,
      IdxColumnName.WIDTH.ToString, Info.WIDTH,
      IdxColumnName.HEIGHT.ToString, Info.HEIGHT,
      IdxColumnName.WEIGHT.ToString, Info.WEIGHT,
      IdxColumnName.ITEM_VALUE.ToString, Info.ITEM_VALUE,
      IdxColumnName.CONTRACT_NO.ToString, Info.CONTRACT_NO,
      IdxColumnName.CONTRACT_SERIAL_NO.ToString, Info.CONTRACT_SERIAL_NO,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.STORAGE_TYPE),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.INBOUND_STATUS.ToString, CInt(Info.INBOUND_STATUS),
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
      IdxColumnName.QC_TIME.ToString, Info.QC_TIME,
      IdxColumnName.INBOUND_TIME.ToString, Info.INBOUND_TIME,
      IdxColumnName.RECEIPT_DATE.ToString, Info.RECEIPT_DATE,
      IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
      IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
      IdxColumnName.EFFECTIVE_DATE.ToString, Info.EFFECTIVE_DATE,
      IdxColumnName.LOCATION_NO.ToString, Info.LOCATION_NO,
      IdxColumnName.DEST_FACTORY_NO.ToString, Info.DEST_FACTORY_NO,
      IdxColumnName.DEST_AREA_NO.ToString, Info.DEST_AREA_NO,
      IdxColumnName.DEST_BLOCK_NO.ToString, Info.DEST_BLOCK_NO,
      IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_NO,
      IdxColumnName.ACTUAL_AREA_NO.ToString, Info.ACTUAL_AREA_NO,
      IdxColumnName.ACTUAL_LOCATION_NO.ToString, Info.ACTUAL_LOCATION_NO,
      IdxColumnName.ACTUAL_SUBLOCATION_X.ToString, Info.ACTUAL_SUBLOCATION_X,
      IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString, Info.ACTUAL_SUBLOCATION_Y,
      IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString, Info.ACTUAL_SUBLOCATION_Z,
      IdxColumnName.USER_ID.ToString, Info.USER_ID,
      IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
      IdxColumnName.COMMAND_ID.ToString, Info.COMMAND_ID,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.CREATE_CMD_TIME.ToString, Info.CREATE_CMD_TIME,
      IdxColumnName.COMPLETED_TIME.ToString, Info.COMPLETED_TIME
     )
      'Dim NewSQL As String = ""
      'If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
      Return strSQL
      'End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsINBOUND_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}',{94}='{95}',{96}='{97}',{98}='{99}',{100}='{101}',{102}='{103}',{104}='{105}',{106}='{107}',{108}='{109}',{110}='{111}',{112}='{113}',{114}='{115}',{116}='{117}',{118}='{119}',{120}='{121}',{122}='{123}',{124}='{125}',{126}='{127}',{128}='{129}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.KEY_NO.ToString, Info.KEY_NO,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_SERIAL_NO,
      IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.QTY_INBOUND.ToString, Info.QTY_INBOUND,
      IdxColumnName.ITEM_KEY_NO.ToString, Info.ITEM_KEY_NO,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.PACKAGE_ID,
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
      IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
      IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
      IdxColumnName.LENGTH.ToString, Info.LENGTH,
      IdxColumnName.WIDTH.ToString, Info.WIDTH,
      IdxColumnName.HEIGHT.ToString, Info.HEIGHT,
      IdxColumnName.WEIGHT.ToString, Info.WEIGHT,
      IdxColumnName.ITEM_VALUE.ToString, Info.ITEM_VALUE,
      IdxColumnName.CONTRACT_NO.ToString, Info.CONTRACT_NO,
      IdxColumnName.CONTRACT_SERIAL_NO.ToString, Info.CONTRACT_SERIAL_NO,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.STORAGE_TYPE),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.INBOUND_STATUS.ToString, CInt(Info.INBOUND_STATUS),
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
      IdxColumnName.QC_TIME.ToString, Info.QC_TIME,
      IdxColumnName.INBOUND_TIME.ToString, Info.INBOUND_TIME,
      IdxColumnName.RECEIPT_DATE.ToString, Info.RECEIPT_DATE,
      IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
      IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
      IdxColumnName.EFFECTIVE_DATE.ToString, Info.EFFECTIVE_DATE,
      IdxColumnName.LOCATION_NO.ToString, Info.LOCATION_NO,
      IdxColumnName.DEST_FACTORY_NO.ToString, Info.DEST_FACTORY_NO,
      IdxColumnName.DEST_AREA_NO.ToString, Info.DEST_AREA_NO,
      IdxColumnName.DEST_BLOCK_NO.ToString, Info.DEST_BLOCK_NO,
      IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_NO,
      IdxColumnName.ACTUAL_AREA_NO.ToString, Info.ACTUAL_AREA_NO,
      IdxColumnName.ACTUAL_LOCATION_NO.ToString, Info.ACTUAL_LOCATION_NO,
      IdxColumnName.ACTUAL_SUBLOCATION_X.ToString, Info.ACTUAL_SUBLOCATION_X,
      IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString, Info.ACTUAL_SUBLOCATION_Y,
      IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString, Info.ACTUAL_SUBLOCATION_Z,
      IdxColumnName.USER_ID.ToString, Info.USER_ID,
      IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
      IdxColumnName.COMMAND_ID.ToString, Info.COMMAND_ID,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.CREATE_CMD_TIME.ToString, Info.CREATE_CMD_TIME,
      IdxColumnName.COMPLETED_TIME.ToString, Info.COMPLETED_TIME
      )
      'Dim NewSQL As String = ""
      'If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
      Return strSQL
      'End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  ' Public Shared Function GetUpdateSQLForChangeValue(ByRef Info As clsINBOUND_DTL, ByRef dicChangeColumnValue As Dictionary(Of String, String)  ) As String
  '	Try
  ' Dim strSQL As String = ""
  'Dim strUpdateColumnValue As String = ""
  'If O_Get_UpdateColumnSQL(Of IdxColumnName)(dicChangeColumnValue, strUpdateColumnValue) = True Then
  'If strUpdateColumnValue <> "" Then
  ' strSQL = String.Format("Update {1} SET {2}  WHERE {3}='{4}'",
  ' strSQL,
  ' TableName,
  ' strUpdateColumnValue,
  ' IdxColumnName.KEY_NO.ToString, Info.KEY_NO,
  ' IdxColumnName.WO_ID.ToString, Info.WO_ID,
  ' IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_SERIAL_NO,
  ' IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
  ' IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
  ' IdxColumnName.QTY_INBOUND.ToString, Info.QTY_INBOUND,
  ' IdxColumnName.ITEM_KEY_NO.ToString, Info.ITEM_KEY_NO,
  ' IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
  ' IdxColumnName.PACKAGE_ID.ToString, Info.PACKAGE_ID,
  ' IdxColumnName.LOT_NO.ToString, Info.LOT_NO,
  ' IdxColumnName.ITEM_COMMON1.ToString, Info.ITEM_COMMON1,
  ' IdxColumnName.ITEM_COMMON2.ToString, Info.ITEM_COMMON2,
  ' IdxColumnName.ITEM_COMMON3.ToString, Info.ITEM_COMMON3,
  ' IdxColumnName.ITEM_COMMON4.ToString, Info.ITEM_COMMON4,
  ' IdxColumnName.ITEM_COMMON5.ToString, Info.ITEM_COMMON5,
  ' IdxColumnName.ITEM_COMMON6.ToString, Info.ITEM_COMMON6,
  ' IdxColumnName.ITEM_COMMON7.ToString, Info.ITEM_COMMON7,
  ' IdxColumnName.ITEM_COMMON8.ToString, Info.ITEM_COMMON8,
  ' IdxColumnName.ITEM_COMMON9.ToString, Info.ITEM_COMMON9,
  ' IdxColumnName.ITEM_COMMON10.ToString, Info.ITEM_COMMON10,
  ' IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
  ' IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
  ' IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
  ' IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
  ' IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
  ' IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
  ' IdxColumnName.CUSTOMER_NO.ToString, Info.CUSTOMER_NO,
  ' IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
  ' IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
  ' IdxColumnName.LENGTH.ToString, Info.LENGTH,
  ' IdxColumnName.WIDTH.ToString, Info.WIDTH,
  ' IdxColumnName.HEIGHT.ToString, Info.HEIGHT,
  ' IdxColumnName.WEIGHT.ToString, Info.WEIGHT,
  ' IdxColumnName.ITEM_VALUE.ToString, Info.ITEM_VALUE,
  ' IdxColumnName.CONTRACT_NO.ToString, Info.CONTRACT_NO,
  ' IdxColumnName.CONTRACT_SERIAL_NO.ToString, Info.CONTRACT_SERIAL_NO,
  ' IdxColumnName.PO_ID.ToString, Info.PO_ID,
  ' IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
  ' IdxColumnName.STORAGE_TYPE.ToString, CINT(Info.STORAGE_TYPE),
  ' IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
  ' IdxColumnName.INBOUND_STATUS.ToString, CINT(Info.INBOUND_STATUS),
  ' IdxColumnName.QC_STATUS.ToString, CINT(Info.QC_STATUS),
  ' IdxColumnName.QC_TIME.ToString, Info.QC_TIME,
  ' IdxColumnName.INBOUND_TIME.ToString, Info.INBOUND_TIME,
  ' IdxColumnName.RECEIPT_DATE.ToString, Info.RECEIPT_DATE,
  ' IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
  ' IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
  ' IdxColumnName.EFFECTIVE_DATE.ToString, Info.EFFECTIVE_DATE,
  ' IdxColumnName.LOCATION_NO.ToString, Info.LOCATION_NO,
  ' IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
  ' IdxColumnName.DEST_AREA_NO.ToString, Info.DEST_AREA_NO,
  ' IdxColumnName.DEST_BLOCK_NO.ToString, Info.DEST_BLOCK_NO,
  ' IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_NO,
  ' IdxColumnName.ACTUAL_AREA_NO.ToString, Info.ACTUAL_AREA_NO,
  ' IdxColumnName.ACTUAL_LOCATION_NO.ToString, Info.ACTUAL_LOCATION_NO,
  ' IdxColumnName.ACTUAL_SUBLOCATION_X.ToString, Info.ACTUAL_SUBLOCATION_X,
  ' IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString, Info.ACTUAL_SUBLOCATION_Y,
  ' IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString, Info.ACTUAL_SUBLOCATION_Z,
  ' IdxColumnName.USER_ID.ToString, Info.USER_ID,
  ' IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
  ' IdxColumnName.COMMAND_ID.ToString, Info.COMMAND_ID,
  ' IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
  ' IdxColumnName.CREATE_CMD_TIME.ToString, Info.CREATE_CMD_TIME,
  ' IdxColumnName.COMPLETED_TIME.ToString, Info.COMPLETED_TIME
  ' )
  ' Dim NewSQL As String = ""
  ' If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
  ' Return NewSQL
  ' End If
  ' End If
  ' End If
  '  Return ""
  ' Catch ex As Exception
  ' SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  ' Return ""
  ' End Try
  ' End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsINBOUND_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.KEY_NO.ToString, Info.KEY_NO,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_SERIAL_NO,
      IdxColumnName.CARRIER_ID.ToString, Info.CARRIER_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.QTY_INBOUND.ToString, Info.QTY_INBOUND,
      IdxColumnName.ITEM_KEY_NO.ToString, Info.ITEM_KEY_NO,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.PACKAGE_ID,
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
      IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
      IdxColumnName.SUB_OWNER_NO.ToString, Info.SUB_OWNER_NO,
      IdxColumnName.LENGTH.ToString, Info.LENGTH,
      IdxColumnName.WIDTH.ToString, Info.WIDTH,
      IdxColumnName.HEIGHT.ToString, Info.HEIGHT,
      IdxColumnName.WEIGHT.ToString, Info.WEIGHT,
      IdxColumnName.ITEM_VALUE.ToString, Info.ITEM_VALUE,
      IdxColumnName.CONTRACT_NO.ToString, Info.CONTRACT_NO,
      IdxColumnName.CONTRACT_SERIAL_NO.ToString, Info.CONTRACT_SERIAL_NO,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.STORAGE_TYPE),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.INBOUND_STATUS.ToString, CInt(Info.INBOUND_STATUS),
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
      IdxColumnName.QC_TIME.ToString, Info.QC_TIME,
      IdxColumnName.INBOUND_TIME.ToString, Info.INBOUND_TIME,
      IdxColumnName.RECEIPT_DATE.ToString, Info.RECEIPT_DATE,
      IdxColumnName.MANUFACETURE_DATE.ToString, Info.MANUFACETURE_DATE,
      IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
      IdxColumnName.EFFECTIVE_DATE.ToString, Info.EFFECTIVE_DATE,
      IdxColumnName.LOCATION_NO.ToString, Info.LOCATION_NO,
      IdxColumnName.DEST_FACTORY_NO.ToString, Info.DEST_FACTORY_NO,
      IdxColumnName.DEST_AREA_NO.ToString, Info.DEST_AREA_NO,
      IdxColumnName.DEST_BLOCK_NO.ToString, Info.DEST_BLOCK_NO,
      IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_NO,
      IdxColumnName.ACTUAL_AREA_NO.ToString, Info.ACTUAL_AREA_NO,
      IdxColumnName.ACTUAL_LOCATION_NO.ToString, Info.ACTUAL_LOCATION_NO,
      IdxColumnName.ACTUAL_SUBLOCATION_X.ToString, Info.ACTUAL_SUBLOCATION_X,
      IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString, Info.ACTUAL_SUBLOCATION_Y,
      IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString, Info.ACTUAL_SUBLOCATION_Z,
      IdxColumnName.USER_ID.ToString, Info.USER_ID,
      IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
      IdxColumnName.COMMAND_ID.ToString, Info.COMMAND_ID,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.CREATE_CMD_TIME.ToString, Info.CREATE_CMD_TIME,
      IdxColumnName.COMPLETED_TIME.ToString, Info.COMPLETED_TIME
      )
      'Dim NewSQL As String = ""
      'If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
      Return strSQL
      'End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Private Shared Function SetInfoFromDB(ByRef Info As clsINBOUND_DTL, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim KEY_NO = "" & RowData.Item(IdxColumnName.KEY_NO.ToString)
        Dim WO_ID = "" & RowData.Item(IdxColumnName.WO_ID.ToString)
        Dim WO_SERIAL_NO = "" & RowData.Item(IdxColumnName.WO_SERIAL_NO.ToString)
        Dim CARRIER_ID = "" & RowData.Item(IdxColumnName.CARRIER_ID.ToString)
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim QTY_INBOUND = If(IsNumeric(RowData.Item(IdxColumnName.QTY_INBOUND.ToString)), RowData.Item(IdxColumnName.QTY_INBOUND.ToString), 0 & RowData.Item(IdxColumnName.QTY_INBOUND.ToString))
        Dim ITEM_KEY_NO = "" & RowData.Item(IdxColumnName.ITEM_KEY_NO.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim PACKAGE_ID = "" & RowData.Item(IdxColumnName.PACKAGE_ID.ToString)
        Dim LOT_NO = "" & RowData.Item(IdxColumnName.LOT_NO.ToString)
        Dim ITEM_COMMON1 = "" & RowData.Item(IdxColumnName.ITEM_COMMON1.ToString)
        Dim ITEM_COMMON2 = "" & RowData.Item(IdxColumnName.ITEM_COMMON2.ToString)
        Dim ITEM_COMMON3 = "" & RowData.Item(IdxColumnName.ITEM_COMMON3.ToString)
        Dim ITEM_COMMON4 = "" & RowData.Item(IdxColumnName.ITEM_COMMON4.ToString)
        Dim ITEM_COMMON5 = "" & RowData.Item(IdxColumnName.ITEM_COMMON5.ToString)
        Dim ITEM_COMMON6 = "" & RowData.Item(IdxColumnName.ITEM_COMMON6.ToString)
        Dim ITEM_COMMON7 = "" & RowData.Item(IdxColumnName.ITEM_COMMON7.ToString)
        Dim ITEM_COMMON8 = "" & RowData.Item(IdxColumnName.ITEM_COMMON8.ToString)
        Dim ITEM_COMMON9 = "" & RowData.Item(IdxColumnName.ITEM_COMMON9.ToString)
        Dim ITEM_COMMON10 = "" & RowData.Item(IdxColumnName.ITEM_COMMON10.ToString)
        Dim SORT_ITEM_COMMON1 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON1.ToString)
        Dim SORT_ITEM_COMMON2 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON2.ToString)
        Dim SORT_ITEM_COMMON3 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON3.ToString)
        Dim SORT_ITEM_COMMON4 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON4.ToString)
        Dim SORT_ITEM_COMMON5 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON5.ToString)
        Dim SUPPLIER_NO = "" & RowData.Item(IdxColumnName.SUPPLIER_NO.ToString)
        Dim CUSTOMER_NO = "" & RowData.Item(IdxColumnName.CUSTOMER_NO.ToString)
        Dim OWNER_NO = "" & RowData.Item(IdxColumnName.OWNER_NO.ToString)
        Dim SUB_OWNER_NO = "" & RowData.Item(IdxColumnName.SUB_OWNER_NO.ToString)
        Dim LENGTH = If(IsNumeric(RowData.Item(IdxColumnName.LENGTH.ToString)), RowData.Item(IdxColumnName.LENGTH.ToString), 0 & RowData.Item(IdxColumnName.LENGTH.ToString))
        Dim WIDTH = If(IsNumeric(RowData.Item(IdxColumnName.WIDTH.ToString)), RowData.Item(IdxColumnName.WIDTH.ToString), 0 & RowData.Item(IdxColumnName.WIDTH.ToString))
        Dim HEIGHT = If(IsNumeric(RowData.Item(IdxColumnName.HEIGHT.ToString)), RowData.Item(IdxColumnName.HEIGHT.ToString), 0 & RowData.Item(IdxColumnName.HEIGHT.ToString))
        Dim WEIGHT = If(IsNumeric(RowData.Item(IdxColumnName.WEIGHT.ToString)), RowData.Item(IdxColumnName.WEIGHT.ToString), 0 & RowData.Item(IdxColumnName.WEIGHT.ToString))
        Dim ITEM_VALUE = If(IsNumeric(RowData.Item(IdxColumnName.ITEM_VALUE.ToString)), RowData.Item(IdxColumnName.ITEM_VALUE.ToString), 0 & RowData.Item(IdxColumnName.ITEM_VALUE.ToString))
        Dim CONTRACT_NO = "" & RowData.Item(IdxColumnName.CONTRACT_NO.ToString)
        Dim CONTRACT_SERIAL_NO = "" & RowData.Item(IdxColumnName.CONTRACT_SERIAL_NO.ToString)
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_SERIAL_NO = "" & RowData.Item(IdxColumnName.PO_SERIAL_NO.ToString)
        Dim STORAGE_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.STORAGE_TYPE.ToString)), RowData.Item(IdxColumnName.STORAGE_TYPE.ToString), 0 & RowData.Item(IdxColumnName.STORAGE_TYPE.ToString))
        Dim BND = IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.BND.ToString))
        Dim INBOUND_STATUS = If(IsNumeric(RowData.Item(IdxColumnName.INBOUND_STATUS.ToString)), RowData.Item(IdxColumnName.INBOUND_STATUS.ToString), 0 & RowData.Item(IdxColumnName.INBOUND_STATUS.ToString))
        Dim QC_STATUS = If(IsNumeric(RowData.Item(IdxColumnName.QC_STATUS.ToString)), RowData.Item(IdxColumnName.QC_STATUS.ToString), 0 & RowData.Item(IdxColumnName.QC_STATUS.ToString))
        Dim QC_TIME = "" & RowData.Item(IdxColumnName.QC_TIME.ToString)
        Dim INBOUND_TIME = "" & RowData.Item(IdxColumnName.INBOUND_TIME.ToString)
        Dim RECEIPT_DATE = "" & RowData.Item(IdxColumnName.RECEIPT_DATE.ToString)
        Dim MANUFACETURE_DATE = "" & RowData.Item(IdxColumnName.MANUFACETURE_DATE.ToString)
        Dim EXPIRED_DATE = "" & RowData.Item(IdxColumnName.EXPIRED_DATE.ToString)
        Dim EFFECTIVE_DATE = "" & RowData.Item(IdxColumnName.EFFECTIVE_DATE.ToString)
        Dim LOCATION_NO = "" & RowData.Item(IdxColumnName.LOCATION_NO.ToString)
        Dim DEST_FACTORY_NO = "" & RowData.Item(IdxColumnName.DEST_FACTORY_NO.ToString)
        Dim DEST_AREA_NO = "" & RowData.Item(IdxColumnName.DEST_AREA_NO.ToString)
        Dim DEST_BLOCK_NO = "" & RowData.Item(IdxColumnName.DEST_BLOCK_NO.ToString)
        Dim DEST_LOCATION_NO = "" & RowData.Item(IdxColumnName.DEST_LOCATION_NO.ToString)
        Dim ACTUAL_AREA_NO = "" & RowData.Item(IdxColumnName.ACTUAL_AREA_NO.ToString)
        Dim ACTUAL_LOCATION_NO = "" & RowData.Item(IdxColumnName.ACTUAL_LOCATION_NO.ToString)
        Dim ACTUAL_SUBLOCATION_X = "" & RowData.Item(IdxColumnName.ACTUAL_SUBLOCATION_X.ToString)
        Dim ACTUAL_SUBLOCATION_Y = "" & RowData.Item(IdxColumnName.ACTUAL_SUBLOCATION_Y.ToString)
        Dim ACTUAL_SUBLOCATION_Z = "" & RowData.Item(IdxColumnName.ACTUAL_SUBLOCATION_Z.ToString)
        Dim USER_ID = "" & RowData.Item(IdxColumnName.USER_ID.ToString)
        Dim CLIENT_ID = "" & RowData.Item(IdxColumnName.CLIENT_ID.ToString)
        Dim COMMAND_ID = "" & RowData.Item(IdxColumnName.COMMAND_ID.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim CREATE_CMD_TIME = "" & RowData.Item(IdxColumnName.CREATE_CMD_TIME.ToString)
        Dim COMPLETED_TIME = "" & RowData.Item(IdxColumnName.COMPLETED_TIME.ToString)
        Info = New clsINBOUND_DTL(KEY_NO, WO_ID, WO_SERIAL_NO, CARRIER_ID, SKU_NO, QTY_INBOUND, ITEM_KEY_NO, COMMENTS, PACKAGE_ID, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, SUPPLIER_NO, CUSTOMER_NO, OWNER_NO, SUB_OWNER_NO, LENGTH, WIDTH, HEIGHT, WEIGHT, ITEM_VALUE, CONTRACT_NO, CONTRACT_SERIAL_NO, PO_ID, PO_SERIAL_NO, STORAGE_TYPE, BND, INBOUND_STATUS, QC_STATUS, QC_TIME, INBOUND_TIME, RECEIPT_DATE, MANUFACETURE_DATE, EXPIRED_DATE, EFFECTIVE_DATE, LOCATION_NO, DEST_FACTORY_NO, DEST_AREA_NO, DEST_BLOCK_NO, DEST_LOCATION_NO, ACTUAL_AREA_NO, ACTUAL_LOCATION_NO, ACTUAL_SUBLOCATION_X, ACTUAL_SUBLOCATION_Y, ACTUAL_SUBLOCATION_Z, USER_ID, CLIENT_ID, COMMAND_ID, CREATE_TIME, CREATE_CMD_TIME, COMPLETED_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetDataDicByALL() As Dictionary(Of String, clsINBOUND_DTL)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsINBOUND_DTL)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} ",
        strSQL,
        TableName
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsINBOUND_DTL = Nothing
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
  Public Shared Function GetDataDicByPO_ID(ByVal PO_ID As String) As Dictionary(Of String, clsINBOUND_DTL)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsINBOUND_DTL)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE {2}='{3}' ",
        strSQL,
        TableName,
        IdxColumnName.PO_ID.ToString,
        PO_ID
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsINBOUND_DTL = Nothing
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
