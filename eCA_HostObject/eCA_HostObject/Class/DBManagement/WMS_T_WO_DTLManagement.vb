Imports System.Collections.Concurrent
Partial Class WMS_T_WO_DTLManagement
  Public Shared TableName As String = "WMS_T_WO_DTL"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    WO_ID
    WO_SERIAL_NO
    WORKING_TYPE
    WORKING_SERIAL_NO
    WORKING_SERIAL_SEQ
    QC_METHOD
    SKU_NO
    SKU_CATALOG
    LOT_NO
    QTY
    QTY_TRANSFERRED
    QTY_PROCESS
    QTY_REPLENISHMENT
    QTY_ABORT
    CARRIER_ID
    PACKAGE_ID
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
    COMMENTS
    SL_NO
    EXPIRED_DATE
    STORAGE_TYPE
    BND
    QC_STATUS
    FROM_OWNER_NO
    FROM_SUB_OWNER_NO
    TO_OWNER_NO
    TO_SUB_OWNER_NO
    FACTORY_NO
    DEST_AREA_NO
    DEST_LOCATION_NO
    SOURCE_AREA_NO
    SOURCE_LOCATION_NO
    START_TIME
    START_TRANSFER_TIME
    FINISH_TRANSFER_TIME
    FINISH_TIME
    WO_DTL_DC_STATUS
    DTL_COMMON1
    DTL_COMMON2
    DTL_COMMON3
    DTL_COMMON4
    DTL_COMMON5
  End Enum

  '- GetSQL
  '-請將 clsWO_DTL 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsWO_DTL) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92},{94},{96},{98},{100},{102},{104},{106},{108},{110},{112}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}','{95}','{97}','{99}','{101}','{103}','{105}','{107}','{109}','{111}','{113}')",
      strSQL,
      TableName,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_Serial_No,
      IdxColumnName.WORKING_TYPE.ToString, CInt(Info.WORKING_TYPE),
      IdxColumnName.WORKING_SERIAL_NO.ToString, Info.WORKING_SERIAL_NO,
      IdxColumnName.WORKING_SERIAL_SEQ.ToString, Info.WORKING_SERIAL_SEQ,
      IdxColumnName.QC_METHOD.ToString, CInt(Info.QC_Method),
      IdxColumnName.SKU_NO.ToString, Info.SKU_No,
      IdxColumnName.SKU_CATALOG.ToString, CInt(Info.SKU_Catalog),
      IdxColumnName.LOT_NO.ToString, Info.Lot_No,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_TRANSFERRED.ToString, Info.QTY_Transferred,
      IdxColumnName.QTY_PROCESS.ToString, Info.QTY_Process,
      IdxColumnName.QTY_REPLENISHMENT.ToString, Info.QTY_Replenishment,
      IdxColumnName.QTY_ABORT.ToString, Info.QTY_Abort,
      IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
      IdxColumnName.PACKAGE_ID.ToString, Info.Package_ID,
      IdxColumnName.ITEM_COMMON1.ToString, Info.Item_Common1,
      IdxColumnName.ITEM_COMMON2.ToString, Info.Item_Common2,
      IdxColumnName.ITEM_COMMON3.ToString, Info.Item_Common3,
      IdxColumnName.ITEM_COMMON4.ToString, Info.Item_Common4,
      IdxColumnName.ITEM_COMMON5.ToString, Info.Item_Common5,
      IdxColumnName.ITEM_COMMON6.ToString, Info.Item_Common6,
      IdxColumnName.ITEM_COMMON7.ToString, Info.Item_Common7,
      IdxColumnName.ITEM_COMMON8.ToString, Info.Item_Common8,
      IdxColumnName.ITEM_COMMON9.ToString, Info.Item_Common9,
      IdxColumnName.ITEM_COMMON10.ToString, Info.Item_Common10,
      IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
      IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
      IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
      IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
      IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
      IdxColumnName.COMMENTS.ToString, Info.Comments,
      IdxColumnName.SL_NO.ToString, Info.SL_NO,
      IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.Storage_Type),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.QC_STATUS.ToString, Info.QC_Status,
      IdxColumnName.FROM_OWNER_NO.ToString, Info.FROM_OWNER_No,
      IdxColumnName.FROM_SUB_OWNER_NO.ToString, Info.FROM_SUB_OWNER_No,
      IdxColumnName.TO_OWNER_NO.ToString, Info.TO_OWNER_No,
      IdxColumnName.TO_SUB_OWNER_NO.ToString, Info.TO_SUB_OWNER_No,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_No,
      IdxColumnName.DEST_AREA_NO.ToString, Info.DEST_AREA_No,
      IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_No,
      IdxColumnName.SOURCE_AREA_NO.ToString, Info.SOURCE_AREA_No,
      IdxColumnName.SOURCE_LOCATION_NO.ToString, Info.SOURCE_LOCATION_NO,
      IdxColumnName.START_TIME.ToString, Info.Start_Time,
      IdxColumnName.START_TRANSFER_TIME.ToString, Info.Start_Transfer_Time,
      IdxColumnName.FINISH_TRANSFER_TIME.ToString, Info.Finish_Transfer_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.Finish_Time,
      IdxColumnName.WO_DTL_DC_STATUS.ToString, CInt(Info.WO_DTL_DC_Status),
      IdxColumnName.DTL_COMMON1.ToString, Info.DTL_Common1,
      IdxColumnName.DTL_COMMON2.ToString, Info.DTL_Common2,
      IdxColumnName.DTL_COMMON3.ToString, Info.DTL_Common3,
      IdxColumnName.DTL_COMMON4.ToString, Info.DTL_Common4,
      IdxColumnName.DTL_COMMON5.ToString, Info.DTL_Common5
     )
      Dim NewSQL As String = ""
      If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsWO_DTL) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' ",
      strSQL,
      TableName,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_Serial_No
      )
      Dim NewSQL As String = ""
      If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsWO_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}',{94}='{95}',{96}='{97}',{98}='{99}',{100}='{101}',{102}='{103}',{104}='{105}',{106}='{107}',{108}='{109}',{110}='{111}',{112}='{113}' WHERE {2}='{3}' And {4}='{5}'",
      strSQL,
      TableName,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_Serial_No,
      IdxColumnName.WORKING_TYPE.ToString, CInt(Info.WORKING_TYPE),
      IdxColumnName.WORKING_SERIAL_NO.ToString, Info.WORKING_SERIAL_NO,
      IdxColumnName.WORKING_SERIAL_SEQ.ToString, Info.WORKING_SERIAL_SEQ,
      IdxColumnName.QC_METHOD.ToString, CInt(Info.QC_Method),
      IdxColumnName.SKU_NO.ToString, Info.SKU_No,
      IdxColumnName.SKU_CATALOG.ToString, CInt(Info.SKU_Catalog),
      IdxColumnName.LOT_NO.ToString, Info.Lot_No,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_TRANSFERRED.ToString, Info.QTY_Transferred,
      IdxColumnName.QTY_PROCESS.ToString, Info.QTY_Process,
      IdxColumnName.QTY_REPLENISHMENT.ToString, Info.QTY_Replenishment,
      IdxColumnName.QTY_ABORT.ToString, Info.QTY_Abort,
      IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
      IdxColumnName.PACKAGE_ID.ToString, Info.Package_ID,
      IdxColumnName.ITEM_COMMON1.ToString, Info.Item_Common1,
      IdxColumnName.ITEM_COMMON2.ToString, Info.Item_Common2,
      IdxColumnName.ITEM_COMMON3.ToString, Info.Item_Common3,
      IdxColumnName.ITEM_COMMON4.ToString, Info.Item_Common4,
      IdxColumnName.ITEM_COMMON5.ToString, Info.Item_Common5,
      IdxColumnName.ITEM_COMMON6.ToString, Info.Item_Common6,
      IdxColumnName.ITEM_COMMON7.ToString, Info.Item_Common7,
      IdxColumnName.ITEM_COMMON8.ToString, Info.Item_Common8,
      IdxColumnName.ITEM_COMMON9.ToString, Info.Item_Common9,
      IdxColumnName.ITEM_COMMON10.ToString, Info.Item_Common10,
      IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
      IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
      IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
      IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
      IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
      IdxColumnName.COMMENTS.ToString, Info.Comments,
      IdxColumnName.SL_NO.ToString, Info.SL_NO,
      IdxColumnName.EXPIRED_DATE.ToString, Info.EXPIRED_DATE,
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.Storage_Type),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.QC_STATUS.ToString, Info.QC_Status,
      IdxColumnName.FROM_OWNER_NO.ToString, Info.FROM_OWNER_No,
      IdxColumnName.FROM_SUB_OWNER_NO.ToString, Info.FROM_SUB_OWNER_No,
      IdxColumnName.TO_OWNER_NO.ToString, Info.TO_OWNER_No,
      IdxColumnName.TO_SUB_OWNER_NO.ToString, Info.TO_SUB_OWNER_No,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_No,
      IdxColumnName.DEST_AREA_NO.ToString, Info.DEST_AREA_No,
      IdxColumnName.DEST_LOCATION_NO.ToString, Info.DEST_LOCATION_No,
      IdxColumnName.SOURCE_AREA_NO.ToString, Info.SOURCE_AREA_No,
      IdxColumnName.SOURCE_LOCATION_NO.ToString, Info.SOURCE_LOCATION_NO,
      IdxColumnName.START_TIME.ToString, Info.Start_Time,
      IdxColumnName.START_TRANSFER_TIME.ToString, Info.Start_Transfer_Time,
      IdxColumnName.FINISH_TRANSFER_TIME.ToString, Info.Finish_Transfer_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.Finish_Time,
      IdxColumnName.WO_DTL_DC_STATUS.ToString, CInt(Info.WO_DTL_DC_Status),
      IdxColumnName.DTL_COMMON1.ToString, Info.DTL_Common1,
      IdxColumnName.DTL_COMMON2.ToString, Info.DTL_Common2,
      IdxColumnName.DTL_COMMON3.ToString, Info.DTL_Common3,
      IdxColumnName.DTL_COMMON4.ToString, Info.DTL_Common4,
      IdxColumnName.DTL_COMMON5.ToString, Info.DTL_Common5
      )
      Dim NewSQL As String = ""
      If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- GET
  Public Shared Function GetWMS_T_WO_DTLDataListByALL() As List(Of clsWO_DTL)
    Try
      Dim _lstReturn As New List(Of clsWO_DTL)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
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
              Dim Info As clsWO_DTL = Nothing
              SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
              _lstReturn.Add(Info)
            Next
          End If
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  Public Shared Function GetWMS_T_WO_DTLDataListByWO_ID(ByVal WO_ID As String) As List(Of clsWO_DTL)
    Try
      Dim _lstReturn As New List(Of clsWO_DTL)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strSQL As String = String.Empty
          Dim rs As DataSet = Nothing
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select * from {0} WHERE {1} = '{2}'", TableName, IdxColumnName.WO_ID.ToString, WO_ID)
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          'Dim OLEDBAdapter As New OleDbDataAdapter
          'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsWO_DTL = Nothing
              SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
              _lstReturn.Add(Info)
            Next
          End If
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsWO_DTL, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim WO_ID = "" & RowData.Item(IdxColumnName.WO_ID.ToString)
        Dim WO_SERIAL_NO = "" & RowData.Item(IdxColumnName.WO_SERIAL_NO.ToString)
        Dim WORKING_TYPE = IIf(IsNumeric(RowData.Item(IdxColumnName.WORKING_TYPE.ToString)), RowData.Item(IdxColumnName.WORKING_TYPE.ToString), 0)
        Dim WORKING_SERIAL_NO = "" & RowData.Item(IdxColumnName.WORKING_SERIAL_NO.ToString)
        Dim WORKING_SERIAL_SEQ = "" & RowData.Item(IdxColumnName.WORKING_SERIAL_SEQ.ToString)
        Dim QC_METHOD = IIf(IsNumeric(RowData.Item(IdxColumnName.QC_METHOD.ToString)), RowData.Item(IdxColumnName.QC_METHOD.ToString), 0)
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim SKU_CATALOG = IIf(IsNumeric(RowData.Item(IdxColumnName.SKU_CATALOG.ToString)), RowData.Item(IdxColumnName.SKU_CATALOG.ToString), 0)
        Dim LOT_NO = "" & RowData.Item(IdxColumnName.LOT_NO.ToString)
        Dim QTY = IIf(IsNumeric(RowData.Item(IdxColumnName.QTY.ToString)), RowData.Item(IdxColumnName.QTY.ToString), 0)
        Dim QTY_TRANSFERRED = IIf(IsNumeric(RowData.Item(IdxColumnName.QTY_TRANSFERRED.ToString)), RowData.Item(IdxColumnName.QTY_TRANSFERRED.ToString), 0)
        Dim QTY_PROCESS = IIf(IsNumeric(RowData.Item(IdxColumnName.QTY_PROCESS.ToString)), RowData.Item(IdxColumnName.QTY_PROCESS.ToString), 0)
        Dim QTY_REPLENISHMENT = IIf(IsNumeric(RowData.Item(IdxColumnName.QTY_REPLENISHMENT.ToString)), RowData.Item(IdxColumnName.QTY_REPLENISHMENT.ToString), 0)
        Dim QTY_ABORT = IIf(IsNumeric(RowData.Item(IdxColumnName.QTY_ABORT.ToString)), RowData.Item(IdxColumnName.QTY_ABORT.ToString), 0)
        Dim CARRIER_ID = "" & RowData.Item(IdxColumnName.CARRIER_ID.ToString)
        Dim PACKAGE_ID = "" & RowData.Item(IdxColumnName.PACKAGE_ID.ToString)
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
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim SL_NO = "" & RowData.Item(IdxColumnName.SL_NO.ToString)
        Dim EXPIRED_DATE = "" & RowData.Item(IdxColumnName.EXPIRED_DATE.ToString)
        Dim STORAGE_TYPE = IIf(IsNumeric(RowData.Item(IdxColumnName.STORAGE_TYPE.ToString)), RowData.Item(IdxColumnName.STORAGE_TYPE.ToString), 0)
        Dim BND = IntegerConvertToBoolean("" & RowData.Item(IdxColumnName.BND.ToString))
        Dim QC_STATUS = "" & RowData.Item(IdxColumnName.QC_STATUS.ToString)
        Dim FROM_OWNER_NO = "" & RowData.Item(IdxColumnName.FROM_OWNER_NO.ToString)
        Dim FROM_SUB_OWNER_NO = "" & RowData.Item(IdxColumnName.FROM_SUB_OWNER_NO.ToString)
        Dim TO_OWNER_NO = "" & RowData.Item(IdxColumnName.TO_OWNER_NO.ToString)
        Dim TO_SUB_OWNER_NO = "" & RowData.Item(IdxColumnName.TO_SUB_OWNER_NO.ToString)
        Dim FACTORY_NO = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim DEST_AREA_NO = "" & RowData.Item(IdxColumnName.DEST_AREA_NO.ToString)
        Dim DEST_LOCATION_NO = "" & RowData.Item(IdxColumnName.DEST_LOCATION_NO.ToString)
        Dim SOURCE_AREA_NO = "" & RowData.Item(IdxColumnName.SOURCE_AREA_NO.ToString)
        Dim SOURCE_LOCATION_NO = "" & RowData.Item(IdxColumnName.SOURCE_LOCATION_NO.ToString)
        Dim START_TIME = "" & RowData.Item(IdxColumnName.START_TIME.ToString)
        Dim START_TRANSFER_TIME = "" & RowData.Item(IdxColumnName.START_TRANSFER_TIME.ToString)
        Dim FINISH_TRANSFER_TIME = "" & RowData.Item(IdxColumnName.FINISH_TRANSFER_TIME.ToString)
        Dim FINISH_TIME = "" & RowData.Item(IdxColumnName.FINISH_TIME.ToString)
        Dim WO_DTL_DC_STATUS = IIf(IsNumeric(RowData.Item(IdxColumnName.WO_DTL_DC_STATUS.ToString)), RowData.Item(IdxColumnName.WO_DTL_DC_STATUS.ToString), 0)
        Dim DTL_COMMON1 = "" & RowData.Item(IdxColumnName.DTL_COMMON1.ToString)
        Dim DTL_COMMON2 = "" & RowData.Item(IdxColumnName.DTL_COMMON2.ToString)
        Dim DTL_COMMON3 = "" & RowData.Item(IdxColumnName.DTL_COMMON3.ToString)
        Dim DTL_COMMON4 = "" & RowData.Item(IdxColumnName.DTL_COMMON4.ToString)
        Dim DTL_COMMON5 = "" & RowData.Item(IdxColumnName.DTL_COMMON5.ToString)
        Info = New clsWO_DTL(WO_ID, WO_SERIAL_NO,
                             WORKING_TYPE, WORKING_SERIAL_NO, WORKING_SERIAL_SEQ, QC_METHOD,
                             SKU_NO, SKU_CATALOG, LOT_NO,
                             QTY, QTY_TRANSFERRED, QTY_PROCESS, QTY_REPLENISHMENT, QTY_ABORT,
                             CARRIER_ID, PACKAGE_ID,
                             ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5,
                             ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10,
                             SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5,
                             COMMENTS, EXPIRED_DATE, SL_NO, STORAGE_TYPE, BND, QC_STATUS,
                             FROM_OWNER_NO, FROM_SUB_OWNER_NO, TO_OWNER_NO, TO_SUB_OWNER_NO,
                             FACTORY_NO, DEST_AREA_NO, DEST_LOCATION_NO, SOURCE_AREA_NO, SOURCE_LOCATION_NO,
                             START_TIME, START_TRANSFER_TIME, FINISH_TRANSFER_TIME, FINISH_TIME,
                             WO_DTL_DC_STATUS, DTL_COMMON1, DTL_COMMON2, DTL_COMMON3, DTL_COMMON4, DTL_COMMON5)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
