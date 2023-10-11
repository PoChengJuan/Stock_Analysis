﻿Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Public Class WMS_T_PO_DTLManagement
  Public Shared TableName As String = "WMS_T_PO_DTL"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsPO_DTL)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
  Public Shared LogTool As eCALogTool._ILogTool = Nothing

  Enum IdxColumnName As Integer
    PO_ID
    PO_LINE_NO
    PO_SERIAL_NO
    WORKING_TYPE
    WORKING_SERIAL_NO
    WORKING_SERIAL_SEQ
    SKU_NO
    LOT_NO
    QTY
    QTY_PROCESS
    QTY_FINISH
    COMMENTS
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
    STORAGE_TYPE
    BND
    QC_STATUS
    FROM_OWNER_ID
    FROM_SUB_OWNER_ID
    TO_OWNER_ID
    TO_SUB_OWNER_ID
    FACTORY_ID
    DEST_AREA_ID
    DEST_LOCATION_ID
    H_POD_STEP_NO
    H_POD_MOVE_TYPE
    H_POD_FINISH_TIME
    H_POD_BILLING_DATE
    H_POD_CREATE_TIME
    H_POD1
    H_POD2
    H_POD3
    H_POD4
    H_POD5
    H_POD6
    H_POD7
    H_POD8
    H_POD9
    H_POD10
    H_POD11
    H_POD12
    H_POD13
    H_POD14
    H_POD15
    H_POD16
    H_POD17
    H_POD18
    H_POD19
    H_POD20
    H_POD21
    H_POD22
    H_POD23
    H_POD24
    H_POD25
    PODTL_STATUS
    CLOSE_ABLE
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPO_DTL) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92},{94},{96},{98},{100},{102},{104},{106},{108},{110},{112},{114},{116},{118},{120},{122},{124},{126},{128},{130},{132},{134},{136},{138},{140}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}','{95}','{97}','{99}','{101}','{103}','{105}','{107}','{109}','{111}','{113}','{115}','{117}','{119}','{121}','{123}','{125}','{127}','{129}','{131}','{133}','{135}','{137}','{139}','{141}')",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.WORKING_TYPE.ToString, Info.WORKING_TYPE,
      IdxColumnName.WORKING_SERIAL_NO.ToString, Info.WORKING_SERIAL_NO,
      IdxColumnName.WORKING_SERIAL_SEQ.ToString, Info.WORKING_SERIAL_SEQ,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.LOT_NO.ToString, Info.LOT_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.PACKAGE_ID,
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
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.STORAGE_TYPE),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
      IdxColumnName.FROM_OWNER_ID.ToString, Info.FROM_OWNER_ID,
      IdxColumnName.FROM_SUB_OWNER_ID.ToString, Info.FROM_SUB_OWNER_ID,
      IdxColumnName.TO_OWNER_ID.ToString, Info.TO_OWNER_ID,
      IdxColumnName.TO_SUB_OWNER_ID.ToString, Info.TO_SUB_OWNER_ID,
      IdxColumnName.FACTORY_ID.ToString, Info.FACTORY_ID,
      IdxColumnName.DEST_AREA_ID.ToString, Info.DEST_AREA_ID,
      IdxColumnName.DEST_LOCATION_ID.ToString, Info.DEST_LOCATION_ID,
      IdxColumnName.H_POD_STEP_NO.ToString, CInt(Info.H_POD_STEP_NO),
      IdxColumnName.H_POD_MOVE_TYPE.ToString, Info.H_POD_MOVE_TYPE,
      IdxColumnName.H_POD_FINISH_TIME.ToString, Info.H_POD_FINISH_TIME,
      IdxColumnName.H_POD_BILLING_DATE.ToString, Info.H_POD_BILLING_DATE,
      IdxColumnName.H_POD_CREATE_TIME.ToString, Info.H_POD_CREATE_TIME,
      IdxColumnName.H_POD1.ToString, Info.H_POD1,
      IdxColumnName.H_POD2.ToString, Info.H_POD2,
      IdxColumnName.H_POD3.ToString, Info.H_POD3,
      IdxColumnName.H_POD4.ToString, Info.H_POD4,
      IdxColumnName.H_POD5.ToString, Info.H_POD5,
      IdxColumnName.H_POD6.ToString, Info.H_POD6,
      IdxColumnName.H_POD7.ToString, Info.H_POD7,
      IdxColumnName.H_POD8.ToString, Info.H_POD8,
      IdxColumnName.H_POD9.ToString, Info.H_POD9,
      IdxColumnName.H_POD10.ToString, Info.H_POD10,
      IdxColumnName.H_POD11.ToString, Info.H_POD11,
      IdxColumnName.H_POD12.ToString, Info.H_POD12,
      IdxColumnName.H_POD13.ToString, Info.H_POD13,
      IdxColumnName.H_POD14.ToString, Info.H_POD14,
      IdxColumnName.H_POD15.ToString, Info.H_POD15,
      IdxColumnName.H_POD16.ToString, Info.H_POD16,
      IdxColumnName.H_POD17.ToString, Info.H_POD17,
      IdxColumnName.H_POD18.ToString, Info.H_POD18,
      IdxColumnName.H_POD19.ToString, Info.H_POD19,
      IdxColumnName.H_POD20.ToString, Info.H_POD20,
      IdxColumnName.H_POD21.ToString, Info.H_POD21,
      IdxColumnName.H_POD22.ToString, Info.H_POD22,
      IdxColumnName.H_POD23.ToString, Info.H_POD23,
      IdxColumnName.H_POD24.ToString, Info.H_POD24,
      IdxColumnName.H_POD25.ToString, Info.H_POD25,
                             IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
                             IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
                             IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
                             IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
                             IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
                             IdxColumnName.PODTL_STATUS, CInt(Info.PODTL_STATUS),
                             IdxColumnName.CLOSE_ABLE, BooleanConvertToInteger(Info.CLOSE_ABLE)
     )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsPO_DTL) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.WORKING_TYPE.ToString, Info.WORKING_TYPE,
      IdxColumnName.WORKING_SERIAL_NO.ToString, Info.WORKING_SERIAL_NO,
      IdxColumnName.WORKING_SERIAL_SEQ.ToString, Info.WORKING_SERIAL_SEQ,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.LOT_NO.ToString, Info.LOT_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.PACKAGE_ID,
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
      IdxColumnName.STORAGE_TYPE.ToString, Info.STORAGE_TYPE,
      IdxColumnName.BND.ToString, Info.BND,
      IdxColumnName.QC_STATUS.ToString, Info.QC_STATUS,
      IdxColumnName.FROM_OWNER_ID.ToString, Info.FROM_OWNER_ID,
      IdxColumnName.FROM_SUB_OWNER_ID.ToString, Info.FROM_SUB_OWNER_ID,
      IdxColumnName.TO_OWNER_ID.ToString, Info.TO_OWNER_ID,
      IdxColumnName.TO_SUB_OWNER_ID.ToString, Info.TO_SUB_OWNER_ID,
      IdxColumnName.FACTORY_ID.ToString, Info.FACTORY_ID,
      IdxColumnName.DEST_AREA_ID.ToString, Info.DEST_AREA_ID,
      IdxColumnName.DEST_LOCATION_ID.ToString, Info.DEST_LOCATION_ID,
      IdxColumnName.H_POD_STEP_NO.ToString, Info.H_POD_STEP_NO,
      IdxColumnName.H_POD_MOVE_TYPE.ToString, Info.H_POD_MOVE_TYPE,
      IdxColumnName.H_POD_FINISH_TIME.ToString, Info.H_POD_FINISH_TIME,
      IdxColumnName.H_POD_BILLING_DATE.ToString, Info.H_POD_BILLING_DATE,
      IdxColumnName.H_POD_CREATE_TIME.ToString, Info.H_POD_CREATE_TIME,
      IdxColumnName.H_POD1.ToString, Info.H_POD1,
      IdxColumnName.H_POD2.ToString, Info.H_POD2,
      IdxColumnName.H_POD3.ToString, Info.H_POD3,
      IdxColumnName.H_POD4.ToString, Info.H_POD4,
      IdxColumnName.H_POD5.ToString, Info.H_POD5,
      IdxColumnName.H_POD6.ToString, Info.H_POD6,
      IdxColumnName.H_POD7.ToString, Info.H_POD7,
      IdxColumnName.H_POD8.ToString, Info.H_POD8,
      IdxColumnName.H_POD9.ToString, Info.H_POD9,
      IdxColumnName.H_POD10.ToString, Info.H_POD10,
      IdxColumnName.H_POD11.ToString, Info.H_POD11,
      IdxColumnName.H_POD12.ToString, Info.H_POD12,
      IdxColumnName.H_POD13.ToString, Info.H_POD13,
      IdxColumnName.H_POD14.ToString, Info.H_POD14,
      IdxColumnName.H_POD15.ToString, Info.H_POD15,
      IdxColumnName.H_POD16.ToString, Info.H_POD16,
      IdxColumnName.H_POD17.ToString, Info.H_POD17,
      IdxColumnName.H_POD18.ToString, Info.H_POD18,
      IdxColumnName.H_POD19.ToString, Info.H_POD19,
      IdxColumnName.H_POD20.ToString, Info.H_POD20,
      IdxColumnName.H_POD21.ToString, Info.H_POD21,
      IdxColumnName.H_POD22.ToString, Info.H_POD22,
      IdxColumnName.H_POD23.ToString, Info.H_POD23,
      IdxColumnName.H_POD24.ToString, Info.H_POD24,
      IdxColumnName.H_POD25.ToString, Info.H_POD25
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsPO_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}',{94}='{95}',{96}='{97}',{98}='{99}',{100}='{101}',{102}='{103}',{104}='{105}',{106}='{107}',{108}='{109}',{110}='{111}',{112}='{113}',{114}='{115}',{116}='{117}',{118}='{119}',{120}='{121}',{122}='{123}',{124}='{125}',{126}='{127}',{128}='{129}',{130}='{131}',{132}='{133}',{134}='{135}',{136}='{137}',{138}='{139}',{140}='{141}' WHERE {2}='{3}' And {6}='{7}'",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.WORKING_TYPE.ToString, Info.WORKING_TYPE,
      IdxColumnName.WORKING_SERIAL_NO.ToString, Info.WORKING_SERIAL_NO,
      IdxColumnName.WORKING_SERIAL_SEQ.ToString, Info.WORKING_SERIAL_SEQ,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.LOT_NO.ToString, Info.LOT_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.PACKAGE_ID,
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
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.STORAGE_TYPE),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
      IdxColumnName.FROM_OWNER_ID.ToString, Info.FROM_OWNER_ID,
      IdxColumnName.FROM_SUB_OWNER_ID.ToString, Info.FROM_SUB_OWNER_ID,
      IdxColumnName.TO_OWNER_ID.ToString, Info.TO_OWNER_ID,
      IdxColumnName.TO_SUB_OWNER_ID.ToString, Info.TO_SUB_OWNER_ID,
      IdxColumnName.FACTORY_ID.ToString, Info.FACTORY_ID,
      IdxColumnName.DEST_AREA_ID.ToString, Info.DEST_AREA_ID,
      IdxColumnName.DEST_LOCATION_ID.ToString, Info.DEST_LOCATION_ID,
      IdxColumnName.H_POD_STEP_NO.ToString, CInt(Info.H_POD_STEP_NO),
      IdxColumnName.H_POD_MOVE_TYPE.ToString, Info.H_POD_MOVE_TYPE,
      IdxColumnName.H_POD_FINISH_TIME.ToString, Info.H_POD_FINISH_TIME,
      IdxColumnName.H_POD_BILLING_DATE.ToString, Info.H_POD_BILLING_DATE,
      IdxColumnName.H_POD_CREATE_TIME.ToString, Info.H_POD_CREATE_TIME,
      IdxColumnName.H_POD1.ToString, Info.H_POD1,
      IdxColumnName.H_POD2.ToString, Info.H_POD2,
      IdxColumnName.H_POD3.ToString, Info.H_POD3,
      IdxColumnName.H_POD4.ToString, Info.H_POD4,
      IdxColumnName.H_POD5.ToString, Info.H_POD5,
      IdxColumnName.H_POD6.ToString, Info.H_POD6,
      IdxColumnName.H_POD7.ToString, Info.H_POD7,
      IdxColumnName.H_POD8.ToString, Info.H_POD8,
      IdxColumnName.H_POD9.ToString, Info.H_POD9,
      IdxColumnName.H_POD10.ToString, Info.H_POD10,
      IdxColumnName.H_POD11.ToString, Info.H_POD11,
      IdxColumnName.H_POD12.ToString, Info.H_POD12,
      IdxColumnName.H_POD13.ToString, Info.H_POD13,
      IdxColumnName.H_POD14.ToString, Info.H_POD14,
      IdxColumnName.H_POD15.ToString, Info.H_POD15,
      IdxColumnName.H_POD16.ToString, Info.H_POD16,
      IdxColumnName.H_POD17.ToString, Info.H_POD17,
      IdxColumnName.H_POD18.ToString, Info.H_POD18,
      IdxColumnName.H_POD19.ToString, Info.H_POD19,
      IdxColumnName.H_POD20.ToString, Info.H_POD20,
      IdxColumnName.H_POD21.ToString, Info.H_POD21,
      IdxColumnName.H_POD22.ToString, Info.H_POD22,
      IdxColumnName.H_POD23.ToString, Info.H_POD23,
      IdxColumnName.H_POD24.ToString, Info.H_POD24,
      IdxColumnName.H_POD25.ToString, Info.H_POD25,
                             IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
                             IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
                             IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
                             IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
                             IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
                             IdxColumnName.PODTL_STATUS, CInt(Info.PODTL_STATUS),
                             IdxColumnName.CLOSE_ABLE, BooleanConvertToInteger(Info.CLOSE_ABLE)
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function


  '- GET
  Public Shared Function GetclsWMS_T_PO_DTLListByALL() As List(Of clsPO_DTL)
    SyncLock objLock
      Try
        Dim _lstReturn As New List(Of clsPO_DTL)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {0}", TableName)
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_DTL
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
    End SyncLock
  End Function
  Public Shared Function GetclsWMS_T_PO_DTLListByPO_ID_PO_SERIAL_NO(ByVal po_id As String, po_serial_no As String) As List(Of clsPO_DTL)
    SyncLock objLock
      Try
        Dim _lstReturn As New List(Of clsPO_DTL)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4} = '{5}' ",
            strSQL,
            TableName,
            IdxColumnName.PO_ID.ToString, po_id,
            IdxColumnName.PO_SERIAL_NO.ToString, po_serial_no
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_DTL
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
    End SyncLock
  End Function
  Public Shared Function GetclsWMS_T_PO_DTLListByPO_ID(ByVal po_id As String) As Dictionary(Of String, clsPO_DTL)
    SyncLock objLock
      Try
        Dim ret_dic As New Dictionary(Of String, clsPO_DTL)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
            strSQL,
            TableName,
            IdxColumnName.PO_ID.ToString, po_id
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_DTL
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                  If Info IsNot Nothing Then
                    If ret_dic.ContainsKey(Info.gid) = False Then
                      ret_dic.Add(Info.gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Next
            End If
          End If
        End If
        Return ret_dic
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function

  '根據運單號
  Public Shared Function GetclsWMS_T_PO_DTLListByTKNUM(ByVal TKNUM As String) As Dictionary(Of String, clsPO_DTL)
    SyncLock objLock
      Try
        Dim ret_dic As New Dictionary(Of String, clsPO_DTL)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
            strSQL,
            TableName,
            IdxColumnName.H_POD12.ToString, TKNUM
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_DTL
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                  If Info IsNot Nothing Then
                    If ret_dic.ContainsKey(Info.gid) = False Then
                      ret_dic.Add(Info.gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Next
            End If
          End If
        End If
        Return ret_dic
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function

  Public Shared Function GetclsWMS_T_PO_DTLListByUniqueID(ByVal UniqueID As String) As Dictionary(Of String, clsPO_DTL)
    SyncLock objLock
      Try
        Dim ret_dic As New Dictionary(Of String, clsPO_DTL)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
            strSQL,
            TableName,
            IdxColumnName.H_POD1.ToString, UniqueID
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_DTL
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                  If Info IsNot Nothing Then
                    If ret_dic.ContainsKey(Info.gid) = False Then
                      ret_dic.Add(Info.gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Next
            End If
          End If
        End If
        Return ret_dic
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function


  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsPO_DTL, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_LINE_NO = "" & RowData.Item(IdxColumnName.PO_LINE_NO.ToString)
        Dim PO_SERIAL_NO = "" & RowData.Item(IdxColumnName.PO_SERIAL_NO.ToString)
        Dim WORKING_TYPE = "" & RowData.Item(IdxColumnName.WORKING_TYPE.ToString)
        Dim WORKING_SERIAL_NO = "" & RowData.Item(IdxColumnName.WORKING_SERIAL_NO.ToString)
        Dim WORKING_SERIAL_SEQ = "" & RowData.Item(IdxColumnName.WORKING_SERIAL_SEQ.ToString)
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim LOT_NO = "" & RowData.Item(IdxColumnName.LOT_NO.ToString)
        Dim QTY = 0 & RowData.Item(IdxColumnName.QTY.ToString)
        Dim QTY_PROCESS = 0 & RowData.Item(IdxColumnName.QTY_PROCESS.ToString)
        Dim QTY_FINISH = 0 & RowData.Item(IdxColumnName.QTY_FINISH.ToString)
        Dim COMMENTS = "" '移除了
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
        Dim STORAGE_TYPE = 0 & RowData.Item(IdxColumnName.STORAGE_TYPE.ToString)
        Dim BND = 0 & RowData.Item(IdxColumnName.BND.ToString)
        Dim QC_STATUS = 0 & RowData.Item(IdxColumnName.QC_STATUS.ToString)
        Dim FROM_OWNER_ID = "" & RowData.Item(IdxColumnName.FROM_OWNER_ID.ToString)
        Dim FROM_SUB_OWNER_ID = "" & RowData.Item(IdxColumnName.FROM_SUB_OWNER_ID.ToString)
        Dim TO_OWNER_ID = "" & RowData.Item(IdxColumnName.TO_OWNER_ID.ToString)
        Dim TO_SUB_OWNER_ID = "" & RowData.Item(IdxColumnName.TO_SUB_OWNER_ID.ToString)
        Dim FACTORY_ID = "" & RowData.Item(IdxColumnName.FACTORY_ID.ToString)
        Dim DEST_AREA_ID = "" & RowData.Item(IdxColumnName.DEST_AREA_ID.ToString)
        Dim DEST_LOCATION_ID = "" & RowData.Item(IdxColumnName.DEST_LOCATION_ID.ToString)
        Dim H_POD_STEP_NO = 0 & RowData.Item(IdxColumnName.H_POD_STEP_NO.ToString)
        Dim H_POD_MOVE_TYPE = "" & RowData.Item(IdxColumnName.H_POD_MOVE_TYPE.ToString)
        Dim H_POD_FINISH_TIME = "" & RowData.Item(IdxColumnName.H_POD_FINISH_TIME.ToString)
        Dim H_POD_BILLING_DATE = "" & RowData.Item(IdxColumnName.H_POD_BILLING_DATE.ToString)
        Dim H_POD_CREATE_TIME = "" & RowData.Item(IdxColumnName.H_POD_CREATE_TIME.ToString)
        Dim H_POD1 = "" & RowData.Item(IdxColumnName.H_POD1.ToString)
        Dim H_POD2 = "" & RowData.Item(IdxColumnName.H_POD2.ToString)
        Dim H_POD3 = "" & RowData.Item(IdxColumnName.H_POD3.ToString)
        Dim H_POD4 = "" & RowData.Item(IdxColumnName.H_POD4.ToString)
        Dim H_POD5 = "" & RowData.Item(IdxColumnName.H_POD5.ToString)
        Dim H_POD6 = "" & RowData.Item(IdxColumnName.H_POD6.ToString)
        Dim H_POD7 = "" & RowData.Item(IdxColumnName.H_POD7.ToString)
        Dim H_POD8 = "" & RowData.Item(IdxColumnName.H_POD8.ToString)
        Dim H_POD9 = "" & RowData.Item(IdxColumnName.H_POD9.ToString)
        Dim H_POD10 = "" & RowData.Item(IdxColumnName.H_POD10.ToString)
        Dim H_POD11 = "" & RowData.Item(IdxColumnName.H_POD11.ToString)
        Dim H_POD12 = "" & RowData.Item(IdxColumnName.H_POD12.ToString)
        Dim H_POD13 = "" & RowData.Item(IdxColumnName.H_POD13.ToString)
        Dim H_POD14 = "" & RowData.Item(IdxColumnName.H_POD14.ToString)
        Dim H_POD15 = "" & RowData.Item(IdxColumnName.H_POD15.ToString)
        Dim H_POD16 = "" & RowData.Item(IdxColumnName.H_POD16.ToString)
        Dim H_POD17 = "" & RowData.Item(IdxColumnName.H_POD17.ToString)
        Dim H_POD18 = "" & RowData.Item(IdxColumnName.H_POD18.ToString)
        Dim H_POD19 = "" & RowData.Item(IdxColumnName.H_POD19.ToString)
        Dim H_POD20 = "" & RowData.Item(IdxColumnName.H_POD20.ToString)
        Dim H_POD21 = "" & RowData.Item(IdxColumnName.H_POD21.ToString)
        Dim H_POD22 = "" & RowData.Item(IdxColumnName.H_POD22.ToString)
        Dim H_POD23 = "" & RowData.Item(IdxColumnName.H_POD23.ToString)
        Dim H_POD24 = "" & RowData.Item(IdxColumnName.H_POD24.ToString)
        Dim H_POD25 = "" & RowData.Item(IdxColumnName.H_POD25.ToString)

        Dim SORT_ITEM_COMMON1 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON1.ToString())
        Dim SORT_ITEM_COMMON2 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON2.ToString())
        Dim SORT_ITEM_COMMON3 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON3.ToString())
        Dim SORT_ITEM_COMMON4 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON4.ToString())
        Dim SORT_ITEM_COMMON5 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON5.ToString())
        Dim PODTL_STATUS = IIf(IsNumeric(RowData.Item(IdxColumnName.PODTL_STATUS.ToString)), RowData.Item(IdxColumnName.PODTL_STATUS.ToString), 0)
        Dim CLOSE_ABLE = IntegerConvertToBoolean(RowData.Item(IdxColumnName.CLOSE_ABLE.ToString()))
        Info = New clsPO_DTL(PO_ID, PO_LINE_NO, PO_SERIAL_NO, WORKING_TYPE, WORKING_SERIAL_NO, WORKING_SERIAL_SEQ, SKU_NO, LOT_NO, QTY, QTY_PROCESS, QTY_FINISH, COMMENTS, PACKAGE_ID, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6,
                             ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, STORAGE_TYPE, BND, QC_STATUS, FROM_OWNER_ID, FROM_SUB_OWNER_ID, TO_OWNER_ID, TO_SUB_OWNER_ID, FACTORY_ID, DEST_AREA_ID, DEST_LOCATION_ID,
                             H_POD_STEP_NO, H_POD_MOVE_TYPE, H_POD_FINISH_TIME, H_POD_BILLING_DATE, H_POD_CREATE_TIME, H_POD1, H_POD2, H_POD3, H_POD4, H_POD5, H_POD6, H_POD7, H_POD8, H_POD9, H_POD10, H_POD11, H_POD12, H_POD13, H_POD14, H_POD15, H_POD16, H_POD17, H_POD18, H_POD19, H_POD20, H_POD21, H_POD22, H_POD23, H_POD24, H_POD25, PODTL_STATUS, CLOSE_ABLE)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function SendSQLToDB(ByRef lstSQL As List(Of String)) As Boolean
    Try
      If lstSQL Is Nothing Then Return False
      If lstSQL.Count = 0 Then Return True
      For i = 0 To lstSQL.Count - 1
        SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next
      If fUseBatchUpdate_DynamicConnection = 0 Then
        For i = 0 To lstSQL.Count - 1
          DBTool.O_AddSQLQueue(TableName, lstSQL(i))
        Next
      Else
        Dim rtnMsg As String = DBTool.BatchUpdate_DynamicConnection(lstSQL)
        If rtnMsg.StartsWith("OK") Then
          SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        End If
      End If
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  '從資料庫抓取PO_DTL的資料
  Public Shared Function GetPO_DTLDictionaryByPOID_PO_Serial_No(ByVal PO_ID As String, ByVal PO_Serial_NO As String) As Dictionary(Of String, clsPO_DTL)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_DTL)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If PO_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_ID.ToString, PO_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_ID.ToString, PO_ID)
            End If
          End If
          If PO_Serial_NO <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_SERIAL_NO.ToString, PO_Serial_NO)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_SERIAL_NO.ToString, PO_Serial_NO)
            End If
          End If
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select * from {1} {2} ",
              strSQL,
            TableName,
            strWhere
            )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPO_DTL = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid()) = False Then
                    ret_dic.Add(Info.gid(), Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO_DTL Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO_DTL Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從資料庫抓取PO_DTL的資料
  Public Shared Function GetPO_DTLDictionaryByPOID_LINE_NO_Serial_No(ByVal PO_ID As String, ByVal PO_Line_NO As String, ByVal PO_Serial_NO As String) As Dictionary(Of String, clsPO_DTL)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_DTL)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If PO_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_ID.ToString, PO_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_ID.ToString, PO_ID)
            End If
          End If
          If PO_Serial_NO <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_SERIAL_NO.ToString, PO_Serial_NO)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_SERIAL_NO.ToString, PO_Serial_NO)
            End If
          End If
          If PO_Line_NO <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_LINE_NO.ToString, PO_Line_NO)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_LINE_NO.ToString, PO_Line_NO)
            End If
          End If
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select * from {1} {2} ",
              strSQL,
            TableName,
            strWhere
            )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPO_DTL = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid()) = False Then
                    ret_dic.Add(Info.gid(), Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO_DTL Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO_DTL Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從資料庫抓取PO_DTL的資料
  Public Shared Function GetPO_DTLDictionaryByPOID_PO_Line_No(ByVal PO_ID As String, ByVal PO_Line_NO As String) As Dictionary(Of String, clsPO_DTL)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_DTL)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If PO_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_ID.ToString, PO_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_ID.ToString, PO_ID)
            End If
          End If
          If PO_Line_NO <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_LINE_NO.ToString, PO_Line_NO)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_LINE_NO.ToString, PO_Line_NO)
            End If
          End If
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select * from {1} {2} ",
              strSQL,
            TableName,
            strWhere
            )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPO_DTL = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid()) = False Then
                    ret_dic.Add(Info.gid(), Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO_DTL Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO_DTL Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從資料庫抓取PO_DTL的資料
  Public Shared Function GetPO_DTLDictionaryBydicPOID(ByVal dicPOID As Dictionary(Of String, String)) As Dictionary(Of String, clsPO_DTL)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_DTL)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          Dim strPOList As String = ""
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          'For Each PO_ID As String In dicPOID.Values
          '  If strPOList = "" Then
          '    strPOList = "'" & PO_ID & "'"
          '  Else
          '    strPOList = strPOList & ",'" & PO_ID & "'"
          '  End If
          'Next
          'If strWhere = "" Then
          '  strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.PO_ID.ToString, strPOList)
          'Else
          '  strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.PO_ID.ToString, strPOList)
          'End If
          'Dim strSQL As String = String.Empty
          'Dim DatasetMessage As New DataSet
          'strSQL = String.Format("Select * from {1} {2} ",
          '    strSQL,
          '  TableName,
          '  strWhere
          '  )
          'SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          'DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          'If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          '  For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
          '    Dim Info As clsPO_DTL = Nothing
          '    If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
          '      If Info IsNot Nothing Then
          '        If ret_dic.ContainsKey(Info.gid()) = False Then
          '          ret_dic.Add(Info.gid(), Info)
          '        End If
          '      Else
          '        SendMessageToLog("Get clsPO_DTL Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '      End If
          '    Else
          '      SendMessageToLog("Get clsPO_DTL Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '    End If

          '  Next
          'End If

          Dim count_flag = 0
          For i = 0 To dicPOID.Count - 1
            If strPOList = "" Then
              strPOList = "'" & dicPOID.Keys(i) & "'"
            Else
              strPOList = strPOList & ",'" & dicPOID.Keys(i) & "'"
            End If
            If i - count_flag > 800 OrElse i = (dicPOID.Count - 1) Then
              count_flag = i
              strWhere = ""
              If strWhere = "" Then
                strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.PO_ID.ToString, strPOList)
              Else
                strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.PO_ID.ToString, strPOList)
              End If
              strSQL = String.Format("Select * from {1} {2} ",
                  strSQL,
                  TableName,
                  strWhere
              )
              SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
              If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
                For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                  Dim Info As clsPO_DTL = Nothing
                  SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Next
              End If
              strPOList = ""
            End If
          Next



        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從資料庫抓取PO_DTL的資料
  Public Shared Function GetPO_DTLDictionaryByALL() As Dictionary(Of String, clsPO_DTL)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_DTL)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then

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
              Dim Info As clsPO_DTL = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid()) = False Then
                    ret_dic.Add(Info.gid(), Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO_DTL Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO_DTL Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
