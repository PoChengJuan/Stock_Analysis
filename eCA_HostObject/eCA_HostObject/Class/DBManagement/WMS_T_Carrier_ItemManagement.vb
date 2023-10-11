Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_T_Carrier_ItemManagement
  Public Shared TableName As String = "WMS_T_Carrier_Item"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsCarrierItem)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
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
    SORT_ITEM_COMMON1
    SORT_ITEM_COMMON2
    SORT_ITEM_COMMON3
    SORT_ITEM_COMMON4
    SORT_ITEM_COMMON5
    QTY
    OWNER_NO
    SUB_OWNER_NO
    RECEIPT_DATE
    MANUFACETURE_DATE
    EXPIRED_DATE
    ORG_EXPIRED_DATE
    EXPIRED_COMMENTS
    EXPIRED_WARNING_FLAG
    CREATE_TIME
    IN_TIME
    IN_CLIENT_ID
    RECEIPT_WO_ID
    RECEIPT_WO_SERIAL_NO
    PICKING_WO_ID
    PICKING_WO_SERIAL_NO
    ACCEPTING_STATUS
    STORAGE_TYPE
    BND
    QC_STATUS
    QC_TIME
    TO_ERP
    TO_ERP_TIME
    EFFECTIVE_DATE
  End Enum

  '- GetSQL
  '-請將 clsCarrierItem 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱



  Public Shared Function GetInsertSQL(ByRef Info As clsCarrierItem) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}',{41},'{43}','{45}','{47}','{49}','{51}','{53}','{55}',{57},'{59}','{61}','{63}','{65}','{67}','{69}',{71},{73},{75},{77},'{79}',{81},'{83}')",
      strSQL,
      TableName,
      IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
      IdxColumnName.PACKAGE_ID.ToString, Info.Package_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_No,
      IdxColumnName.LOT_NO.ToString, Info.Lot_No,
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
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.OWNER_NO.ToString, Info.Owner_No,
      IdxColumnName.SUB_OWNER_NO.ToString, Info.Sub_Owner_No,
      IdxColumnName.RECEIPT_DATE.ToString, Info.Receipt_Date,
      IdxColumnName.MANUFACETURE_DATE.ToString, Info.Manufaceture_Date,
      IdxColumnName.EXPIRED_DATE.ToString, Info.Expired_Date,
      IdxColumnName.ORG_EXPIRED_DATE.ToString, Info.Org_Expired_Date,
      IdxColumnName.EXPIRED_COMMENTS.ToString, Info.Expired_Comments,
      IdxColumnName.EXPIRED_WARNING_FLAG.ToString, BooleanConvertToInteger(Info.Expired_Warning_Flag),
      IdxColumnName.EFFECTIVE_DATE.ToString, Info.Effective_Date,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.IN_TIME.ToString, Info.In_Time,
      IdxColumnName.IN_CLIENT_ID.ToString, Info.In_Client_ID,
      IdxColumnName.RECEIPT_WO_ID.ToString, Info.Receipt_WO_ID,
      IdxColumnName.RECEIPT_WO_SERIAL_NO.ToString, Info.Receipt_WO_Serial_No,
      IdxColumnName.ACCEPTING_STATUS.ToString, CInt(Info.Accepting_Status),
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.Storage_Type),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_Status),
      IdxColumnName.QC_TIME.ToString, Info.QC_Time,
      IdxColumnName.TO_ERP.ToString, BooleanConvertToInteger(Info.To_ERP),
      IdxColumnName.TO_ERP_TIME.ToString, Info.To_ERP_Time
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsCarrierItem) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' AND {10}='{11}' AND {12}='{13}' AND {14}='{15}' AND {16}='{17}' AND {18}='{19}' AND {20}='{21}' AND {22}='{23}' AND {24}='{25}' AND {26}='{27}' AND {28}='{29}' AND {30}='{31}' AND {32}='{33}' AND {34}='{35}' AND {36}='{37}' AND {38}='{39}' AND {42}='{43}' AND {44}='{45}' AND {66}='{67}' AND {68}='{69}' AND {72}={73} AND {74}={75} AND {76}={77} ",
      strSQL,
      TableName,
      IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
      IdxColumnName.PACKAGE_ID.ToString, Info.Package_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_No,
      IdxColumnName.LOT_NO.ToString, Info.Lot_No,
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
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.OWNER_NO.ToString, Info.Owner_No,
      IdxColumnName.SUB_OWNER_NO.ToString, Info.Sub_Owner_No,
      IdxColumnName.RECEIPT_DATE.ToString, Info.Receipt_Date,
      IdxColumnName.MANUFACETURE_DATE.ToString, Info.Manufaceture_Date,
      IdxColumnName.EXPIRED_DATE.ToString, Info.Expired_Date,
      IdxColumnName.ORG_EXPIRED_DATE.ToString, Info.Org_Expired_Date,
      IdxColumnName.EXPIRED_COMMENTS.ToString, Info.Expired_Comments,
      IdxColumnName.EXPIRED_WARNING_FLAG.ToString, Info.Expired_Warning_Flag,
      IdxColumnName.EFFECTIVE_DATE.ToString, Info.Effective_Date,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.IN_TIME.ToString, Info.In_Time,
      IdxColumnName.IN_CLIENT_ID.ToString, Info.In_Client_ID,
      IdxColumnName.RECEIPT_WO_ID.ToString, Info.Receipt_WO_ID,
      IdxColumnName.RECEIPT_WO_SERIAL_NO.ToString, Info.Receipt_WO_Serial_No,
      IdxColumnName.ACCEPTING_STATUS.ToString, CInt(Info.Accepting_Status),
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.Storage_Type),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_Status),
      IdxColumnName.QC_TIME.ToString, Info.QC_Time,
      IdxColumnName.TO_ERP.ToString, BooleanConvertToInteger(Info.To_ERP),
      IdxColumnName.TO_ERP_TIME.ToString, Info.To_ERP_Time
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsCarrierItem) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {40}={41},{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}={57},{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{70}={71},{78}='{79}',{80}={81},{82}='{83}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}' And {10}='{11}' And {12}='{13}' And {14}='{15}' And {16}='{17}' And {18}='{19}' And {20}='{21}' And {22}='{23}' And {24}='{25}' And {26}='{27}' And {28}='{29}' And {30}='{31}' And {32}='{33}' And {34}='{35}' And {36}='{37}' And {38}='{39}' And {42}='{43}' And {44}='{45}' And {66}='{67}' And {68}='{69}' And {72}={73} And {74}={75} And {76}={77}",
      strSQL,
      TableName,
      IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
      IdxColumnName.PACKAGE_ID.ToString, Info.Package_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_No,
      IdxColumnName.LOT_NO.ToString, Info.Lot_No,
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
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.OWNER_NO.ToString, Info.Owner_No,
      IdxColumnName.SUB_OWNER_NO.ToString, Info.Sub_Owner_No,
      IdxColumnName.RECEIPT_DATE.ToString, Info.Receipt_Date,
      IdxColumnName.MANUFACETURE_DATE.ToString, Info.Manufaceture_Date,
      IdxColumnName.EXPIRED_DATE.ToString, Info.Expired_Date,
      IdxColumnName.ORG_EXPIRED_DATE.ToString, Info.Org_Expired_Date,
      IdxColumnName.EXPIRED_COMMENTS.ToString, Info.Expired_Comments,
      IdxColumnName.EXPIRED_WARNING_FLAG.ToString, BooleanConvertToInteger(Info.Expired_Warning_Flag),
      IdxColumnName.EFFECTIVE_DATE.ToString, Info.Effective_Date,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.IN_TIME.ToString, Info.In_Time,
      IdxColumnName.IN_CLIENT_ID.ToString, Info.In_Client_ID,
      IdxColumnName.RECEIPT_WO_ID.ToString, Info.Receipt_WO_ID,
      IdxColumnName.RECEIPT_WO_SERIAL_NO.ToString, Info.Receipt_WO_Serial_No,
      IdxColumnName.ACCEPTING_STATUS.ToString, CInt(Info.Accepting_Status),
      IdxColumnName.STORAGE_TYPE.ToString, CInt(Info.Storage_Type),
      IdxColumnName.BND.ToString, BooleanConvertToInteger(Info.BND),
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_Status),
      IdxColumnName.QC_TIME.ToString, Info.QC_Time,
      IdxColumnName.TO_ERP.ToString, BooleanConvertToInteger(Info.To_ERP),
      IdxColumnName.TO_ERP_TIME.ToString, Info.To_ERP_Time
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

  '- GET DB Data
  Public Shared Function GetWMS_T_Carrier_Item_ALL() As Dictionary(Of String, clsCarrierItem)
    Try
      Dim dicCarrierItem As New Dictionary(Of String, clsCarrierItem)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1}",
            strSQL,
            TableName
            )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCarrierItem = Nothing
            If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
              If Info IsNot Nothing Then
                If dicCarrierItem.ContainsKey(Info.gid) = False Then
                  dicCarrierItem.Add(Info.gid, Info)
                End If
              Else
                SendMessageToLog("Get clsCarrierItem Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Else
              SendMessageToLog("Get clsCarrierItem Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            End If

          Next
        End If
        'End If
      End If
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetWMS_T_Carrier_ItemDataListByCarrierID_ReceiptWOID_WOSeriaNo_SKUNo_PackageID(ByVal CarrierID As String, ByVal ReceiptWOID As String, ByVal ReceiptSerialNo As String, ByVal SKUNo As String, ByVal PackageID As String) As Concurrent.ConcurrentDictionary(Of String, clsCarrierItem)
    Try
      Dim dicCarrierItem As New Concurrent.ConcurrentDictionary(Of String, clsCarrierItem)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim strSQLWhere As String = ""
        If CarrierID <> "" Then
          If strSQLWhere = "" Then
            strSQLWhere = String.Format(" WHERE {0}='{1}' ", IdxColumnName.CARRIER_ID.ToString, CarrierID)
          Else
            strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.CARRIER_ID.ToString, CarrierID)
          End If
        End If
        If ReceiptWOID <> "" Then
          If strSQLWhere = "" Then
            strSQLWhere = String.Format(" WHERE {0}='{1}' ", IdxColumnName.RECEIPT_WO_ID.ToString, ReceiptWOID)
          Else
            strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.RECEIPT_WO_ID.ToString, ReceiptWOID)
          End If
        End If
        If ReceiptSerialNo <> "" Then
          If strSQLWhere = "" Then
            strSQLWhere = String.Format(" WHERE {0}='{1}' ", IdxColumnName.RECEIPT_WO_SERIAL_NO.ToString, ReceiptSerialNo)
          Else
            strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.RECEIPT_WO_SERIAL_NO.ToString, ReceiptSerialNo)
          End If
        End If
        If SKUNo <> "" Then
          If strSQLWhere = "" Then
            strSQLWhere = String.Format(" WHERE {0}='{1}' ", IdxColumnName.SKU_NO.ToString, SKUNo)
          Else
            strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.SKU_NO.ToString, SKUNo)
          End If
        End If
        If PackageID <> "" Then
          If strSQLWhere = "" Then
            strSQLWhere = String.Format(" WHERE {0}='{1}' ", IdxColumnName.PACKAGE_ID.ToString, PackageID)
          Else
            strSQLWhere = strSQLWhere & String.Format(" AND {0}='{1}' ", IdxColumnName.PACKAGE_ID.ToString, PackageID)
          End If
        End If
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} {2} ",
            strSQL,
            TableName,
            strSQLWhere
            )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCarrierItem = Nothing
            If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
              If Info IsNot Nothing Then
                If dicCarrierItem.ContainsKey(Info.gid) = False Then
                  dicCarrierItem.TryAdd(Info.gid, Info)
                End If
              Else
                SendMessageToLog("Get clsCarrierItem Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Else
              SendMessageToLog("Get clsCarrierItem Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            End If

          Next
        End If
        'End If
      End If
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetWMS_T_Carrier_ItemDataListByACCEPTING_STATUS_ItemCommon2IsNotNull(ByVal ACCEPTING_STATUS As enuAcceptingStatus) As Dictionary(Of String, clsCarrierItem)
    Try
      Dim dicCarrierItem As New Dictionary(Of String, clsCarrierItem)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim strSQLWhere As String = ""

        strSQLWhere = String.Format(" WHERE {0}={1} ", IdxColumnName.ACCEPTING_STATUS.ToString, CInt(ACCEPTING_STATUS))

        'Item_Common2 不能为空
        If strSQLWhere = "" Then
          strSQLWhere = String.Format(" WHERE {0} is not null ", IdxColumnName.ITEM_COMMON2.ToString)
        Else
          strSQLWhere = strSQLWhere & String.Format(" AND {0} is not null ", IdxColumnName.ITEM_COMMON2.ToString)
        End If
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} {2} ",
    strSQL,
    TableName,
    strSQLWhere
    )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCarrierItem = Nothing
            If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
              If Info IsNot Nothing Then
                If dicCarrierItem.ContainsKey(Info.gid) = False Then
                  dicCarrierItem.Add(Info.gid, Info)
                End If
              Else
                SendMessageToLog("Get clsCarrierItem Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Else
              SendMessageToLog("Get clsCarrierItem Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            End If

          Next
        End If
        'End If
      End If
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetWMS_T_Carrier_ItemDataListByACCEPTING_STATUS(ByVal ACCEPTING_STATUS As enuAcceptingStatus) As Dictionary(Of String, clsCarrierItem)
    Try
      Dim dicCarrierItem As New Dictionary(Of String, clsCarrierItem)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim strSQLWhere As String = ""

        strSQLWhere = String.Format(" WHERE {0}={1} ", IdxColumnName.ACCEPTING_STATUS.ToString, CInt(ACCEPTING_STATUS))

        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} {2} order by {3} ",
            strSQL,
            TableName,
            strSQLWhere,
            IdxColumnName.RECEIPT_DATE
            )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCarrierItem = Nothing
            If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
              If Info IsNot Nothing Then
                If dicCarrierItem.ContainsKey(Info.gid) = False Then
                  dicCarrierItem.Add(Info.gid, Info)
                End If
              Else
                SendMessageToLog("Get clsCarrierItem Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Else
              SendMessageToLog("Get clsCarrierItem Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            End If

          Next
        End If
        'End If
      End If
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  Public Shared Function GetWMS_T_Carrier_ItemDataListByItemCommon2SN(ByVal ItemCommon2 As String, ByVal PACKAGE_ID As String) As Dictionary(Of String, clsCarrierItem)
    Try
      Dim dicCarrierItem As New Dictionary(Of String, clsCarrierItem)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim strSQLWhere As String = ""

        'strSQLWhere = String.Format(" WHERE {0}={1} ", IdxColumnName.ACCEPTING_STATUS.ToString, CInt(ACCEPTING_STATUS))

        'Item_Common2 
        If strSQLWhere = "" Then
          strSQLWhere = String.Format(" WHERE {0} = '{1}' ", IdxColumnName.ITEM_COMMON2.ToString, ItemCommon2)
        Else
          strSQLWhere = strSQLWhere & String.Format(" AND {0} = '{1}' ", IdxColumnName.ITEM_COMMON2.ToString, ItemCommon2)
        End If

        'SN 
        If strSQLWhere = "" Then
          strSQLWhere = String.Format(" WHERE {0} = '{1}' ", IdxColumnName.PACKAGE_ID.ToString, PACKAGE_ID)
        Else
          strSQLWhere = strSQLWhere & String.Format(" AND {0} = '{1}' ", IdxColumnName.PACKAGE_ID.ToString, PACKAGE_ID)
        End If

        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} {2} ",
    strSQL,
    TableName,
    strSQLWhere
    )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCarrierItem = Nothing
            If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
              If Info IsNot Nothing Then
                If dicCarrierItem.ContainsKey(Info.gid) = False Then
                  dicCarrierItem.Add(Info.gid, Info)
                End If
              Else
                SendMessageToLog("Get clsCarrierItem Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Else
              SendMessageToLog("Get clsCarrierItem Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            End If

          Next
        End If
        'End If
      End If
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetWMS_T_Carrier_ItemDataListByACCEPTING_STATUS_SortItemCommon4Is101(ByVal ACCEPTING_STATUS As enuAcceptingStatus) As Dictionary(Of String, clsCarrierItem)
    Try
      Dim dicCarrierItem As New Dictionary(Of String, clsCarrierItem)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim strSQLWhere As String = ""

        strSQLWhere = String.Format(" WHERE {0}={1} ", IdxColumnName.ACCEPTING_STATUS.ToString, CInt(ACCEPTING_STATUS))

        'SORT_ITEM_COMMON4 = 101
        If strSQLWhere = "" Then
          strSQLWhere = String.Format(" WHERE {0}='101' ", IdxColumnName.SORT_ITEM_COMMON4.ToString)
        Else
          strSQLWhere = strSQLWhere & String.Format(" AND {0}='101' ", IdxColumnName.SORT_ITEM_COMMON4.ToString)
        End If
        'Item_Common2 为空
        If strSQLWhere = "" Then
          strSQLWhere = String.Format(" WHERE {0} is null ", IdxColumnName.ITEM_COMMON2.ToString)
        Else
          strSQLWhere = strSQLWhere & String.Format(" AND {0} is null ", IdxColumnName.ITEM_COMMON2.ToString)
        End If
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} {2} ",
    strSQL,
    TableName,
    strSQLWhere
    )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCarrierItem = Nothing
            If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
              If Info IsNot Nothing Then
                If dicCarrierItem.ContainsKey(Info.gid) = False Then
                  dicCarrierItem.Add(Info.gid, Info)
                End If
              Else
                SendMessageToLog("Get clsCarrierItem Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Else
              SendMessageToLog("Get clsCarrierItem Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            End If

          Next
        End If
        'End If
      End If
      Return dicCarrierItem
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsCarrierItem, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim CARRIER_ID = "" & RowData.Item(IdxColumnName.CARRIER_ID.ToString)
        Dim PACKAGE_ID = "" & RowData.Item(IdxColumnName.PACKAGE_ID.ToString)
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
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
        Dim QTY = 0 & RowData.Item(IdxColumnName.QTY.ToString)
        Dim OWNER_NO = "" & RowData.Item(IdxColumnName.OWNER_NO.ToString)
        Dim SUB_OWNER_NO = "" & RowData.Item(IdxColumnName.SUB_OWNER_NO.ToString)
        Dim RECEIPT_DATE = "" & RowData.Item(IdxColumnName.RECEIPT_DATE.ToString)
        Dim MANUFACETURE_DATE = "" & RowData.Item(IdxColumnName.MANUFACETURE_DATE.ToString)
        Dim EXPIRED_DATE = "" & RowData.Item(IdxColumnName.EXPIRED_DATE.ToString)
        Dim ORG_EXPIRED_DATE = "" & RowData.Item(IdxColumnName.ORG_EXPIRED_DATE.ToString)
        Dim EXPIRED_COMMENTS = "" & RowData.Item(IdxColumnName.EXPIRED_COMMENTS.ToString)
        Dim EXPIRED_WARNING_FLAG = IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.EXPIRED_WARNING_FLAG.ToString))
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim IN_TIME = "" & RowData.Item(IdxColumnName.IN_TIME.ToString)
        Dim IN_CLIENT_ID = "" & RowData.Item(IdxColumnName.IN_CLIENT_ID.ToString)
        Dim RECEIPT_WO_ID = "" & RowData.Item(IdxColumnName.RECEIPT_WO_ID.ToString)
        Dim RECEIPT_WO_SERIAL_NO = "" & RowData.Item(IdxColumnName.RECEIPT_WO_SERIAL_NO.ToString)
        Dim ACCEPTING_STATUS = 0 & RowData.Item(IdxColumnName.ACCEPTING_STATUS.ToString)
        Dim TEMPORART_STORAGE = 0 & RowData.Item(IdxColumnName.STORAGE_TYPE.ToString)
        Dim BND = IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.BND.ToString))
        Dim QC_STATUS = 0 & RowData.Item(IdxColumnName.QC_STATUS.ToString)
        Dim QC_TIME = "" & RowData.Item(IdxColumnName.QC_TIME.ToString)
        Dim TO_ERP = IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.TO_ERP.ToString))
        Dim TO_ERP_TIME = "" & RowData.Item(IdxColumnName.TO_ERP_TIME.ToString)
        Dim EFFECTIVE_DATE = "" & RowData.Item(IdxColumnName.EFFECTIVE_DATE.ToString)
        Info = New clsCarrierItem(CARRIER_ID, PACKAGE_ID, SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10,
                                  SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5,
                                  QTY, OWNER_NO, SUB_OWNER_NO, RECEIPT_DATE, MANUFACETURE_DATE, EXPIRED_DATE, ORG_EXPIRED_DATE, EXPIRED_COMMENTS, EXPIRED_WARNING_FLAG, CREATE_TIME, IN_TIME, IN_CLIENT_ID, RECEIPT_WO_ID, RECEIPT_WO_SERIAL_NO,
                                  ACCEPTING_STATUS, TEMPORART_STORAGE, BND, QC_STATUS, QC_TIME, TO_ERP, TO_ERP_TIME, EFFECTIVE_DATE)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
