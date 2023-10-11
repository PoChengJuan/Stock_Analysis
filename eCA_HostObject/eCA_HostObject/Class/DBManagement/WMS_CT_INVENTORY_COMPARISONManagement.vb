Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_CT_INVENTORY_COMPARISONManagement
  Public Shared TableName As String = "WMS_CT_INVENTORY_COMPARISON"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsWMS_CT_INVENTORY_COMPARISON)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
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
    STORAGE_TYPE
    BND
    QC_STATUS
    WMS_STOCK_QTY
    WMS_UNFINISH_QTY
    WMS_COMPARSON_QTY
    ERP_STOCK_QTY
    ERP_UNFINISH_QTY
    ERP_COMPARSON_QTY
    QUANTITY_VARIANCE
    ERP_SYSTEM
    CREATE_TIME
    ACC_COMMON1
    ACC_COMMON2
    ACC_COMMON3
    ACC_COMMON4
    ACC_COMMON5
    ACC_COMMON6
    ACC_COMMON7
    ACC_COMMON8
    ACC_COMMON9
    ACC_COMMON10
  End Enum

  Public Enum UpdateOption As Integer
    UpdateDic = 0
    UpdateDB = 1
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsWMS_CT_INVENTORY_COMPARISON) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}',{41},{43},{45},{47},{49},{51},{53},{55},{57},{59},'{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}')",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.LOT_NO.ToString, CI.LOT_NO,
      IdxColumnName.ITEM_COMMON1.ToString, CI.ITEM_COMMON1,
      IdxColumnName.ITEM_COMMON2.ToString, CI.ITEM_COMMON2,
      IdxColumnName.ITEM_COMMON3.ToString, CI.ITEM_COMMON3,
      IdxColumnName.ITEM_COMMON4.ToString, CI.ITEM_COMMON4,
      IdxColumnName.ITEM_COMMON5.ToString, CI.ITEM_COMMON5,
      IdxColumnName.ITEM_COMMON6.ToString, CI.ITEM_COMMON6,
      IdxColumnName.ITEM_COMMON7.ToString, CI.ITEM_COMMON7,
      IdxColumnName.ITEM_COMMON8.ToString, CI.ITEM_COMMON8,
      IdxColumnName.ITEM_COMMON9.ToString, CI.ITEM_COMMON9,
      IdxColumnName.ITEM_COMMON10.ToString, CI.ITEM_COMMON10,
      IdxColumnName.SORT_ITEM_COMMON1.ToString, CI.SORT_ITEM_COMMON1,
      IdxColumnName.SORT_ITEM_COMMON2.ToString, CI.SORT_ITEM_COMMON2,
      IdxColumnName.SORT_ITEM_COMMON3.ToString, CI.SORT_ITEM_COMMON3,
      IdxColumnName.SORT_ITEM_COMMON4.ToString, CI.SORT_ITEM_COMMON4,
      IdxColumnName.SORT_ITEM_COMMON5.ToString, CI.SORT_ITEM_COMMON5,
      IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
      IdxColumnName.SUB_OWNER_NO.ToString, CI.SUB_OWNER_NO,
      IdxColumnName.STORAGE_TYPE.ToString, CI.STORAGE_TYPE,
      IdxColumnName.BND.ToString, CI.BND,
      IdxColumnName.QC_STATUS.ToString, CI.QC_STATUS,
      IdxColumnName.WMS_STOCK_QTY.ToString, CI.WMS_STOCK_QTY,
      IdxColumnName.WMS_UNFINISH_QTY.ToString, CI.WMS_UNFINISH_QTY,
      IdxColumnName.WMS_COMPARSON_QTY.ToString, CI.WMS_COMPARSON_QTY,
      IdxColumnName.ERP_STOCK_QTY.ToString, CI.ERP_STOCK_QTY,
      IdxColumnName.ERP_UNFINISH_QTY.ToString, CI.ERP_UNFINISH_QTY,
      IdxColumnName.ERP_COMPARSON_QTY.ToString, CI.ERP_COMPARSON_QTY,
      IdxColumnName.QUANTITY_VARIANCE.ToString, CI.QUANTITY_VARIANCE,
      IdxColumnName.ERP_SYSTEM.ToString, CI.ERP_SYSTEM,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
      IdxColumnName.ACC_COMMON1.ToString, CI.ACC_COMMON1,
      IdxColumnName.ACC_COMMON2.ToString, CI.ACC_COMMON2,
      IdxColumnName.ACC_COMMON3.ToString, CI.ACC_COMMON3,
      IdxColumnName.ACC_COMMON4.ToString, CI.ACC_COMMON4,
      IdxColumnName.ACC_COMMON5.ToString, CI.ACC_COMMON5,
      IdxColumnName.ACC_COMMON6.ToString, CI.ACC_COMMON6,
      IdxColumnName.ACC_COMMON7.ToString, CI.ACC_COMMON7,
      IdxColumnName.ACC_COMMON8.ToString, CI.ACC_COMMON8,
      IdxColumnName.ACC_COMMON9.ToString, CI.ACC_COMMON9,
      IdxColumnName.ACC_COMMON10.ToString, CI.ACC_COMMON10
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
  Public Shared Function GetDeleteSQL(ByRef CI As clsWMS_CT_INVENTORY_COMPARISON) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} ",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.LOT_NO.ToString, CI.LOT_NO,
      IdxColumnName.ITEM_COMMON1.ToString, CI.ITEM_COMMON1,
      IdxColumnName.ITEM_COMMON2.ToString, CI.ITEM_COMMON2,
      IdxColumnName.ITEM_COMMON3.ToString, CI.ITEM_COMMON3,
      IdxColumnName.ITEM_COMMON4.ToString, CI.ITEM_COMMON4,
      IdxColumnName.ITEM_COMMON5.ToString, CI.ITEM_COMMON5,
      IdxColumnName.ITEM_COMMON6.ToString, CI.ITEM_COMMON6,
      IdxColumnName.ITEM_COMMON7.ToString, CI.ITEM_COMMON7,
      IdxColumnName.ITEM_COMMON8.ToString, CI.ITEM_COMMON8,
      IdxColumnName.ITEM_COMMON9.ToString, CI.ITEM_COMMON9,
      IdxColumnName.ITEM_COMMON10.ToString, CI.ITEM_COMMON10,
      IdxColumnName.SORT_ITEM_COMMON1.ToString, CI.SORT_ITEM_COMMON1,
      IdxColumnName.SORT_ITEM_COMMON2.ToString, CI.SORT_ITEM_COMMON2,
      IdxColumnName.SORT_ITEM_COMMON3.ToString, CI.SORT_ITEM_COMMON3,
      IdxColumnName.SORT_ITEM_COMMON4.ToString, CI.SORT_ITEM_COMMON4,
      IdxColumnName.SORT_ITEM_COMMON5.ToString, CI.SORT_ITEM_COMMON5,
      IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
      IdxColumnName.SUB_OWNER_NO.ToString, CI.SUB_OWNER_NO,
      IdxColumnName.STORAGE_TYPE.ToString, CI.STORAGE_TYPE,
      IdxColumnName.BND.ToString, CI.BND,
      IdxColumnName.QC_STATUS.ToString, CI.QC_STATUS,
      IdxColumnName.WMS_STOCK_QTY.ToString, CI.WMS_STOCK_QTY,
      IdxColumnName.WMS_UNFINISH_QTY.ToString, CI.WMS_UNFINISH_QTY,
      IdxColumnName.WMS_COMPARSON_QTY.ToString, CI.WMS_COMPARSON_QTY,
      IdxColumnName.ERP_STOCK_QTY.ToString, CI.ERP_STOCK_QTY,
      IdxColumnName.ERP_UNFINISH_QTY.ToString, CI.ERP_UNFINISH_QTY,
      IdxColumnName.ERP_COMPARSON_QTY.ToString, CI.ERP_COMPARSON_QTY,
      IdxColumnName.QUANTITY_VARIANCE.ToString, CI.QUANTITY_VARIANCE,
      IdxColumnName.ERP_SYSTEM.ToString, CI.ERP_SYSTEM,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
      IdxColumnName.ACC_COMMON1.ToString, CI.ACC_COMMON1,
      IdxColumnName.ACC_COMMON2.ToString, CI.ACC_COMMON2,
      IdxColumnName.ACC_COMMON3.ToString, CI.ACC_COMMON3,
      IdxColumnName.ACC_COMMON4.ToString, CI.ACC_COMMON4,
      IdxColumnName.ACC_COMMON5.ToString, CI.ACC_COMMON5,
      IdxColumnName.ACC_COMMON6.ToString, CI.ACC_COMMON6,
      IdxColumnName.ACC_COMMON7.ToString, CI.ACC_COMMON7,
      IdxColumnName.ACC_COMMON8.ToString, CI.ACC_COMMON8,
      IdxColumnName.ACC_COMMON9.ToString, CI.ACC_COMMON9,
      IdxColumnName.ACC_COMMON10.ToString, CI.ACC_COMMON10
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
  Public Shared Function GetUpdateSQL(ByRef CI As clsWMS_CT_INVENTORY_COMPARISON) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {2}='{3}',{4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}={41},{42}={43},{44}={45},{46}={47},{48}={49},{50}={51},{52}={53},{54}={55},{56}={57},{58}={59},{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}' WHERE ",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.LOT_NO.ToString, CI.LOT_NO,
      IdxColumnName.ITEM_COMMON1.ToString, CI.ITEM_COMMON1,
      IdxColumnName.ITEM_COMMON2.ToString, CI.ITEM_COMMON2,
      IdxColumnName.ITEM_COMMON3.ToString, CI.ITEM_COMMON3,
      IdxColumnName.ITEM_COMMON4.ToString, CI.ITEM_COMMON4,
      IdxColumnName.ITEM_COMMON5.ToString, CI.ITEM_COMMON5,
      IdxColumnName.ITEM_COMMON6.ToString, CI.ITEM_COMMON6,
      IdxColumnName.ITEM_COMMON7.ToString, CI.ITEM_COMMON7,
      IdxColumnName.ITEM_COMMON8.ToString, CI.ITEM_COMMON8,
      IdxColumnName.ITEM_COMMON9.ToString, CI.ITEM_COMMON9,
      IdxColumnName.ITEM_COMMON10.ToString, CI.ITEM_COMMON10,
      IdxColumnName.SORT_ITEM_COMMON1.ToString, CI.SORT_ITEM_COMMON1,
      IdxColumnName.SORT_ITEM_COMMON2.ToString, CI.SORT_ITEM_COMMON2,
      IdxColumnName.SORT_ITEM_COMMON3.ToString, CI.SORT_ITEM_COMMON3,
      IdxColumnName.SORT_ITEM_COMMON4.ToString, CI.SORT_ITEM_COMMON4,
      IdxColumnName.SORT_ITEM_COMMON5.ToString, CI.SORT_ITEM_COMMON5,
      IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
      IdxColumnName.SUB_OWNER_NO.ToString, CI.SUB_OWNER_NO,
      IdxColumnName.STORAGE_TYPE.ToString, CI.STORAGE_TYPE,
      IdxColumnName.BND.ToString, CI.BND,
      IdxColumnName.QC_STATUS.ToString, CI.QC_STATUS,
      IdxColumnName.WMS_STOCK_QTY.ToString, CI.WMS_STOCK_QTY,
      IdxColumnName.WMS_UNFINISH_QTY.ToString, CI.WMS_UNFINISH_QTY,
      IdxColumnName.WMS_COMPARSON_QTY.ToString, CI.WMS_COMPARSON_QTY,
      IdxColumnName.ERP_STOCK_QTY.ToString, CI.ERP_STOCK_QTY,
      IdxColumnName.ERP_UNFINISH_QTY.ToString, CI.ERP_UNFINISH_QTY,
      IdxColumnName.ERP_COMPARSON_QTY.ToString, CI.ERP_COMPARSON_QTY,
      IdxColumnName.QUANTITY_VARIANCE.ToString, CI.QUANTITY_VARIANCE,
      IdxColumnName.ERP_SYSTEM.ToString, CI.ERP_SYSTEM,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
      IdxColumnName.ACC_COMMON1.ToString, CI.ACC_COMMON1,
      IdxColumnName.ACC_COMMON2.ToString, CI.ACC_COMMON2,
      IdxColumnName.ACC_COMMON3.ToString, CI.ACC_COMMON3,
      IdxColumnName.ACC_COMMON4.ToString, CI.ACC_COMMON4,
      IdxColumnName.ACC_COMMON5.ToString, CI.ACC_COMMON5,
      IdxColumnName.ACC_COMMON6.ToString, CI.ACC_COMMON6,
      IdxColumnName.ACC_COMMON7.ToString, CI.ACC_COMMON7,
      IdxColumnName.ACC_COMMON8.ToString, CI.ACC_COMMON8,
      IdxColumnName.ACC_COMMON9.ToString, CI.ACC_COMMON9,
      IdxColumnName.ACC_COMMON10.ToString, CI.ACC_COMMON10
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
  Public Shared Function GetEndFlagByDB() As Boolean
    Try
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0} WHERE {1} = '{2}'", TableName, IdxColumnName.ACC_COMMON10, "1")
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsWMS_CT_INVENTORY_COMPARISON = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If Info IsNot Nothing Then Return True
          Next
        End If
        'End If
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function



  Private Shared Function SetInfoFromDB(ByRef Info As clsWMS_CT_INVENTORY_COMPARISON, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
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
        Dim OWNER_NO = "" & RowData.Item(IdxColumnName.OWNER_NO.ToString)
        Dim SUB_OWNER_NO = "" & RowData.Item(IdxColumnName.SUB_OWNER_NO.ToString)
        Dim STORAGE_TYPE = RowData.Item(IdxColumnName.STORAGE_TYPE.ToString)
        Dim BND = RowData.Item(IdxColumnName.BND.ToString)
        Dim QC_STATUS = RowData.Item(IdxColumnName.QC_STATUS.ToString)
        Dim WMS_STOCK_QTY = RowData.Item(IdxColumnName.WMS_STOCK_QTY.ToString)
        Dim WMS_UNFINISH_QTY = RowData.Item(IdxColumnName.WMS_UNFINISH_QTY.ToString)
        Dim WMS_COMPARSON_QTY = RowData.Item(IdxColumnName.WMS_COMPARSON_QTY.ToString)
        Dim ERP_STOCK_QTY = RowData.Item(IdxColumnName.ERP_STOCK_QTY.ToString)
        Dim ERP_UNFINISH_QTY = RowData.Item(IdxColumnName.ERP_UNFINISH_QTY.ToString)
        Dim ERP_COMPARSON_QTY = RowData.Item(IdxColumnName.ERP_COMPARSON_QTY.ToString)
        Dim QUANTITY_VARIANCE = RowData.Item(IdxColumnName.QUANTITY_VARIANCE.ToString)
        Dim ERP_SYSTEM = "" & RowData.Item(IdxColumnName.ERP_SYSTEM.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim ACC_COMMON1 = "" & RowData.Item(IdxColumnName.ACC_COMMON1.ToString)
        Dim ACC_COMMON2 = "" & RowData.Item(IdxColumnName.ACC_COMMON2.ToString)
        Dim ACC_COMMON3 = "" & RowData.Item(IdxColumnName.ACC_COMMON3.ToString)
        Dim ACC_COMMON4 = "" & RowData.Item(IdxColumnName.ACC_COMMON4.ToString)
        Dim ACC_COMMON5 = "" & RowData.Item(IdxColumnName.ACC_COMMON5.ToString)
        Dim ACC_COMMON6 = "" & RowData.Item(IdxColumnName.ACC_COMMON6.ToString)
        Dim ACC_COMMON7 = "" & RowData.Item(IdxColumnName.ACC_COMMON7.ToString)
        Dim ACC_COMMON8 = "" & RowData.Item(IdxColumnName.ACC_COMMON8.ToString)
        Dim ACC_COMMON9 = "" & RowData.Item(IdxColumnName.ACC_COMMON9.ToString)
        Dim ACC_COMMON10 = "" & RowData.Item(IdxColumnName.ACC_COMMON10.ToString)
        Info = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)

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
End Class
