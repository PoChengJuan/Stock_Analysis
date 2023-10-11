Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_CT_PRODUCTION_REPORTManagement
  Public Shared TableName As String = "WMS_CT_PRODUCTION_REPORT"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsWMS_CT_PRODUCTION_REPORT)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    PO_ID
    SKU_NO
    REPORT_STATUS
    QTY
    QTY_NG
    REPORT_QTY
    REPORT_QTY_NG
    TB003
    TB004
    TB005
    TB008
    TB007
    TB010
    TC003
    TC004
    TC005
    TC006
    TC007
    TC008
    TC009
    TC010
    TC014
    TC016
    TC020
    TC021
    TC200
    TC201
    CREATE_TIME
    UPDATE_TIME
    FINISH_TIME
  End Enum

  Public Enum UpdateOption As Integer
    UpdateDic = 0
    UpdateDB = 1
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsWMS_CT_PRODUCTION_REPORT) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64}) values ('{3}','{5}','{7}','{9}',{11},{13},{15},{17},{19},'{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}',{49},{51},'{53}','{55}','{57}','{59}','{61}','{63}','{65}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
      IdxColumnName.PO_ID.ToString, CI.PO_ID,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.REPORT_STATUS.ToString, CI.REPORT_STATUS,
      IdxColumnName.QTY.ToString, CI.QTY,
      IdxColumnName.QTY_NG.ToString, CI.QTY_NG,
      IdxColumnName.REPORT_QTY.ToString, CI.REPORT_QTY,
      IdxColumnName.REPORT_QTY_NG.ToString, CI.REPORT_QTY_NG,
      IdxColumnName.TB003.ToString, CI.TB003,
      IdxColumnName.TB004.ToString, CI.TB004,
      IdxColumnName.TB005.ToString, CI.TB005,
      IdxColumnName.TB008.ToString, CI.TB008,
      IdxColumnName.TB007.ToString, CI.TB007,
      IdxColumnName.TB010.ToString, CI.TB010,
      IdxColumnName.TC003.ToString, CI.TC003,
      IdxColumnName.TC004.ToString, CI.TC004,
      IdxColumnName.TC005.ToString, CI.TC005,
      IdxColumnName.TC006.ToString, CI.TC006,
      IdxColumnName.TC007.ToString, CI.TC007,
      IdxColumnName.TC008.ToString, CI.TC008,
      IdxColumnName.TC009.ToString, CI.TC009,
      IdxColumnName.TC010.ToString, CI.TC010,
      IdxColumnName.TC014.ToString, CI.TC014,
      IdxColumnName.TC016.ToString, CI.TC016,
      IdxColumnName.TC020.ToString, CI.TC020,
      IdxColumnName.TC021.ToString, CI.TC021,
      IdxColumnName.TC200.ToString, CI.TC200,
      IdxColumnName.TC201.ToString, CI.TC201,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME,
      IdxColumnName.FINISH_TIME.ToString, CI.FINISH_TIME
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
  Public Shared Function GetDeleteSQL(ByRef CI As clsWMS_CT_PRODUCTION_REPORT) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
      IdxColumnName.PO_ID.ToString, CI.PO_ID,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.REPORT_STATUS.ToString, CI.REPORT_STATUS,
      IdxColumnName.QTY.ToString, CI.QTY,
      IdxColumnName.QTY_NG.ToString, CI.QTY_NG,
      IdxColumnName.REPORT_QTY.ToString, CI.REPORT_QTY,
      IdxColumnName.REPORT_QTY_NG.ToString, CI.REPORT_QTY_NG,
      IdxColumnName.TB003.ToString, CI.TB003,
      IdxColumnName.TB004.ToString, CI.TB004,
      IdxColumnName.TB005.ToString, CI.TB005,
      IdxColumnName.TB008.ToString, CI.TB008,
      IdxColumnName.TB007.ToString, CI.TB007,
      IdxColumnName.TB010.ToString, CI.TB010,
      IdxColumnName.TC003.ToString, CI.TC003,
      IdxColumnName.TC004.ToString, CI.TC004,
      IdxColumnName.TC005.ToString, CI.TC005,
      IdxColumnName.TC006.ToString, CI.TC006,
      IdxColumnName.TC007.ToString, CI.TC007,
      IdxColumnName.TC008.ToString, CI.TC008,
      IdxColumnName.TC009.ToString, CI.TC009,
      IdxColumnName.TC010.ToString, CI.TC010,
      IdxColumnName.TC014.ToString, CI.TC014,
      IdxColumnName.TC016.ToString, CI.TC016,
      IdxColumnName.TC020.ToString, CI.TC020,
      IdxColumnName.TC021.ToString, CI.TC021,
      IdxColumnName.TC200.ToString, CI.TC200,
      IdxColumnName.TC201.ToString, CI.TC201,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME,
      IdxColumnName.FINISH_TIME.ToString, CI.FINISH_TIME
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
  Public Shared Function GetUpdateSQL(ByRef CI As clsWMS_CT_PRODUCTION_REPORT) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}={11},{12}={13},{14}={15},{16}={17},{18}={19},{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}={49},{50}={51},{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
      IdxColumnName.PO_ID.ToString, CI.PO_ID,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.REPORT_STATUS.ToString, CI.REPORT_STATUS,
      IdxColumnName.QTY.ToString, CI.QTY,
      IdxColumnName.QTY_NG.ToString, CI.QTY_NG,
      IdxColumnName.REPORT_QTY.ToString, CI.REPORT_QTY,
      IdxColumnName.REPORT_QTY_NG.ToString, CI.REPORT_QTY_NG,
      IdxColumnName.TB003.ToString, CI.TB003,
      IdxColumnName.TB004.ToString, CI.TB004,
      IdxColumnName.TB005.ToString, CI.TB005,
      IdxColumnName.TB008.ToString, CI.TB008,
      IdxColumnName.TB007.ToString, CI.TB007,
      IdxColumnName.TB010.ToString, CI.TB010,
      IdxColumnName.TC003.ToString, CI.TC003,
      IdxColumnName.TC004.ToString, CI.TC004,
      IdxColumnName.TC005.ToString, CI.TC005,
      IdxColumnName.TC006.ToString, CI.TC006,
      IdxColumnName.TC007.ToString, CI.TC007,
      IdxColumnName.TC008.ToString, CI.TC008,
      IdxColumnName.TC009.ToString, CI.TC009,
      IdxColumnName.TC010.ToString, CI.TC010,
      IdxColumnName.TC014.ToString, CI.TC014,
      IdxColumnName.TC016.ToString, CI.TC016,
      IdxColumnName.TC020.ToString, CI.TC020,
      IdxColumnName.TC021.ToString, CI.TC021,
      IdxColumnName.TC200.ToString, CI.TC200,
      IdxColumnName.TC201.ToString, CI.TC201,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME,
      IdxColumnName.FINISH_TIME.ToString, CI.FINISH_TIME
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
  Public Shared Function AddWMS_CT_PRODUCTION_REPORTData(ByVal Info As clsWMS_CT_PRODUCTION_REPORT, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If AddlstWMS_CT_PRODUCTION_REPORTData(New List(Of clsWMS_CT_PRODUCTION_REPORT)({Info}), SendToDB) = True Then
          Return True
        End If '-載不載入記憶體都是呼叫同一個function
        Return False
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function AddlstWMS_CT_PRODUCTION_REPORTData(ByVal Info As List(Of clsWMS_CT_PRODUCTION_REPORT), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True
        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).gid
            If dicData.ContainsKey(key) = True Then
              SendMessageToLog("Add the same key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next
          If SendToDB Then
            If InsertWMS_CT_PRODUCTION_REPORTDataToDB(Info) Then
              If AddOrUpdateWMS_CT_PRODUCTION_REPORTDataToDictionary(Info) Then
                SendMessageToLog("InsertDic WMS_CT_PRODUCTION_REPORTData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("InsertDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("InsertDB WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If AddOrUpdateWMS_CT_PRODUCTION_REPORTDataToDictionary(Info) Then
              SendMessageToLog("InsertDic WMS_CT_PRODUCTION_REPORTData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("InsertDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If InsertWMS_CT_PRODUCTION_REPORTDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("InsertDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function UpdateWMS_CT_PRODUCTION_REPORTData(ByVal Info As clsWMS_CT_PRODUCTION_REPORT, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If UpdatelstWMS_CT_PRODUCTION_REPORTData(New List(Of clsWMS_CT_PRODUCTION_REPORT)({Info}), SendToDB) = True Then
          Return True
        End If
        Return False
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function UpdatelstWMS_CT_PRODUCTION_REPORTData(ByVal Info As List(Of clsWMS_CT_PRODUCTION_REPORT), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True
        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).gid
            If dicData.ContainsKey(key) = True Then
              SendMessageToLog("There is no key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next
          If SendToDB Then
            If UpdateWMS_CT_PRODUCTION_REPORTDataToDB(Info) Then
              If AddOrUpdateWMS_CT_PRODUCTION_REPORTDataToDictionary(Info) Then
                SendMessageToLog("UpdateDic WMS_CT_PRODUCTION_REPORTData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("UpdateDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("UpdateDB WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If AddOrUpdateWMS_CT_PRODUCTION_REPORTDataToDictionary(Info) Then
              SendMessageToLog("UpdateDic WMS_CT_PRODUCTION_REPORTData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("UpdateDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If UpdateWMS_CT_PRODUCTION_REPORTDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("UpdateDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function DeleteWMS_CT_PRODUCTION_REPORTData(ByVal Info As clsWMS_CT_PRODUCTION_REPORT, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If DeletelstWMS_CT_PRODUCTION_REPORTData(New List(Of clsWMS_CT_PRODUCTION_REPORT)({Info}), SendToDB) = True Then
          Return True
        End If '-載不載入記憶體都是呼叫同一個function
        Return False
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function DeletelstWMS_CT_PRODUCTION_REPORTData(ByVal Info As List(Of clsWMS_CT_PRODUCTION_REPORT), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True
        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).gid
            If dicData.ContainsKey(key) = True Then
              SendMessageToLog("There is no key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next
          If SendToDB Then
            If DeleteWMS_CT_PRODUCTION_REPORTDataToDB(Info) Then
              If DeleteWMS_CT_PRODUCTION_REPORTDataToDictionary(Info) Then
                SendMessageToLog("DeleteDic WMS_CT_PRODUCTION_REPORTData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("DeleteDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("DeleteDB WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If DeleteWMS_CT_PRODUCTION_REPORTDataToDictionary(Info) Then
              SendMessageToLog("DeleteDic WMS_CT_PRODUCTION_REPORTData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("DeleteDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If DeleteWMS_CT_PRODUCTION_REPORTDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("DeleteDic WMS_CT_PRODUCTION_REPORTData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function GetWMS_CT_PRODUCTION_REPORTDataListByKey_FACTORY_NO_AREA_NO_PO_ID_SKU_NO(ByVal factory_no As String, area_no As String, po_id As String, sku_no As String) As Dictionary(Of String, clsWMS_CT_PRODUCTION_REPORT)
    SyncLock objLock
      Try
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim rs As ADODB.Recordset = Nothing

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4} = '{5}' AND {6} = '{7}' AND {8} = '{9}' ",
            strSQL,
            TableName,
            IdxColumnName.FACTORY_NO.ToString, factory_no,
            IdxColumnName.AREA_NO.ToString, area_no,
            IdxColumnName.PO_ID.ToString, po_id,
            IdxColumnName.SKU_NO.ToString, sku_no
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, rs)
            Dim DatasetMessage As New DataSet
            Dim OLEDBAdapter As New OleDbDataAdapter
            Dim _lstReturn As New Dictionary(Of String, clsWMS_CT_PRODUCTION_REPORT)
            OLEDBAdapter.Fill(DatasetMessage, rs, TableName)
            If DatasetMessage.Tables(TableName).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables(TableName).Rows.Count - 1
                Dim Info As clsWMS_CT_PRODUCTION_REPORT = Nothing
                SetInfoFromDB(Info, DatasetMessage.Tables(TableName).Rows(RowIndex))
                _lstReturn.Add(Info.gid, Info)
              Next
            End If
            Return _lstReturn
          End If
        End If
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function
  Private Shared Function InsertWMS_CT_PRODUCTION_REPORTDataToDB(ByRef Info As List(Of clsWMS_CT_PRODUCTION_REPORT)) As Integer
    Try
      If Info Is Nothing Then Return -1
      If Info.Count = 0 Then Return 0

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64}) values ('{3}','{5}','{7}','{9}',{11},{13},{15},{17},{19},'{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}',{49},{51},'{53}','{55}','{57}','{59}','{61}','{63}','{65}')",
        strSQL,
        TableName,
        IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
        IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
        IdxColumnName.PO_ID.ToString, CI.PO_ID,
        IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
        IdxColumnName.REPORT_STATUS.ToString, CI.REPORT_STATUS,
        IdxColumnName.QTY.ToString, CI.QTY,
        IdxColumnName.QTY_NG.ToString, CI.QTY_NG,
        IdxColumnName.REPORT_QTY.ToString, CI.REPORT_QTY,
        IdxColumnName.REPORT_QTY_NG.ToString, CI.REPORT_QTY_NG,
        IdxColumnName.TB003.ToString, CI.TB003,
        IdxColumnName.TB004.ToString, CI.TB004,
        IdxColumnName.TB005.ToString, CI.TB005,
        IdxColumnName.TB008.ToString, CI.TB008,
        IdxColumnName.TB007.ToString, CI.TB007,
        IdxColumnName.TB010.ToString, CI.TB010,
        IdxColumnName.TC003.ToString, CI.TC003,
        IdxColumnName.TC004.ToString, CI.TC004,
        IdxColumnName.TC005.ToString, CI.TC005,
        IdxColumnName.TC006.ToString, CI.TC006,
        IdxColumnName.TC007.ToString, CI.TC007,
        IdxColumnName.TC008.ToString, CI.TC008,
        IdxColumnName.TC009.ToString, CI.TC009,
        IdxColumnName.TC010.ToString, CI.TC010,
        IdxColumnName.TC014.ToString, CI.TC014,
        IdxColumnName.TC016.ToString, CI.TC016,
        IdxColumnName.TC020.ToString, CI.TC020,
        IdxColumnName.TC021.ToString, CI.TC021,
        IdxColumnName.TC200.ToString, CI.TC200,
        IdxColumnName.TC201.ToString, CI.TC201,
        IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
        IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME,
        IdxColumnName.FINISH_TIME.ToString, CI.FINISH_TIME
        )
        lstSql.Add(strSQL)
      Next

      Dim NewSQL As New List(Of String)
      If SQLCorrect(lstSql, NewSQL) = False Then
        Return Nothing
      End If
      If SendSQLToDB(NewSQL) = True Then
        Return True
      Else
        SendMessageToLog("Insert to WMS_CT_PRODUCTION_REPORT DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateWMS_CT_PRODUCTION_REPORTDataToDB(ByRef Info As List(Of clsWMS_CT_PRODUCTION_REPORT)) As Integer
    Try
      If Info Is Nothing Then Return -1
      If Info.Count = 0 Then Return 0

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Update {1} SET {10}={11},{12}={13},{14}={15},{16}={17},{18}={19},{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}={49},{50}={51},{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
        strSQL,
        TableName,
        IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
        IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
        IdxColumnName.PO_ID.ToString, CI.PO_ID,
        IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
        IdxColumnName.REPORT_STATUS.ToString, CI.REPORT_STATUS,
        IdxColumnName.QTY.ToString, CI.QTY,
        IdxColumnName.QTY_NG.ToString, CI.QTY_NG,
        IdxColumnName.REPORT_QTY.ToString, CI.REPORT_QTY,
        IdxColumnName.REPORT_QTY_NG.ToString, CI.REPORT_QTY_NG,
        IdxColumnName.TB003.ToString, CI.TB003,
        IdxColumnName.TB004.ToString, CI.TB004,
        IdxColumnName.TB005.ToString, CI.TB005,
        IdxColumnName.TB008.ToString, CI.TB008,
        IdxColumnName.TB007.ToString, CI.TB007,
        IdxColumnName.TB010.ToString, CI.TB010,
        IdxColumnName.TC003.ToString, CI.TC003,
        IdxColumnName.TC004.ToString, CI.TC004,
        IdxColumnName.TC005.ToString, CI.TC005,
        IdxColumnName.TC006.ToString, CI.TC006,
        IdxColumnName.TC007.ToString, CI.TC007,
        IdxColumnName.TC008.ToString, CI.TC008,
        IdxColumnName.TC009.ToString, CI.TC009,
        IdxColumnName.TC010.ToString, CI.TC010,
        IdxColumnName.TC014.ToString, CI.TC014,
        IdxColumnName.TC016.ToString, CI.TC016,
        IdxColumnName.TC020.ToString, CI.TC020,
        IdxColumnName.TC021.ToString, CI.TC021,
        IdxColumnName.TC200.ToString, CI.TC200,
        IdxColumnName.TC201.ToString, CI.TC201,
        IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
        IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME,
        IdxColumnName.FINISH_TIME.ToString, CI.FINISH_TIME
        )
        lstSql.Add(strSQL)
      Next

      Dim NewSQL As New List(Of String)
      If SQLCorrect(lstSql, NewSQL) = False Then
        Return Nothing
      End If
      If SendSQLToDB(NewSQL) = True Then
        Return True
      Else
        SendMessageToLog("Update to WMS_CT_PRODUCTION_REPORT DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function DeleteWMS_CT_PRODUCTION_REPORTDataToDB(ByRef Info As List(Of clsWMS_CT_PRODUCTION_REPORT)) As Integer
    Try
      If Info Is Nothing Then Return -1
      If Info.Count = 0 Then Return 0

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
        strSQL,
        TableName,
        IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
        IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
        IdxColumnName.PO_ID.ToString, CI.PO_ID,
        IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
        IdxColumnName.REPORT_STATUS.ToString, CI.REPORT_STATUS,
        IdxColumnName.QTY.ToString, CI.QTY,
        IdxColumnName.QTY_NG.ToString, CI.QTY_NG,
        IdxColumnName.REPORT_QTY.ToString, CI.REPORT_QTY,
        IdxColumnName.REPORT_QTY_NG.ToString, CI.REPORT_QTY_NG,
        IdxColumnName.TB003.ToString, CI.TB003,
        IdxColumnName.TB004.ToString, CI.TB004,
        IdxColumnName.TB005.ToString, CI.TB005,
        IdxColumnName.TB008.ToString, CI.TB008,
        IdxColumnName.TB007.ToString, CI.TB007,
        IdxColumnName.TB010.ToString, CI.TB010,
        IdxColumnName.TC003.ToString, CI.TC003,
        IdxColumnName.TC004.ToString, CI.TC004,
        IdxColumnName.TC005.ToString, CI.TC005,
        IdxColumnName.TC006.ToString, CI.TC006,
        IdxColumnName.TC007.ToString, CI.TC007,
        IdxColumnName.TC008.ToString, CI.TC008,
        IdxColumnName.TC009.ToString, CI.TC009,
        IdxColumnName.TC010.ToString, CI.TC010,
        IdxColumnName.TC014.ToString, CI.TC014,
        IdxColumnName.TC016.ToString, CI.TC016,
        IdxColumnName.TC020.ToString, CI.TC020,
        IdxColumnName.TC021.ToString, CI.TC021,
        IdxColumnName.TC200.ToString, CI.TC200,
        IdxColumnName.TC201.ToString, CI.TC201,
        IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
        IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME,
        IdxColumnName.FINISH_TIME.ToString, CI.FINISH_TIME
        )
        lstSql.Add(strSQL)
      Next

      Dim NewSQL As New List(Of String)
      If SQLCorrect(lstSql, NewSQL) = False Then
        Return Nothing
      End If
      If SendSQLToDB(NewSQL) = True Then
        Return True
      Else
        SendMessageToLog("Delete to WMS_CT_PRODUCTION_REPORT DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '-內部記憶體增刪修
  Private Shared Function AddOrUpdateWMS_CT_PRODUCTION_REPORTDataToDictionary(ByRef Info As List(Of clsWMS_CT_PRODUCTION_REPORT)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True
      For Each CI In Info
        Dim _Data As clsWMS_CT_PRODUCTION_REPORT = CI
        Dim key As String = _Data.gid
        dicData.AddOrUpdate(key,
        _Data,
        Function(dicKey, ExistVal)
          UpdateInfo(dicKey, ExistVal, _Data)
          Return ExistVal
        End Function)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function DeleteWMS_CT_PRODUCTION_REPORTDataToDictionary(ByRef Info As List(Of clsWMS_CT_PRODUCTION_REPORT)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True
      For i = 0 To Info.Count - 1
        Dim key As String = Info(i).gid
        If dicData.TryRemove(key, Nothing) = False Then
          SendMessageToLog("dicData.TryRemove Failed -WMS_CT_PRODUCTION_REPORTData", eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateInfo(ByRef Key As String, ByRef Info As clsWMS_CT_PRODUCTION_REPORT, ByRef objNewTC As clsWMS_CT_PRODUCTION_REPORT) As clsWMS_CT_PRODUCTION_REPORT
    Try
      If Key = Info.gid Then
        Info.Update_To_Memory(objNewTC)
      Else
        SendMessageToLog("Dictionary has the different key", eCALogTool.ILogTool.enuTrcLevel.lvError)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
    Return Info
  End Function
  Private Shared Function SetInfoFromDB(ByRef Info As clsWMS_CT_PRODUCTION_REPORT, ByRef RowData As DataRow) As Boolean
    Try
            If RowData IsNot Nothing Then
                Dim FACTORY_NO = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
                Dim AREA_NO = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
                Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
                Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
                Dim REPORT_STATUS = 0 & RowData.Item(IdxColumnName.REPORT_STATUS.ToString)
                Dim QTY = 0 & RowData.Item(IdxColumnName.QTY.ToString)
                Dim QTY_NG = 0 & RowData.Item(IdxColumnName.QTY_NG.ToString)
                Dim REPORT_QTY = 0 & RowData.Item(IdxColumnName.REPORT_QTY.ToString)
                Dim REPORT_QTY_NG = 0 & RowData.Item(IdxColumnName.REPORT_QTY_NG.ToString)
                Dim TB003 = "" & RowData.Item(IdxColumnName.TB003.ToString)
                Dim TB004 = "" & RowData.Item(IdxColumnName.TB004.ToString)
                Dim TB005 = "" & RowData.Item(IdxColumnName.TB005.ToString)
                Dim TB008 = "" & RowData.Item(IdxColumnName.TB008.ToString)
                Dim TB007 = "" & RowData.Item(IdxColumnName.TB007.ToString)
                Dim TB010 = "" & RowData.Item(IdxColumnName.TB010.ToString)
                Dim TC003 = "" & RowData.Item(IdxColumnName.TC003.ToString)
                Dim TC004 = "" & RowData.Item(IdxColumnName.TC004.ToString)
                Dim TC005 = "" & RowData.Item(IdxColumnName.TC005.ToString)
                Dim TC006 = "" & RowData.Item(IdxColumnName.TC006.ToString)
                Dim TC007 = "" & RowData.Item(IdxColumnName.TC007.ToString)
                Dim TC008 = "" & RowData.Item(IdxColumnName.TC008.ToString)
                Dim TC009 = "" & RowData.Item(IdxColumnName.TC009.ToString)
                Dim TC010 = "" & RowData.Item(IdxColumnName.TC010.ToString)
                Dim TC014 = 0 & RowData.Item(IdxColumnName.TC014.ToString)
                Dim TC016 = 0 & RowData.Item(IdxColumnName.TC016.ToString)
                Dim TC020 = "" & RowData.Item(IdxColumnName.TC020.ToString)
                Dim TC021 = "" & RowData.Item(IdxColumnName.TC021.ToString)
                Dim TC200 = "" & RowData.Item(IdxColumnName.TC200.ToString)
                Dim TC201 = "" & RowData.Item(IdxColumnName.TC201.ToString)
                Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
                Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
                Dim FINISH_TIME = "" & RowData.Item(IdxColumnName.FINISH_TIME.ToString)
                Info = New clsWMS_CT_PRODUCTION_REPORT(FACTORY_NO, AREA_NO, PO_ID, SKU_NO, REPORT_STATUS, QTY, QTY_NG, REPORT_QTY, REPORT_QTY_NG, TB003, TB004, TB005, TB008, TB007, TB010, TC003, TC004, TC005, TC006, TC007, TC008, TC009, TC010, TC014, TC016, TC020, TC021, TC200, TC201, CREATE_TIME, UPDATE_TIME, FINISH_TIME)

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
