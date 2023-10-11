Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class HOST_CT_TMP_PO_DTLManagement
  Public Shared TableName As String = "HOST_CT_TMP_PO_DTL"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsHOST_CT_TMP_PO_DTL)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    PO_ID
    PO_LINE_NO
    PO_SERIAL_NO
    WO_TYPE
    SKU_NO
    LOT_NO
    QTY
    OWNER_ID
    SUB_OWNER_ID
    COMMON1
    COMMON2
    COMMON3
    COMMON4
    COMMON5
    COMMENTS
    CREATE_TIME
  End Enum

  Public Enum UpdateOption As Integer
    UpdateDic = 0
    UpdateDB = 1
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsHOST_CT_TMP_PO_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32}) values ('{3}','{5}','{7}','{9}','{11}','{13}',{15},'{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}')",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, CI.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, CI.PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, CI.PO_SERIAL_NO,
      IdxColumnName.WO_TYPE.ToString, CI.WO_TYPE,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.LOT_NO.ToString, CI.LOT_NO,
      IdxColumnName.QTY.ToString, CI.QTY,
      IdxColumnName.OWNER_ID.ToString, CI.OWNER_ID,
      IdxColumnName.SUB_OWNER_ID.ToString, CI.SUB_OWNER_ID,
      IdxColumnName.COMMON1.ToString, CI.COMMON1,
      IdxColumnName.COMMON2.ToString, CI.COMMON2,
      IdxColumnName.COMMON3.ToString, CI.COMMON3,
      IdxColumnName.COMMON4.ToString, CI.COMMON4,
      IdxColumnName.COMMON5.ToString, CI.COMMON5,
      IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME
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
  Public Shared Function GetDeleteSQL(ByRef CI As clsHOST_CT_TMP_PO_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {6}='{7}' ",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, CI.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, CI.PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, CI.PO_SERIAL_NO,
      IdxColumnName.WO_TYPE.ToString, CI.WO_TYPE,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.LOT_NO.ToString, CI.LOT_NO,
      IdxColumnName.QTY.ToString, CI.QTY,
      IdxColumnName.OWNER_ID.ToString, CI.OWNER_ID,
      IdxColumnName.SUB_OWNER_ID.ToString, CI.SUB_OWNER_ID,
      IdxColumnName.COMMON1.ToString, CI.COMMON1,
      IdxColumnName.COMMON2.ToString, CI.COMMON2,
      IdxColumnName.COMMON3.ToString, CI.COMMON3,
      IdxColumnName.COMMON4.ToString, CI.COMMON4,
      IdxColumnName.COMMON5.ToString, CI.COMMON5,
      IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME
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
  Public Shared Function GetUpdateSQL(ByRef CI As clsHOST_CT_TMP_PO_DTL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{8}='{9}',{10}='{11}',{12}='{13}',{14}={15},{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}' WHERE {2}='{3}' And {6}='{7}'",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, CI.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, CI.PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, CI.PO_SERIAL_NO,
      IdxColumnName.WO_TYPE.ToString, CI.WO_TYPE,
      IdxColumnName.SKU_NO.ToString, CI.SKU_NO,
      IdxColumnName.LOT_NO.ToString, CI.LOT_NO,
      IdxColumnName.QTY.ToString, CI.QTY,
      IdxColumnName.OWNER_ID.ToString, CI.OWNER_ID,
      IdxColumnName.SUB_OWNER_ID.ToString, CI.SUB_OWNER_ID,
      IdxColumnName.COMMON1.ToString, CI.COMMON1,
      IdxColumnName.COMMON2.ToString, CI.COMMON2,
      IdxColumnName.COMMON3.ToString, CI.COMMON3,
      IdxColumnName.COMMON4.ToString, CI.COMMON4,
      IdxColumnName.COMMON5.ToString, CI.COMMON5,
      IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME
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

  Public Shared Function GetHOST_CT_TMP_PO_DTLDataListByKey_PO_ID_PO_SERIAL_NO(ByVal po_id As String, po_serial_no As String) As List(Of clsHOST_CT_TMP_PO_DTL)
    SyncLock objLock
      Try
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim rs As ADODB.Recordset = Nothing

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4} = '{5}' ",
            strSQL,
            TableName,
            IdxColumnName.PO_ID.ToString, po_id,
            IdxColumnName.PO_SERIAL_NO.ToString, po_serial_no
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, rs)
            Dim DatasetMessage As New DataSet
            Dim OLEDBAdapter As New OleDbDataAdapter
            Dim _lstReturn As New List(Of clsHOST_CT_TMP_PO_DTL)
            OLEDBAdapter.Fill(DatasetMessage, rs, TableName)
            If DatasetMessage.Tables(TableName).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables(TableName).Rows.Count - 1
                Dim Info As clsHOST_CT_TMP_PO_DTL = Nothing
                SetInfoFromDB(Info, DatasetMessage.Tables(TableName).Rows(RowIndex))
                _lstReturn.Add(Info)
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

  Public Shared Function GetWCT_PO_DTLDataListByALL() As List(Of clsHOST_CT_TMP_PO_DTL)
    Try
      Dim _lstReturn As New List(Of clsHOST_CT_TMP_PO_DTL)
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
            Dim Info As clsHOST_CT_TMP_PO_DTL = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info)
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-內部記憶體增刪修
  Private Shared Function AddOrUpdateHOST_CT_TMP_PO_DTLDataToDictionary(ByRef Info As List(Of clsHOST_CT_TMP_PO_DTL)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True
      For Each CI In Info
        Dim _Data As clsHOST_CT_TMP_PO_DTL = CI
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
  Private Shared Function DeleteHOST_CT_TMP_PO_DTLDataToDictionary(ByRef Info As List(Of clsHOST_CT_TMP_PO_DTL)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True
      For i = 0 To Info.Count - 1
        Dim key As String = Info(i).gid
        If dicData.TryRemove(key, Nothing) = False Then
          SendMessageToLog("dicData.TryRemove Failed -HOST_CT_TMP_PO_DTLData", eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateInfo(ByRef Key As String, ByRef Info As clsHOST_CT_TMP_PO_DTL, ByRef objNewTC As clsHOST_CT_TMP_PO_DTL) As clsHOST_CT_TMP_PO_DTL
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsHOST_CT_TMP_PO_DTL, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_LINE_NO = "" & RowData.Item(IdxColumnName.PO_LINE_NO.ToString)
        Dim PO_SERIAL_NO = "" & RowData.Item(IdxColumnName.PO_SERIAL_NO.ToString)
        Dim WO_TYPE = "" & RowData.Item(IdxColumnName.WO_TYPE.ToString)
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim LOT_NO = "" & RowData.Item(IdxColumnName.LOT_NO.ToString)
        Dim QTY = 0 & RowData.Item(IdxColumnName.QTY.ToString)
        Dim OWNER_ID = "" & RowData.Item(IdxColumnName.OWNER_ID.ToString)
        Dim SUB_OWNER_ID = "" & RowData.Item(IdxColumnName.SUB_OWNER_ID.ToString)
        Dim COMMON1 = "" & RowData.Item(IdxColumnName.COMMON1.ToString)
        Dim COMMON2 = "" & RowData.Item(IdxColumnName.COMMON2.ToString)
        Dim COMMON3 = "" & RowData.Item(IdxColumnName.COMMON3.ToString)
        Dim COMMON4 = "" & RowData.Item(IdxColumnName.COMMON4.ToString)
        Dim COMMON5 = "" & RowData.Item(IdxColumnName.COMMON5.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Info = New clsHOST_CT_TMP_PO_DTL(PO_ID, PO_LINE_NO, PO_SERIAL_NO, WO_TYPE, SKU_NO, LOT_NO, QTY, OWNER_ID, SUB_OWNER_ID, COMMON1, COMMON2, COMMON3, COMMON4, COMMON5, COMMENTS, CREATE_TIME)

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
