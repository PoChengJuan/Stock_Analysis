Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Public Class WMS_T_PO_POSTINGManagement
  Public Shared TableName As String = "WMS_T_PO_POSTING"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsPO_POSTING)
  Public Shared Property DictionaryNeeded As Integer = 1  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
  Public Shared LogTool As eCALogTool._ILogTool = Nothing

  Enum IdxColumnName As Integer
    PO_ID
    PO_LINE_NO
    WO_ID
    SORT_ITEM_COMMON1
    SORT_ITEM_COMMON2
    SORT_ITEM_COMMON3
    SORT_ITEM_COMMON4
    SORT_ITEM_COMMON5
    QTY
    UUID
    CREATE_TIME
    UPDATE_TIME
    RESULT
    RESULT_MESSAGE
    H_POP1
    H_POP2
    H_POP3
    H_POP4
    H_POP5
    SKU_NO
    CLOSE_USER_ID
    START_TRANSFER_TIME
    FINISH_TRANSFER_TIME
    ORDER_TYPE
    PO_SERIAL_NO
    TKNUM
    LOT_NO
    OWNER
    SUBOWNER
    KEY_NO
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPO_POSTING) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60}) values ('{3}','{5}','{7}','{9}','{11}',{13},'{15}','{17}','{19}',{21},'{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}')",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
      IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.UUID.ToString, Info.UUID,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.RESULT.ToString, Info.RESULT,
      IdxColumnName.RESULT_MESSAGE.ToString, Info.RESULT_MESSAGE,
      IdxColumnName.H_POP1.ToString, Info.H_POP1,
      IdxColumnName.H_POP2.ToString, Info.H_POP2,
      IdxColumnName.H_POP3.ToString, Info.H_POP3,
      IdxColumnName.H_POP4.ToString, Info.H_POP4,
      IdxColumnName.H_POP5.ToString, Info.H_POP5,
      IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
      IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
      IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.CLOSE_USER_ID.ToString, Info.CLOSE_USER_ID,
      IdxColumnName.START_TRANSFER_TIME.ToString, Info.START_TRANSFER_TIME,
      IdxColumnName.FINISH_TRANSFER_TIME.ToString, Info.FINISH_TRANSFER_TIME,
      IdxColumnName.ORDER_TYPE.ToString, CInt(Info.ORDER_TYPE),
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.TKNUM.ToString, Info.TKNUM,
      IdxColumnName.LOT_NO.ToString, Info.LOT_NO,
      IdxColumnName.OWNER.ToString, Info.OWNER,
      IdxColumnName.SUBOWNER.ToString, Info.SUBOWNER,
      IdxColumnName.KEY_NO.ToString, Info.KEY_NO
)
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsPO_POSTING) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}'",'' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' AND {10}='{11}' AND {12}='{13}' AND {14}='{15}' AND {16}='{17}' AND {18}='{19}' ",
      strSQL,
      TableName,
      IdxColumnName.KEY_NO.ToString, Info.KEY_NO
      )

      'IdxColumnName.PO_ID.ToString, Info.PO_ID,
      'IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      'IdxColumnName.WO_ID.ToString, Info.WO_ID,
      'IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
      'IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
      'IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
      'IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
      'IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
      'IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO

      Dim New_SQL = ""
      ModuleHelpFunc.SQLCorrect(strSQL, New_SQL)

      Return New_SQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsPO_POSTING) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {14}='{15}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{40}='{41}',{42}='{43}',{44}='{45}',{48}='{49}',{50}='{51}',{52}='{53}',{2}='{3}', {4}='{5}', {6}='{7}', {8}='{9}', {10}='{11}', {12}={13}, {34}='{35}', {36}='{37}', {38}='{39}', {46}='{47}' WHERE {54}='{55}' ",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
      IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.UUID.ToString, Info.UUID,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.RESULT.ToString, Info.RESULT,
      IdxColumnName.RESULT_MESSAGE.ToString, Info.RESULT_MESSAGE,
      IdxColumnName.H_POP1.ToString, Info.H_POP1,
      IdxColumnName.H_POP2.ToString, Info.H_POP2,
      IdxColumnName.H_POP3.ToString, Info.H_POP3,
      IdxColumnName.H_POP4.ToString, Info.H_POP4,
      IdxColumnName.H_POP5.ToString, Info.H_POP5,
      IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
      IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
      IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.CLOSE_USER_ID.ToString, Info.CLOSE_USER_ID,
      IdxColumnName.START_TRANSFER_TIME.ToString, Info.START_TRANSFER_TIME,
      IdxColumnName.FINISH_TRANSFER_TIME.ToString, Info.FINISH_TRANSFER_TIME,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.LOT_NO.ToString, Info.LOT_NO,
      IdxColumnName.OWNER.ToString, Info.OWNER,
      IdxColumnName.SUBOWNER.ToString, Info.SUBOWNER,
      IdxColumnName.KEY_NO.ToString, Info.KEY_NO
      )
      Dim New_SQL = ""
      ModuleHelpFunc.SQLCorrect(strSQL, New_SQL)

      Return New_SQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  Public Shared Function GetWMS_T_PO_POSTINGDataListByPO_ID_PO_LINE_NO_WO_ID_SORT_ITEM_COMMON1_5_LOT_NO(ByVal po_id As String, ByVal po_line_no As String, ByVal wo_id As String,
                                                                                                     ByVal sort_item_common1 As String, ByVal sort_item_common2 As String,
                                                                                                     ByVal sort_item_common3 As String, ByVal sort_item_common4 As String,
                                                                                                     ByVal sort_item_common5 As String, ByVal PO_SERIAL_NO As String,
                                                                                                     ByVal LOT_NO As String) As Dictionary(Of String, clsPO_POSTING)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsPO_POSTING)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} ='{3}' AND {4} ='{5}' AND {6} ='{7}' AND {8} ='{9}' AND {10} ='{11}' AND {12} ='{13}' AND {14} ='{15}' AND {16} ='{17}' AND {18} ='{19}'",
            strSQL,
            TableName,
            IdxColumnName.PO_ID.ToString, po_id,
            IdxColumnName.PO_LINE_NO.ToString, po_line_no,
            IdxColumnName.WO_ID.ToString, wo_id,
            IdxColumnName.SORT_ITEM_COMMON1.ToString, sort_item_common1,
            IdxColumnName.SORT_ITEM_COMMON2.ToString, sort_item_common2,
            IdxColumnName.SORT_ITEM_COMMON3.ToString, sort_item_common3,
            IdxColumnName.SORT_ITEM_COMMON4.ToString, sort_item_common4,
            IdxColumnName.SORT_ITEM_COMMON5.ToString, sort_item_common5,
            IdxColumnName.PO_SERIAL_NO.ToString, PO_SERIAL_NO,
            IdxColumnName.LOT_NO.ToString, LOT_NO
            )

            Dim NewSQL As String = ""
            If SQLCorrect(strSQL, NewSQL) Then
            End If
            SendMessageToLog(NewSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(NewSQL, DatasetMessage)
            'Dim DatasetMessage As New DataSet
            'Dim OLEDBAdapter As New OleDbDataAdapter
            'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_POSTING = Nothing
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                  If Info IsNot Nothing Then
                    If _lstReturn.ContainsKey(Info.gid) = False Then
                      _lstReturn.Add(Info.gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
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

  Public Shared Function GetWMS_T_PO_POSTINGDataListByAll() As Dictionary(Of String, clsPO_POSTING)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsPO_POSTING)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1}",
            strSQL,
            TableName
            )

            Dim NewSQL As String = ""
            If SQLCorrect(strSQL, NewSQL) Then
            End If
            SendMessageToLog(NewSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(NewSQL, DatasetMessage)
            'Dim DatasetMessage As New DataSet
            'Dim OLEDBAdapter As New OleDbDataAdapter
            'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_POSTING = Nothing
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                  If Info IsNot Nothing Then
                    If _lstReturn.ContainsKey(Info.gid) = False Then
                      _lstReturn.Add(Info.gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
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
  Public Shared Function GetWMS_T_PO_POSTINGDataListByPO_ID_PO_LINE_NO_WO_ID(ByVal po_id As String, ByVal po_line_no As String, ByVal wo_id As String) As Dictionary(Of String, clsPO_POSTING)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsPO_POSTING)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} ='{3}' AND {4} ='{5}' AND {6} ='{7}' ",
            strSQL,
            TableName,
            IdxColumnName.PO_ID.ToString, po_id,
            IdxColumnName.PO_LINE_NO.ToString, po_line_no,
            IdxColumnName.WO_ID.ToString, wo_id
            )

            Dim NewSQL As String = ""
            If SQLCorrect(strSQL, NewSQL) Then
            End If
            SendMessageToLog(NewSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(NewSQL, DatasetMessage)
            'Dim DatasetMessage As New DataSet
            'Dim OLEDBAdapter As New OleDbDataAdapter
            'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_POSTING = Nothing
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                  If Info IsNot Nothing Then
                    If _lstReturn.ContainsKey(Info.gid) = False Then
                      _lstReturn.Add(Info.gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
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

  Public Shared Function GetWMS_T_PO_POSTINGDataListByWO_ID(ByVal wo_id As String) As Dictionary(Of String, clsPO_POSTING)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsPO_POSTING)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty

            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
            strSQL,
            TableName,
            IdxColumnName.WO_ID.ToString, wo_id
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
            'Dim DatasetMessage As New DataSet
            'Dim OLEDBAdapter As New OleDbDataAdapter
            'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_POSTING = Nothing
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                  If Info IsNot Nothing Then
                    If _lstReturn.ContainsKey(Info.gid) = False Then
                      _lstReturn.Add(Info.gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
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
  Public Shared Function GetWMS_T_PO_POSTINGDataListBydicWO_ID(ByVal dicWO_ID As Dictionary(Of String, String)) As Dictionary(Of String, clsPO_POSTING)
    SyncLock objLock
      Try
        Dim ret_dic As New Dictionary(Of String, clsPO_POSTING)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim strIDList As String = ""
            Dim strWhere As String = ""
            Dim DatasetMessage As New DataSet

            Dim count_flag = 0
            For i = 0 To dicWO_ID.Count - 1
              If strIDList = "" Then
                strIDList = "'" & dicWO_ID.Keys(i) & "'"
              Else
                strIDList = strIDList & ",'" & dicWO_ID.Keys(i) & "'"
              End If
              If i - count_flag > 800 OrElse i = (dicWO_ID.Count - 1) Then
                count_flag = i
                strWhere = ""
                If strWhere = "" Then
                  strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.WO_ID.ToString, strIDList)
                Else
                  strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.WO_ID.ToString, strIDList)
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
                    Dim Info As clsPO_POSTING = Nothing
                    SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                    If ret_dic.ContainsKey(Info.gid) = False Then
                      ret_dic.Add(Info.gid, Info)
                    End If
                  Next
                End If
                strIDList = ""
              End If
            Next
          End If
        End If
        Return ret_dic
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function

  '-Function



  '-內部記憶體增刪修

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsPO_POSTING, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_LINE_NO = "" & RowData.Item(IdxColumnName.PO_LINE_NO.ToString)
        Dim PO_SERIAL_NO = "" & RowData.Item(IdxColumnName.PO_SERIAL_NO.ToString)
        Dim WO_ID = "" & RowData.Item(IdxColumnName.WO_ID.ToString)
        Dim SORT_ITEM_COMMON1 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON1.ToString)
        Dim SORT_ITEM_COMMON2 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON2.ToString)
        Dim SORT_ITEM_COMMON3 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON3.ToString)
        Dim SORT_ITEM_COMMON4 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON4.ToString)
        Dim SORT_ITEM_COMMON5 = "" & RowData.Item(IdxColumnName.SORT_ITEM_COMMON5.ToString)
        Dim QTY = 0 & RowData.Item(IdxColumnName.QTY.ToString)
        Dim UUID = "" & RowData.Item(IdxColumnName.UUID.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim RESULT = IIf(IsNumeric(RowData.Item(IdxColumnName.RESULT.ToString)), RowData.Item(IdxColumnName.RESULT.ToString), 0)
        Dim RESULT_MESSAGE = "" & RowData.Item(IdxColumnName.RESULT_MESSAGE.ToString)
        Dim H_POP1 = "" & RowData.Item(IdxColumnName.H_POP1.ToString)
        Dim H_POP2 = "" & RowData.Item(IdxColumnName.H_POP2.ToString)
        Dim H_POP3 = "" & RowData.Item(IdxColumnName.H_POP3.ToString)
        Dim H_POP4 = "" & RowData.Item(IdxColumnName.H_POP4.ToString)
        Dim H_POP5 = "" & RowData.Item(IdxColumnName.H_POP5.ToString)
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim CLOSE_USER_ID = "" & RowData.Item(IdxColumnName.CLOSE_USER_ID.ToString)
        Dim START_TRANSFER_TIME = "" & RowData.Item(IdxColumnName.START_TRANSFER_TIME.ToString)
        Dim FINISH_TRANSFER_TIME = "" & RowData.Item(IdxColumnName.FINISH_TRANSFER_TIME.ToString)
        Dim ORDER_TYPE = IIf(IsNumeric(RowData.Item(IdxColumnName.ORDER_TYPE.ToString)), RowData.Item(IdxColumnName.ORDER_TYPE.ToString), 0)
        Dim TKNUM = "" & RowData.Item(IdxColumnName.TKNUM.ToString)
        Dim LOT_NO = "" & RowData.Item(IdxColumnName.LOT_NO.ToString)
        Dim OWNER = "" & RowData.Item(IdxColumnName.OWNER.ToString)
        Dim SUBOWNER = "" & RowData.Item(IdxColumnName.SUBOWNER.ToString)
        Dim KEY_NO = "" & RowData.Item(IdxColumnName.KEY_NO.ToString)
        Info = New clsPO_POSTING(PO_ID, PO_LINE_NO, WO_ID, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, QTY, UUID, CREATE_TIME, UPDATE_TIME, RESULT, RESULT_MESSAGE, H_POP1, H_POP2, H_POP3, H_POP4, H_POP5, SKU_NO, CLOSE_USER_ID, START_TRANSFER_TIME, FINISH_TRANSFER_TIME, ORDER_TYPE, PO_SERIAL_NO, TKNUM, LOT_NO, OWNER, SUBOWNER, KEY_NO)

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
