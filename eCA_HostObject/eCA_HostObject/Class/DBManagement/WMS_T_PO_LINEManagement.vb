Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Public Class WMS_T_PO_LINEManagement
  Public Shared TableName As String = "WMS_T_PO_LINE"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsPO_LINE)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
  Public Shared LogTool As eCALogTool._ILogTool = Nothing

  Enum IdxColumnName As Integer
    PO_ID
    PO_LINE_NO
    QTY
    QTY_FINISH
    H_QTY_PROCESS
    H_POL1
    H_POL2
    H_POL3
    H_POL4
    H_POL5
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPO_LINE) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20}) values ('{3}','{5}',{7},{9},{11},'{13}','{15}','{17}','{19}','{21}')",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.H_QTY_PROCESS.ToString, Info.H_QTY_PROCESS,
                             IdxColumnName.H_POL1.ToString, Info.H_POL1,
                             IdxColumnName.H_POL2.ToString, Info.H_POL2,
                             IdxColumnName.H_POL3.ToString, Info.H_POL3,
                             IdxColumnName.H_POL4.ToString, Info.H_POL4,
                             IdxColumnName.H_POL5.ToString, Info.H_POL5
     )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsPO_LINE) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.H_QTY_PROCESS.ToString, Info.H_QTY_PROCESS
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsPO_LINE) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}={7},{8}={9},{10}={11},{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}' WHERE {2}='{3}' And {4}='{5}'",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.H_QTY_PROCESS.ToString, Info.H_QTY_PROCESS,
                             IdxColumnName.H_POL1.ToString, Info.H_POL1,
                             IdxColumnName.H_POL2.ToString, Info.H_POL2,
                             IdxColumnName.H_POL3.ToString, Info.H_POL3,
                             IdxColumnName.H_POL4.ToString, Info.H_POL4,
                             IdxColumnName.H_POL5.ToString, Info.H_POL5
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '- Add & Insert
  'Public Shared Function AddclsWMS_T_PO_LINE(ByVal Info As clsPO_LINE) As Boolean
  '  SyncLock objLock
  '    Try
  '      If Info Is Nothing Then Return False
  '      If AddlstclsWMS_T_PO_LINE(New List(Of clsPO_LINE)({Info})) = True Then
  '        Return True
  '      End If '-載不載入記憶體都是呼叫同一個function
  '      Return False
  '    Catch ex As Exception
  '      Return False
  '    End Try
  '  End SyncLock
  'End Function
  'Public Shared Function AddlstclsWMS_T_PO_LINE(ByVal Info As List(Of clsPO_LINE)) As Boolean
  '  SyncLock objLock
  '    Try
  '      Dim AddWork As New List(Of clsPO_LINE)
  '      If Info Is Nothing Then Return False
  '      If Info.Count = 0 Then Return True

  '      If InsertclsWMS_T_PO_LINEToDB(Info) = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("InsertDB clsWMS_T_PO_LINE Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If
  '    Catch ex As Exception
  '      Return False
  '    End Try
  '  End SyncLock
  'End Function

  ''- Update
  'Public Shared Function UpdateclsWMS_T_PO_LINE(ByVal Info As clsPO_LINE) As Boolean
  '  SyncLock objLock
  '    Try
  '      If Info Is Nothing Then Return False
  '      If UpdatelstclsWMS_T_PO_LINE(New List(Of clsPO_LINE)({Info})) = True Then
  '        Return True
  '      End If
  '      Return False
  '    Catch ex As Exception
  '      Return False
  '    End Try
  '  End SyncLock
  'End Function
  'Public Shared Function UpdatelstclsWMS_T_PO_LINE(ByVal Info As List(Of clsPO_LINE)) As Boolean
  '  SyncLock objLock
  '    Try
  '      Dim UpdateWork As New List(Of clsPO_LINE)
  '      If Info Is Nothing Then Return -1
  '      If Info.Count = 0 Then Return 0
  '      If UpdateclsWMS_T_PO_LINEToDB(Info) = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("UpdateDB clsWMS_T_PO_LINE Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If
  '    Catch ex As Exception
  '      Return False
  '    End Try
  '  End SyncLock
  'End Function

  ''- Delete
  'Public Shared Function DeleteclsWMS_T_PO_LINE(ByVal Info As clsPO_LINE) As Boolean
  '  SyncLock objLock
  '    Try
  '      If Info Is Nothing Then Return False
  '      If DeletelstclsWMS_T_PO_LINE(New List(Of clsPO_LINE)({Info})) = True Then
  '        Return True
  '      End If '-載不載入記憶體都是呼叫同一個function
  '      Return False
  '    Catch ex As Exception
  '      Return False
  '    End Try
  '  End SyncLock
  'End Function
  'Public Shared Function DeletelstclsWMS_T_PO_LINE(ByVal Info As List(Of clsPO_LINE)) As Boolean
  '  SyncLock objLock
  '    Try
  '      Dim DeleteWork As New List(Of clsPO_LINE)
  '      If Info Is Nothing Then Return False
  '      If Info.Count = 0 Then Return True

  '      If DeleteclsWMS_T_PO_LINEToDB(Info) = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("DeleteDB clsWMS_T_PO_LINE Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If
  '    Catch ex As Exception
  '      Return False
  '    End Try
  '  End SyncLock
  'End Function

  '- GET
  Public Shared Function GetdicWMS_T_PO_LINEListByALL() As Dictionary(Of String, clsPO_LINE)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsPO_LINE)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {0}", TableName)
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_LINE
                SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                If _lstReturn.ContainsKey(Info.gid) = False Then
                  _lstReturn.Add(Info.gid, Info)
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
  Public Shared Function GetclsWMS_T_PO_LINEListByPO_ID_PO_LINE_NO(ByVal po_id As String, po_line_no As String) As List(Of clsPO_LINE)
    SyncLock objLock
      Try
        Dim _lstReturn As New List(Of clsPO_LINE)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4} = '{5}' ",
            strSQL,
            TableName,
            IdxColumnName.PO_ID.ToString, po_id,
            IdxColumnName.PO_LINE_NO.ToString, po_line_no
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsPO_LINE
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

  '-Function


  '-以下為內部私人用
  'Private Shared Function InsertclsWMS_T_PO_LINEToDB(ByRef Info As List(Of clsPO_LINE)) As Boolean
  '  Try
  '    If Info Is Nothing Then Return False
  '    If Info.Count = 0 Then Return True

  '    Dim strSQL As String = ""
  '    Dim rs As ADODB.Recordset = Nothing
  '    Dim lstSql As New List(Of String)
  '    For Each CI In Info
  '      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10}) values ('{3}','{5}',{7},{9},{11})",
  '      strSQL,
  '      TableName,
  '      IdxColumnName.PO_ID.ToString, CI.PO_ID,
  '      IdxColumnName.PO_LINE_NO.ToString, CI.PO_LINE_NO,
  '      IdxColumnName.QTY.ToString, CI.QTY,
  '      IdxColumnName.QTY_FINISH.ToString, CI.QTY_FINISH,
  '      IdxColumnName.H_QTY_PROCESS.ToString, CI.H_QTY_PROCESS
  '      )
  '      lstSql.Add(strSQL)
  '    Next
  '    If SendSQLToDB(lstSql) = True Then
  '      Return True
  '    Else
  '      SendMessageToLog("Insert to clsWMS_T_PO_LINE DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return False
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  'Private Shared Function UpdateclsWMS_T_PO_LINEToDB(ByRef Info As List(Of clsPO_LINE)) As Boolean
  '  Try
  '    If Info Is Nothing Then Return False
  '    If Info.Count = 0 Then Return True

  '    Dim strSQL As String = ""
  '    Dim rs As ADODB.Recordset = Nothing
  '    Dim lstSql As New List(Of String)
  '    For Each CI In Info
  '      strSQL = String.Format("Update {1} SET {6}={7},{8}={9},{10}={11} WHERE {2}='{3}' And {4}='{5}'",
  '      strSQL,
  '      TableName,
  '      IdxColumnName.PO_ID.ToString, CI.PO_ID,
  '      IdxColumnName.PO_LINE_NO.ToString, CI.PO_LINE_NO,
  '      IdxColumnName.QTY.ToString, CI.QTY,
  '      IdxColumnName.QTY_FINISH.ToString, CI.QTY_FINISH,
  '      IdxColumnName.H_QTY_PROCESS.ToString, CI.H_QTY_PROCESS
  '      )
  '      lstSql.Add(strSQL)
  '    Next

  '    If SendSQLToDB(lstSql) = True Then
  '      Return True
  '    Else
  '      SendMessageToLog("Update to clsWMS_T_PO_LINE DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return False
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  'Private Shared Function DeleteclsWMS_T_PO_LINEToDB(ByRef Info As List(Of clsPO_LINE)) As Boolean
  '  Try
  '    If Info Is Nothing Then Return False
  '    If Info.Count = 0 Then Return True

  '    Dim strSQL As String = ""
  '    Dim rs As ADODB.Recordset = Nothing
  '    Dim lstSql As New List(Of String)
  '    For Each CI In Info
  '      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' ",
  '      strSQL,
  '      TableName,
  '      IdxColumnName.PO_ID.ToString, CI.PO_ID,
  '      IdxColumnName.PO_LINE_NO.ToString, CI.PO_LINE_NO,
  '      IdxColumnName.QTY.ToString, CI.QTY,
  '      IdxColumnName.QTY_FINISH.ToString, CI.QTY_FINISH,
  '      IdxColumnName.H_QTY_PROCESS.ToString, CI.H_QTY_PROCESS
  '      )
  '      lstSql.Add(strSQL)
  '    Next

  '    If SendSQLToDB(lstSql) = True Then
  '      Return True
  '    Else
  '      SendMessageToLog("Delete clsWMS_T_PO_LINE DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return False
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function

  '-內部記憶體增刪修

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsPO_LINE, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_LINE_NO = "" & RowData.Item(IdxColumnName.PO_LINE_NO.ToString)
        Dim QTY = 0 & RowData.Item(IdxColumnName.QTY.ToString)
        Dim QTY_FINISH = 0 & RowData.Item(IdxColumnName.QTY_FINISH.ToString)
        Dim H_QTY_PROCESS = 0 & RowData.Item(IdxColumnName.H_QTY_PROCESS.ToString)
        Dim H_POL1 = "" & RowData.Item(IdxColumnName.H_POL1.ToString)
        Dim H_POL2 = "" & RowData.Item(IdxColumnName.H_POL2.ToString)
        Dim H_POL3 = "" & RowData.Item(IdxColumnName.H_POL3.ToString)
        Dim H_POL4 = "" & RowData.Item(IdxColumnName.H_POL4.ToString)
        Dim H_POL5 = "" & RowData.Item(IdxColumnName.H_POL5.ToString)
        Info = New clsPO_LINE(PO_ID, PO_LINE_NO, QTY, QTY_FINISH, H_QTY_PROCESS, H_POL1, H_POL2, H_POL3, H_POL4, H_POL5)
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
  '從資料庫抓取PO的資料
  Public Shared Function GetPOLineDictionaryByPOID_POLineNo(ByVal PO_ID As String, ByVal PO_Line_No As String) As Dictionary(Of String, clsPO_LINE)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_LINE)
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
          If PO_Line_No <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_LINE_NO.ToString, PO_Line_No)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_LINE_NO.ToString, PO_Line_No)
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
              Dim Info As clsPO_LINE = Nothing
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
  End Function
  '從資料庫抓取PO的資料
  Public Shared Function GetPOLineDictionaryByPOID(ByVal PO_ID As String) As Dictionary(Of String, clsPO_LINE)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_LINE)
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
              Dim Info As clsPO_LINE = Nothing
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
  End Function
  '從資料庫抓取PO的資料
  Public Shared Function GetPOLineDictionaryBydicPOID(ByVal dicPOID As Dictionary(Of String, String)) As Dictionary(Of String, clsPO_LINE)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_LINE)
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
          '    Dim Info As clsPO_LINE = Nothing
          '    If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
          '      If Info IsNot Nothing Then
          '        If ret_dic.ContainsKey(Info.gid) = False Then
          '          ret_dic.Add(Info.gid, Info)
          '        End If
          '      Else
          '        SendMessageToLog("Get clsPO_Line Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '      End If
          '    Else
          '      SendMessageToLog("Get clsPO_Line Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
                  Dim Info As clsPO_LINE = Nothing
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
End Class
