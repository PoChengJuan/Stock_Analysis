Imports System.Collections.Concurrent
Partial Class WMS_T_PO_MERGEManagement
  Public Shared TableName As String = "WMS_T_PO_MERGE"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    PO_ID = 0
    PO_SERIAL_NO = 1
    WO_ID = 2
    WO_SERIAL_NO = 3
    QTY = 4
    QTY_PROCESS = 5
    QTY_FINISH = 6
    CLOSE_UUID = 7
    COMMENTS = 8
    CREATE_TIME = 9
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPO_MERGE) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}')",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_SERIAL_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.CLOSE_UUID.ToString, Info.CLOSE_UUID,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsPO_MERGE) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_SERIAL_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.CLOSE_UUID.ToString, Info.CLOSE_UUID,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME
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
  'Public Shared Function GetUpdateSQLForChangeValue(ByRef Info As clsPO_MERGE, ByRef dicChangeColumnValue As Dictionary(Of String, String)) As String
  '  Try
  '    Dim strSQL As String = ""
  '    Dim strUpdateColumnValue As String = ""
  '    If O_Get_UpdateColumnSQL(Of IdxColumnName)(dicChangeColumnValue, strUpdateColumnValue) = True Then
  '      If strUpdateColumnValue <> "" Then
  '        strSQL = String.Format("Update {1} SET {2}  WHERE {3}='{4}' And {5}='{6}' And {7}='{8}' And {9}='{10}'",
  '        strSQL,
  '        TableName,
  '        strUpdateColumnValue,
  '        IdxColumnName.PO_ID.ToString, Info.PO_ID,
  '        IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
  '        IdxColumnName.WO_ID.ToString, Info.WO_ID,
  '        IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_SERIAL_NO,
  '        IdxColumnName.QTY.ToString, Info.QTY,
  '        IdxColumnName.QTY_PROCESS.ToString, Info.QTY_PROCESS,
  '        IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
  '        IdxColumnName.CLOSE_UUID.ToString, Info.CLOSE_UUID,
  '        IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
  '        IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME
  '        )
  '        Dim NewSQL As String = ""
  '        If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
  '          Return NewSQL
  '        End If
  '      End If
  '    End If
  '    Return ""
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return ""
  '  End Try
  'End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsPO_MERGE) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
      IdxColumnName.WO_ID.ToString, Info.WO_ID,
      IdxColumnName.WO_SERIAL_NO.ToString, Info.WO_SERIAL_NO,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.QTY_FINISH,
      IdxColumnName.CLOSE_UUID.ToString, Info.CLOSE_UUID,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsPO_MERGE, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_SERIAL_NO = "" & RowData.Item(IdxColumnName.PO_SERIAL_NO.ToString)
        Dim WO_ID = "" & RowData.Item(IdxColumnName.WO_ID.ToString)
        Dim WO_SERIAL_NO = "" & RowData.Item(IdxColumnName.WO_SERIAL_NO.ToString)
        Dim QTY = If(IsNumeric(RowData.Item(IdxColumnName.QTY.ToString)), RowData.Item(IdxColumnName.QTY.ToString), 0 & RowData.Item(IdxColumnName.QTY.ToString))
        Dim QTY_PROCESS = If(IsNumeric(RowData.Item(IdxColumnName.QTY_PROCESS.ToString)), RowData.Item(IdxColumnName.QTY_PROCESS.ToString), 0 & RowData.Item(IdxColumnName.QTY_PROCESS.ToString))
        Dim QTY_FINISH = If(IsNumeric(RowData.Item(IdxColumnName.QTY_FINISH.ToString)), RowData.Item(IdxColumnName.QTY_FINISH.ToString), 0 & RowData.Item(IdxColumnName.QTY_FINISH.ToString))
        Dim CLOSE_UUID = "" '& RowData.Item(IdxColumnName.CLOSE_UUID.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Info = New clsPO_MERGE(PO_ID, PO_SERIAL_NO, WO_ID, WO_SERIAL_NO, QTY, QTY_PROCESS, QTY_FINISH, CLOSE_UUID, COMMENTS, CREATE_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetDataDicByALL() As Dictionary(Of String, clsPO_MERGE)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsPO_MERGE)
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
            Dim Info As clsPO_MERGE = Nothing
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
  '從資料庫抓取PO_DTL的資料
  Public Shared Function GetPO_MergeDictionaryBydicPOID(ByVal dicPOID As Dictionary(Of String, String)) As Dictionary(Of String, clsPO_MERGE)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_MERGE)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          Dim strPOList As String = ""
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet


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
                  Dim Info As clsPO_MERGE = Nothing
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
  Public Shared Function GetPO_MergeDictionaryByPOID_PO_Serial_No(ByVal PO_ID As String, ByVal PO_Serial_NO As String) As Dictionary(Of String, clsPO_MERGE)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO_MERGE)
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
              Dim Info As clsPO_MERGE = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid()) = False Then
                    ret_dic.Add(Info.gid(), Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO_MERGE Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO_MERGE Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
