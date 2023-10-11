Imports System.Collections.Concurrent
Partial Class MOCTBManagement
  Public Shared TableName As String = "MOCTB"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    TB001
    TB002
    TB003
    TB004
    TB005
    TB006
    TB007
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMOCTB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}')",
      strSQL,
      TableName,
     IdxColumnName.TB001.ToString, Info.TB001,
     IdxColumnName.TB002.ToString, Info.TB002,
     IdxColumnName.TB003.ToString, Info.TB003,
     IdxColumnName.TB004.ToString, CInt(Info.TB004),
     IdxColumnName.TB005.ToString, CInt(Info.TB005),
     IdxColumnName.TB006.ToString, Info.TB006,
     IdxColumnName.TB007.ToString, Info.TB007
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMOCTB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}'",
      strSQL,
      TableName,
     IdxColumnName.TB001.ToString, Info.TB001,
     IdxColumnName.TB002.ToString, Info.TB002,
     IdxColumnName.TB003.ToString, Info.TB003,
     IdxColumnName.TB004.ToString, CInt(Info.TB004),
     IdxColumnName.TB005.ToString, CInt(Info.TB005),
     IdxColumnName.TB006.ToString, Info.TB006,
     IdxColumnName.TB007.ToString, Info.TB007
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMOCTB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}'",
      strSQL,
      TableName,
     IdxColumnName.TB001.ToString, Info.TB001,
     IdxColumnName.TB002.ToString, Info.TB002,
     IdxColumnName.TB003.ToString, Info.TB003,
     IdxColumnName.TB006.ToString, Info.TB006
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

  Public Shared Function GetDataDictionaryByKEY(ByVal TB001 As String, ByVal TB002 As String, ByVal TB003 As String, ByVal TB006 As String) As Dictionary(Of String, clsMOCTB)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCTB)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If TB001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TB001.ToString, TB001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TB001.ToString, TB001)
            End If
          End If
          If TB002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TB002.ToString, TB002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TB002.ToString, TB002)
            End If
          End If
          If TB003 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TB003.ToString, TB003)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TB003.ToString, TB003)
            End If
          End If
          If TB006 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TB006.ToString, TB006)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TB006.ToString, TB006)
            End If
          End If
          Dim strSQL As String = String.Empty
          Dim rs As DataSet = Nothing
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
              Dim Info As clsMOCTB = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCTB Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCTB Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByPO_ID(ByVal PO_TYPE As String, ByVal PO_ID As String) As Dictionary(Of String, clsMOCTB)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCTB)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If PO_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}' ", IdxColumnName.TB001.ToString, PO_TYPE, IdxColumnName.TB002.ToString, PO_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4}' ", strWhere, IdxColumnName.TB001.ToString, PO_TYPE, IdxColumnName.TB002.ToString, PO_ID)
            End If
          End If

          Dim strSQL As String = String.Empty
          Dim rs As DataSet = Nothing
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
              Dim Info As clsMOCTB = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCTB Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCTB Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsMOCTB, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim TB001 = "" & RowData.Item(IdxColumnName.TB001.ToString).ToString.Trim
        Dim TB002 = "" & RowData.Item(IdxColumnName.TB002.ToString).ToString.Trim
        Dim TB003 = "" & RowData.Item(IdxColumnName.TB003.ToString).ToString.Trim
        Dim TB004 = IIf(IsNumeric(RowData.Item(IdxColumnName.TB004.ToString)), RowData.Item(IdxColumnName.TB004.ToString), 0)
        Dim TB005 = IIf(IsNumeric(RowData.Item(IdxColumnName.TB005.ToString)), RowData.Item(IdxColumnName.TB005.ToString), 0)
        Dim TB006 = "" & RowData.Item(IdxColumnName.TB006.ToString).ToString.Trim
        Dim TB007 = "" & RowData.Item(IdxColumnName.TB007.ToString).ToString.Trim
        Info = New clsMOCTB(TB001, TB002, TB003, TB004, TB005, TB006, TB007)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
