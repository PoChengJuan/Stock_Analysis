Imports System.Collections.Concurrent
Partial Class MOCTPManagement
  Public Shared TableName As String = "MOCTP"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    TP001
    TP002
    TP003
    TP004
    TP005
    TP006
    TP007
    TP008
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMOCTP) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}')",
      strSQL,
      TableName,
     IdxColumnName.TP001.ToString, Info.TP001,
     IdxColumnName.TP002.ToString, Info.TP002,
     IdxColumnName.TP003.ToString, Info.TP003,
     IdxColumnName.TP004.ToString, Info.TP004,
     IdxColumnName.TP005.ToString, CInt(Info.TP005),
     IdxColumnName.TP006.ToString, CInt(Info.TP006),
     IdxColumnName.TP007.ToString, Info.TP007,
     IdxColumnName.TP008.ToString, Info.TP008
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMOCTP) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}' WHERE {2}='{3}' and {4}='{5}' and {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.TP001.ToString, Info.TP001,
     IdxColumnName.TP002.ToString, Info.TP002,
     IdxColumnName.TP003.ToString, Info.TP003,
     IdxColumnName.TP004.ToString, Info.TP004,
     IdxColumnName.TP005.ToString, CInt(Info.TP005),
     IdxColumnName.TP006.ToString, CInt(Info.TP006),
     IdxColumnName.TP007.ToString, Info.TP007,
     IdxColumnName.TP008.ToString, Info.TP008
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMOCTP) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.TP001.ToString, Info.TP001,
     IdxColumnName.TP002.ToString, Info.TP002,
     IdxColumnName.TP003.ToString, Info.TP003
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

  Public Shared Function GetDataDictionaryByKEY(ByVal TP001 As String, ByVal TP002 As String, ByVal TP003 As String) As Dictionary(Of String, clsMOCTP)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCTP)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If TP001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TP001.ToString, TP001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TP001.ToString, TP001)
            End If
          End If
          If TP002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TP002.ToString, TP002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TP002.ToString, TP002)
            End If
          End If
          If TP003 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TP003.ToString, TP003)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TP003.ToString, TP003)
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
              Dim Info As clsMOCTP = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCTO Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCTO Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByPO_ID(ByVal TP001 As String, ByVal TP002 As String) As Dictionary(Of String, clsMOCTP)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCTP)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If TP002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}' ", IdxColumnName.TP001.ToString, TP001, IdxColumnName.TP002.ToString, TP002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4}' ", strWhere, IdxColumnName.TP001.ToString, TP001, IdxColumnName.TP002.ToString, TP002)
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
              Dim Info As clsMOCTP = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCTP Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCTP Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsMOCTP, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim TP001 = "" & RowData.Item(IdxColumnName.TP001.ToString).ToString.Trim
        Dim TP002 = "" & RowData.Item(IdxColumnName.TP002.ToString).ToString.Trim
        Dim TP003 = "" & RowData.Item(IdxColumnName.TP003.ToString).ToString.Trim
        Dim TP004 = "" & RowData.Item(IdxColumnName.TP004.ToString).ToString.Trim
        Dim TP005 = IIf(IsNumeric(RowData.Item(IdxColumnName.TP005.ToString)), RowData.Item(IdxColumnName.TP005.ToString), 0)
        Dim TP006 = IIf(IsNumeric(RowData.Item(IdxColumnName.TP006.ToString)), RowData.Item(IdxColumnName.TP006.ToString), 0)
        Dim TP007 = "" & RowData.Item(IdxColumnName.TP007.ToString).ToString.Trim
        Dim TP008 = "" & RowData.Item(IdxColumnName.TP008.ToString).ToString.Trim
        Info = New clsMOCTP(TP001, TP002, TP003, TP004, TP005, TP006, TP007, TP008)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
