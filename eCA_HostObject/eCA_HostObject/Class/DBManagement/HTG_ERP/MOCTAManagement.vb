Imports System.Collections.Concurrent
Partial Class MOCTAManagement
  Public Shared TableName As String = "MOCTA"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    TA001
    TA002
    TA003
    TA004
    TA005
    TA006
    TA007
    TA015
    TA020
    TA033
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMOCTA) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}')",
      strSQL,
      TableName,
     IdxColumnName.TA001.ToString, Info.TA001,
     IdxColumnName.TA002.ToString, Info.TA002,
     IdxColumnName.TA003.ToString, Info.TA003,
     IdxColumnName.TA004.ToString, Info.TA004,
     IdxColumnName.TA005.ToString, Info.TA005,
     IdxColumnName.TA006.ToString, Info.TA006,
     IdxColumnName.TA007.ToString, Info.TA007,
     IdxColumnName.TA015.ToString, Info.TA015,
     IdxColumnName.TA020.ToString, Info.TA020,
     IdxColumnName.TA033.ToString, Info.TA033
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMOCTA) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}' WHERE {2}='{3}' AND {4}='{5}'",
      strSQL,
      TableName,
     IdxColumnName.TA001.ToString, Info.TA001,
     IdxColumnName.TA002.ToString, Info.TA002,
     IdxColumnName.TA003.ToString, Info.TA003,
     IdxColumnName.TA004.ToString, Info.TA004,
     IdxColumnName.TA005.ToString, Info.TA005,
     IdxColumnName.TA006.ToString, Info.TA006,
     IdxColumnName.TA007.ToString, Info.TA007,
     IdxColumnName.TA015.ToString, Info.TA015,
     IdxColumnName.TA020.ToString, Info.TA020,
     IdxColumnName.TA033.ToString, Info.TA033
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMOCTA) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}'",
      strSQL,
      TableName,
     IdxColumnName.TA001.ToString, Info.TA001,
     IdxColumnName.TA002.ToString, Info.TA002
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

  Public Shared Function GetDataDictionaryByKEY(ByVal TA001 As String, ByVal TA002 As String) As Dictionary(Of String, clsMOCTA)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCTA)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If TA001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TA001.ToString, TA001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TA001.ToString, TA001)
            End If
          End If
          If TA002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TA002.ToString, TA002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TA002.ToString, TA002)
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
              Dim Info As clsMOCTA = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCTA Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCTA Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByPO_ID(ByVal PO_TYPE As String, ByVal PO_ID As String) As Dictionary(Of String, clsMOCTA)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCTA)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If PO_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}' ", IdxColumnName.TA001.ToString, PO_TYPE, IdxColumnName.TA002.ToString, PO_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4}'", strWhere, IdxColumnName.TA001.ToString, PO_TYPE, IdxColumnName.TA002.ToString, PO_ID)
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
              Dim Info As clsMOCTA = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCTA Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCTA Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsMOCTA, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim TA001 = "" & RowData.Item(IdxColumnName.TA001.ToString).ToString.Trim
        Dim TA002 = "" & RowData.Item(IdxColumnName.TA002.ToString).ToString.Trim
        Dim TA003 = "" & RowData.Item(IdxColumnName.TA003.ToString).ToString.Trim
        Dim TA004 = "" & RowData.Item(IdxColumnName.TA004.ToString).ToString.Trim
        Dim TA005 = "" & RowData.Item(IdxColumnName.TA005.ToString).ToString.Trim
        Dim TA006 = "" & RowData.Item(IdxColumnName.TA006.ToString).ToString.Trim
        Dim TA007 = "" & RowData.Item(IdxColumnName.TA007.ToString).ToString.Trim
        Dim TA015 = IIf(IsNumeric(RowData.Item(IdxColumnName.TA015.ToString)), RowData.Item(IdxColumnName.TA015.ToString), 0)
        Dim TA020 = "" & RowData.Item(IdxColumnName.TA020.ToString).ToString.Trim
        Dim TA033 = "" & RowData.Item(IdxColumnName.TA033.ToString).ToString.Trim
        Info = New clsMOCTA(TA001, TA002, TA003, TA004, TA005, TA006, TA007, TA015, TA020, TA033)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
