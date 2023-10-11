Imports System.Collections.Concurrent
Partial Class PURTCManagement
  Public Shared TableName As String = "PURTC"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    TC001
    TC002
    TC003
    TC014
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPURTC) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8}) values ('{3}','{5}','{7}','{9}')",
      strSQL,
      TableName,
     IdxColumnName.TC001.ToString, Info.TC001,
     IdxColumnName.TC002.ToString, Info.TC002,
     IdxColumnName.TC003.ToString, Info.TC003,
     IdxColumnName.TC014.ToString, Info.TC014
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsPURTC) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}',{8}='{9}' WHERE {2}='{3}' AND {4}='{5}'",
      strSQL,
      TableName,
     IdxColumnName.TC001.ToString, Info.TC001,
     IdxColumnName.TC002.ToString, Info.TC002,
     IdxColumnName.TC003.ToString, Info.TC003,
     IdxColumnName.TC014.ToString, Info.TC014
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsPURTC) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}'",
      strSQL,
      TableName,
     IdxColumnName.TC001.ToString, Info.TC001,
     IdxColumnName.TC002.ToString, Info.TC002
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

  Public Shared Function GetDataDictionaryByKEY(ByVal TC001 As String, ByVal TC002 As String) As Dictionary(Of String, clsPURTC)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPURTC)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If TC001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TC001.ToString, TC001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TC001.ToString, TC001)
            End If
          End If
          If TC002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TC002.ToString, TC002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TC002.ToString, TC002)
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
              Dim Info As clsPURTC = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPURTC Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPURTC Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByTC002(ByVal TC001 As String, ByVal TC002 As String) As Dictionary(Of String, clsPURTC)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPURTC)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""

          If TC002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}' ", IdxColumnName.TC001.ToString, TC001, IdxColumnName.TC002.ToString, TC002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4}' ", strWhere, IdxColumnName.TC001.ToString, TC001, IdxColumnName.TC002.ToString, TC002)
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
              Dim Info As clsPURTC = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPURTC Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPURTC Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsPURTC, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim TC001 = "" & RowData.Item(IdxColumnName.TC001.ToString).ToString.Trim
        Dim TC002 = "" & RowData.Item(IdxColumnName.TC002.ToString).ToString.Trim
        Dim TC003 = "" & RowData.Item(IdxColumnName.TC003.ToString).ToString.Trim
        Dim TC014 = "" & RowData.Item(IdxColumnName.TC014.ToString).ToString.Trim
        Info = New clsPURTC(TC001, TC002, TC003, TC014)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
