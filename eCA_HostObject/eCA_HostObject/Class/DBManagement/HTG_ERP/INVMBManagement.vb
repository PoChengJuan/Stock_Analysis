Imports System.Collections.Concurrent
Partial Class INVMBManagement
  Public Shared TableName As String = "INVMB"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    MB001
    MB002
    MB003
    MB004
    MB005
    MB006
    MB007
    MB008
    MB009
    MB014
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsINVMB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}')",
      strSQL,
      TableName,
     IdxColumnName.MB001.ToString, Info.MB001,
     IdxColumnName.MB002.ToString, Info.MB002,
     IdxColumnName.MB003.ToString, Info.MB003,
     IdxColumnName.MB004.ToString, Info.MB004,
     IdxColumnName.MB005.ToString, Info.MB005,
     IdxColumnName.MB006.ToString, Info.MB006,
     IdxColumnName.MB007.ToString, Info.MB007,
     IdxColumnName.MB008.ToString, Info.MB008,
     IdxColumnName.MB009.ToString, Info.MB009,
     IdxColumnName.MB014.ToString, CInt(Info.MB014)
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsINVMB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}' WHERE {2}='{3}'",
      strSQL,
      TableName,
     IdxColumnName.MB001.ToString, Info.MB001,
     IdxColumnName.MB002.ToString, Info.MB002,
     IdxColumnName.MB003.ToString, Info.MB003,
     IdxColumnName.MB004.ToString, Info.MB004,
     IdxColumnName.MB005.ToString, Info.MB005,
     IdxColumnName.MB006.ToString, Info.MB006,
     IdxColumnName.MB007.ToString, Info.MB007,
     IdxColumnName.MB008.ToString, Info.MB008,
     IdxColumnName.MB009.ToString, Info.MB009,
     IdxColumnName.MB014.ToString, CInt(Info.MB014)
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsINVMB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}'",
      strSQL,
      TableName,
     IdxColumnName.MB001.ToString, Info.MB001
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

  Public Shared Function GetDataDictionaryByKEY(ByVal MB001 As String) As Dictionary(Of String, clsINVMB)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVMB)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If MB001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.MB001.ToString, MB001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.MB001.ToString, MB001)
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
              Dim Info As clsINVMB = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsINVMB Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsINVMB Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsINVMB, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim MB001 = "" & RowData.Item(IdxColumnName.MB001.ToString).ToString.Trim
        Dim MB002 = "" & RowData.Item(IdxColumnName.MB002.ToString).ToString.Trim
        Dim MB003 = "" & RowData.Item(IdxColumnName.MB003.ToString).ToString.Trim
        Dim MB004 = "" & RowData.Item(IdxColumnName.MB004.ToString).ToString.Trim
        Dim MB005 = "" & RowData.Item(IdxColumnName.MB005.ToString).ToString.Trim
        Dim MB006 = "" & RowData.Item(IdxColumnName.MB006.ToString).ToString.Trim
        Dim MB007 = "" & RowData.Item(IdxColumnName.MB007.ToString).ToString.Trim
        Dim MB008 = "" & RowData.Item(IdxColumnName.MB008.ToString).ToString.Trim
        Dim MB009 = "" & RowData.Item(IdxColumnName.MB009.ToString).ToString.Trim
        Dim MB014 = IIf(IsNumeric(RowData.Item(IdxColumnName.MB014.ToString)), RowData.Item(IdxColumnName.MB014.ToString), 0)
        Info = New clsINVMB(MB001, MB002, MB003, MB004, MB005, MB006, MB007, MB008, MB009, MB014)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
