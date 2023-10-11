Imports System.Collections.Concurrent
Partial Class INVXBManagement
  Public Shared TableName As String = "INVXB"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    XB001
    XB002
    XB003
    XB004
    XB005
    XB007
    XB008
    XB009
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsINVXB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}')",
      strSQL,
      TableName,
     IdxColumnName.XB001.ToString, Info.XB001,
     IdxColumnName.XB002.ToString, Info.XB002,
     IdxColumnName.XB003.ToString, Info.XB003,
     IdxColumnName.XB004.ToString, Info.XB004,
     IdxColumnName.XB005.ToString, Info.XB005,
     IdxColumnName.XB007.ToString, Info.XB007,
     IdxColumnName.XB008.ToString, Info.XB008,
     IdxColumnName.XB009.ToString, Info.XB009
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsINVXB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}' WHERE {2}='{3}'",
      strSQL,
      TableName,
     IdxColumnName.XB001.ToString, Info.XB001,
     IdxColumnName.XB002.ToString, Info.XB002,
     IdxColumnName.XB003.ToString, Info.XB003,
     IdxColumnName.XB004.ToString, Info.XB004,
     IdxColumnName.XB005.ToString, Info.XB005,
     IdxColumnName.XB007.ToString, Info.XB007,
     IdxColumnName.XB008.ToString, Info.XB008,
     IdxColumnName.XB009.ToString, Info.XB009
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsINVXB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}'",
      strSQL,
      TableName,
     IdxColumnName.XB001.ToString, Info.XB001
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

  Public Shared Function GetDataDictionaryByKEY(ByVal XB001 As String) As Dictionary(Of String, clsINVXB)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXB)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If XB001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}%' ", IdxColumnName.XB001.ToString, XB001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}%' ", strWhere, IdxColumnName.XB001.ToString, XB001)
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
              Dim Info As clsINVXB = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsINVXB Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsINVXB Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByXB008_IS_ZERO() As Dictionary(Of String, clsINVXB)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXB)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = "WHERE XB008 IN ('0', '5')"

          Dim strSQL As String = String.Empty
          Dim rs As DataSet = Nothing
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select TOP(1000) * from {1} {2} ",
          strSQL,
          TableName,
          strWhere
          )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsINVXB = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsEPSXB Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsEPSXB Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByXB008_IS_ZEROBySKU(ByVal SKU_NO As String) As Dictionary(Of String, clsINVXB)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXB)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = "WHERE XB008 IN ('0', '5') AND XB001 like '" & SKU_NO & "%'"

          Dim strSQL As String = String.Empty
          Dim rs As DataSet = Nothing
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select TOP(2000) * from {1} {2} ",
          strSQL,
          TableName,
          strWhere
          )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsINVXB = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsEPSXB Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsEPSXB Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsINVXB, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim XB001 = "" & RowData.Item(IdxColumnName.XB001.ToString).ToString.Trim
        Dim XB002 = "" & RowData.Item(IdxColumnName.XB002.ToString).ToString.Trim
        Dim XB003 = "" & RowData.Item(IdxColumnName.XB003.ToString).ToString.Trim
        Dim XB004 = "" & RowData.Item(IdxColumnName.XB004.ToString).ToString.Trim
        Dim XB005 = "" & RowData.Item(IdxColumnName.XB005.ToString).ToString.Trim
        Dim XB007 = "" & RowData.Item(IdxColumnName.XB007.ToString).ToString.Trim
        Dim XB008 = "" & RowData.Item(IdxColumnName.XB008.ToString).ToString.Trim
        Dim XB009 = "" & RowData.Item(IdxColumnName.XB009.ToString).ToString.Trim

        XB001 = XB001.ToString.Replace("'", "''")
        XB002 = XB002.ToString.Replace("'", "''")
        XB003 = XB003.ToString.Replace("'", "''")
        XB004 = XB004.ToString.Replace("'", "''")
        XB005 = XB005.ToString.Replace("'", "''")
        XB007 = XB007.ToString.Replace("'", "''")
        XB008 = XB008.ToString.Replace("'", "''")
        XB009 = XB009.ToString.Replace("'", "''")

        Info = New clsINVXB(XB001, XB002, XB003, XB004, XB005, XB007, XB008, XB009)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
