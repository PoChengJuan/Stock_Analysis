Imports System.Collections.Concurrent
Partial Class MOCXDManagement
  Public Shared TableName As String = "MOCXD"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    XD001
    XD002
    XD003
    XD004
    XD005
    XD006
    XD007
    XD008
    XD009
    XD010
    XD011
    XD012
    XD013
    XD014
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMOCXD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}')",
      strSQL,
      TableName,
     IdxColumnName.XD001.ToString, Info.XD001,
     IdxColumnName.XD002.ToString, Info.XD002,
     IdxColumnName.XD003.ToString, Info.XD003,
     IdxColumnName.XD004.ToString, Info.XD004,
     IdxColumnName.XD005.ToString, Info.XD005,
     IdxColumnName.XD006.ToString, Info.XD006,
     IdxColumnName.XD007.ToString, CInt(Info.XD007),
     IdxColumnName.XD008.ToString, Info.XD008,
     IdxColumnName.XD009.ToString, Info.XD009,
     IdxColumnName.XD010.ToString, Info.XD010,
     IdxColumnName.XD011.ToString, Info.XD011,
     IdxColumnName.XD012.ToString, Info.XD012,
     IdxColumnName.XD013.ToString, Info.XD013,
     IdxColumnName.XD014.ToString, Info.XD014
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMOCXD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{28}='{29}' WHERE {2}='{3}' AND {4}='{5}' AND {26}='{27}'",
      strSQL,
      TableName,
     IdxColumnName.XD001.ToString, Info.XD001,
     IdxColumnName.XD002.ToString, Info.XD002,
     IdxColumnName.XD003.ToString, Info.XD003,
     IdxColumnName.XD004.ToString, Info.XD004,
     IdxColumnName.XD005.ToString, Info.XD005,
     IdxColumnName.XD006.ToString, Info.XD006,
     IdxColumnName.XD007.ToString, CInt(Info.XD007),
     IdxColumnName.XD008.ToString, Info.XD008,
     IdxColumnName.XD009.ToString, Info.XD009,
     IdxColumnName.XD010.ToString, Info.XD010,
     IdxColumnName.XD011.ToString, Info.XD011,
     IdxColumnName.XD012.ToString, Info.XD012,
     IdxColumnName.XD013.ToString, Info.XD013,
     IdxColumnName.XD014.ToString, Info.XD014
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMOCXD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.XD001.ToString, Info.XD001,
     IdxColumnName.XD002.ToString, Info.XD002,
     IdxColumnName.XD013.ToString, Info.XD013
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

  Public Shared Function GetDataDictionaryByKEY(ByVal XD001 As String, ByVal XD002 As String, ByVal XD013 As String) As Dictionary(Of String, clsMOCXD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCXD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If XD001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XD001.ToString, XD001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XD001.ToString, XD001)
            End If
          End If
          If XD002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XD002.ToString, XD002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XD002.ToString, XD002)
            End If
          End If
          If XD013 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XD013.ToString, XD013)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XD013.ToString, XD013)
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
              Dim Info As clsMOCXD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCXD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCXD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByXD011_IS_ZERO() As Dictionary(Of String, clsMOCXD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCXD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = "WHERE XD011 IN ('0','7','9')"

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
              Dim Info As clsMOCXD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCXD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCXD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByXD001_XD002(ByVal XD001 As String, ByVal XD002 As String) As Dictionary(Of String, clsMOCXD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCXD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          'If PO_ID <> "" Then
          If strWhere = "" Then
            strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}'", IdxColumnName.XD001.ToString, XD001, IdxColumnName.XD002.ToString, XD002)
          Else
            strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4}'", strWhere, IdxColumnName.XD001.ToString, XD001, IdxColumnName.XD002.ToString, XD002)
          End If
          'End If

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
              Dim Info As clsMOCXD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCXD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCXD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByXD001_XD002_XD013(ByVal XD001 As String, ByVal XD002 As String, ByVal XD013 As String) As Dictionary(Of String, clsMOCXD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCXD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          'If PO_ID <> "" Then
          If strWhere = "" Then
            strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}' AND {4} = '{5}'", IdxColumnName.XD001.ToString, XD001, IdxColumnName.XD002.ToString, XD002, IdxColumnName.XD013.ToString, XD013)
          Else
            strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4}' AND {5} = '{6}'", strWhere, IdxColumnName.XD001.ToString, XD001, IdxColumnName.XD002.ToString, XD002, IdxColumnName.XD013.ToString, XD013)
          End If
          'End If

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
              Dim Info As clsMOCXD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCXD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCXD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsMOCXD, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim XD001 = "" & RowData.Item(IdxColumnName.XD001.ToString).ToString.Trim
        Dim XD002 = "" & RowData.Item(IdxColumnName.XD002.ToString).ToString.Trim
        Dim XD003 = "" & RowData.Item(IdxColumnName.XD003.ToString).ToString.Trim
        Dim XD004 = "" & RowData.Item(IdxColumnName.XD004.ToString).ToString.Trim
        Dim XD005 = "" & RowData.Item(IdxColumnName.XD005.ToString).ToString.Trim
        Dim XD006 = "" & RowData.Item(IdxColumnName.XD006.ToString).ToString.Trim
        Dim XD007 = IIf(IsNumeric(RowData.Item(IdxColumnName.XD007.ToString)), RowData.Item(IdxColumnName.XD007.ToString), 0)
        Dim XD008 = "" & RowData.Item(IdxColumnName.XD008.ToString).ToString.Trim
        Dim XD009 = "" & RowData.Item(IdxColumnName.XD009.ToString).ToString.Trim
        Dim XD010 = "" & RowData.Item(IdxColumnName.XD010.ToString).ToString.Trim
        Dim XD011 = "" & RowData.Item(IdxColumnName.XD011.ToString).ToString.Trim
        Dim XD012 = "" & RowData.Item(IdxColumnName.XD012.ToString).ToString.Trim
        Dim XD013 = "" & RowData.Item(IdxColumnName.XD013.ToString).ToString.Trim
        Dim XD014 = "" & RowData.Item(IdxColumnName.XD014.ToString).ToString.Trim
        Info = New clsMOCXD(XD001, XD002, XD003, XD004, XD005, XD006, XD007, XD008, XD009, XD010, XD011, XD012, XD013, XD014)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
