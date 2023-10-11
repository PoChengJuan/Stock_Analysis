Imports System.Collections.Concurrent
Partial Class INVXDManagement
  Public Shared TableName As String = "INVXD"
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
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsINVXD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}')",
      strSQL,
      TableName,
     IdxColumnName.XD001.ToString, Info.XD001,
     IdxColumnName.XD002.ToString, Info.XD002,
     IdxColumnName.XD003.ToString, Info.XD003,
     IdxColumnName.XD004.ToString, Info.XD004,
     IdxColumnName.XD005.ToString, Info.XD005,
     IdxColumnName.XD006.ToString, CInt(Info.XD006),
     IdxColumnName.XD007.ToString, Info.XD007,
     IdxColumnName.XD008.ToString, Info.XD008,
     IdxColumnName.XD009.ToString, Info.XD009,
     IdxColumnName.XD010.ToString, Info.XD010,
     IdxColumnName.XD011.ToString, Info.XD011,
     IdxColumnName.XD012.ToString, Info.XD012
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsINVXD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.XD001.ToString, Info.XD001,
     IdxColumnName.XD002.ToString, Info.XD002,
     IdxColumnName.XD003.ToString, Info.XD003,
     IdxColumnName.XD004.ToString, Info.XD004,
     IdxColumnName.XD005.ToString, Info.XD005,
     IdxColumnName.XD006.ToString, CInt(Info.XD006),
     IdxColumnName.XD007.ToString, Info.XD007,
     IdxColumnName.XD008.ToString, Info.XD008,
     IdxColumnName.XD009.ToString, Info.XD009,
     IdxColumnName.XD010.ToString, Info.XD010,
     IdxColumnName.XD011.ToString, Info.XD011,
     IdxColumnName.XD012.ToString, Info.XD012,
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsINVXD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.XD001.ToString, Info.XD001,
     IdxColumnName.XD002.ToString, Info.XD002,
     IdxColumnName.XD004.ToString, Info.XD004
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

  Public Shared Function GetDataDictionaryByKEY(ByVal XD001 As String, ByVal XD002 As String, ByVal XD004 As String) As Dictionary(Of String, clsINVXD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXD)
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
          If XD004 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XD004.ToString, XD004)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XD004.ToString, XD004)
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
              Dim Info As clsINVXD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsINVXD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsINVXD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByPO_ID(ByVal PO_ID As String) As Dictionary(Of String, clsINVXD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""

          If PO_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XD002.ToString, PO_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XD002.ToString, PO_ID)
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
              Dim Info As clsINVXD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsINVXD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsINVXD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByXD009_IS_ZERO() As Dictionary(Of String, clsINVXD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = "WHERE XD009 IN ('0','7','9') AND XD010 like '1'"

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
              Dim Info As clsINVXD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsINVXD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsINVXD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsINVXD, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim XD001 = "" & RowData.Item(IdxColumnName.XD001.ToString).ToString.Trim
        Dim XD002 = "" & RowData.Item(IdxColumnName.XD002.ToString).ToString.Trim
        Dim XD003 = "" & RowData.Item(IdxColumnName.XD003.ToString).ToString.Trim
        Dim XD004 = "" & RowData.Item(IdxColumnName.XD004.ToString).ToString.Trim
        Dim XD005 = "" & RowData.Item(IdxColumnName.XD005.ToString).ToString.Trim
        Dim XD006 = IIf(IsNumeric(RowData.Item(IdxColumnName.XD006.ToString)), RowData.Item(IdxColumnName.XD006.ToString), 0)
        Dim XD007 = "" & RowData.Item(IdxColumnName.XD007.ToString).ToString.Trim
        Dim XD008 = "" & RowData.Item(IdxColumnName.XD008.ToString).ToString.Trim
        Dim XD009 = "" & RowData.Item(IdxColumnName.XD009.ToString).ToString.Trim
        Dim XD010 = "" & RowData.Item(IdxColumnName.XD010.ToString).ToString.Trim
        Dim XD011 = "" & RowData.Item(IdxColumnName.XD011.ToString).ToString.Trim
        Dim XD012 = "" & RowData.Item(IdxColumnName.XD012.ToString).ToString.Trim
        Dim XD013 = "" & RowData.Item(IdxColumnName.XD013.ToString).ToString.Trim
        Info = New clsINVXD(XD001, XD002, XD003, XD004, XD005, XD006, XD007, XD008, XD009, XD010, XD011, XD012, XD013)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
