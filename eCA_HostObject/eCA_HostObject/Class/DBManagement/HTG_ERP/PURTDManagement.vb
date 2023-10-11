Imports System.Collections.Concurrent
Partial Class PURTDManagement
  Public Shared TableName As String = "PURTD"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    TD001
    TD002
    TD003
    TD004
    TD005
    TD006
    TD007
    TD008
    TD009
    TD010
    TD011
    TD012
    TD021
    TD022
    TD023
    TD024
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPURTD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}')",
      strSQL,
      TableName,
     IdxColumnName.TD001.ToString, Info.TD001,
     IdxColumnName.TD002.ToString, Info.TD002,
     IdxColumnName.TD003.ToString, Info.TD003,
     IdxColumnName.TD004.ToString, Info.TD004,
     IdxColumnName.TD005.ToString, Info.TD005,
     IdxColumnName.TD006.ToString, Info.TD006,
     IdxColumnName.TD007.ToString, Info.TD007,
     IdxColumnName.TD008.ToString, CInt(Info.TD008),
     IdxColumnName.TD009.ToString, Info.TD009,
     IdxColumnName.TD010.ToString, CInt(Info.TD010),
     IdxColumnName.TD011.ToString, CInt(Info.TD011),
     IdxColumnName.TD012.ToString, Info.TD012,
     IdxColumnName.TD021.ToString, Info.TD021,
     IdxColumnName.TD022.ToString, Info.TD022,
     IdxColumnName.TD023.ToString, Info.TD023,
     IdxColumnName.TD024.ToString, Info.TD024
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsPURTD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.TD001.ToString, Info.TD001,
     IdxColumnName.TD002.ToString, Info.TD002,
     IdxColumnName.TD003.ToString, Info.TD003,
     IdxColumnName.TD004.ToString, Info.TD004,
     IdxColumnName.TD005.ToString, Info.TD005,
     IdxColumnName.TD006.ToString, Info.TD006,
     IdxColumnName.TD007.ToString, Info.TD007,
     IdxColumnName.TD008.ToString, CInt(Info.TD008),
     IdxColumnName.TD009.ToString, Info.TD009,
     IdxColumnName.TD010.ToString, CInt(Info.TD010),
     IdxColumnName.TD011.ToString, CInt(Info.TD011),
     IdxColumnName.TD012.ToString, Info.TD012,
     IdxColumnName.TD021.ToString, Info.TD021,
     IdxColumnName.TD022.ToString, Info.TD022,
     IdxColumnName.TD023.ToString, Info.TD023,
     IdxColumnName.TD024.ToString, Info.TD024
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsPURTD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.TD001.ToString, Info.TD001,
     IdxColumnName.TD002.ToString, Info.TD002,
     IdxColumnName.TD003.ToString, Info.TD003
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

  Public Shared Function GetDataDictionaryByKEY(ByVal TD001 As String, ByVal TD002 As String, ByVal TD003 As String) As Dictionary(Of String, clsPURTD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPURTD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If TD001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TD001.ToString, TD001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TD001.ToString, TD001)
            End If
          End If
          If TD002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TD002.ToString, TD002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TD002.ToString, TD002)
            End If
          End If
          If TD003 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TD003.ToString, TD003)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TD003.ToString, TD003)
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
              Dim Info As clsPURTD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPURTD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPURTD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByTD002(ByVal TD001 As String, ByVal TD002 As String, TD003 As String) As Dictionary(Of String, clsPURTD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPURTD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""

          If TD002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}' AND {4} = '{5}'", IdxColumnName.TD001.ToString, TD001, IdxColumnName.TD002.ToString, TD002, IdxColumnName.TD003.ToString, TD003)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4} AND {5} = '{6}' ", strWhere, IdxColumnName.TD001.ToString, TD001, IdxColumnName.TD002.ToString, TD002, IdxColumnName.TD003.ToString, TD003)
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
              Dim Info As clsPURTD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPURTD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPURTD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsPURTD, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim TD001 = "" & RowData.Item(IdxColumnName.TD001.ToString).ToString.Trim
        Dim TD002 = "" & RowData.Item(IdxColumnName.TD002.ToString).ToString.Trim
        Dim TD003 = "" & RowData.Item(IdxColumnName.TD003.ToString).ToString.Trim
        Dim TD004 = "" & RowData.Item(IdxColumnName.TD004.ToString).ToString.Trim
        Dim TD005 = "" & RowData.Item(IdxColumnName.TD005.ToString).ToString.Trim
        Dim TD006 = "" & RowData.Item(IdxColumnName.TD006.ToString).ToString.Trim
        Dim TD007 = "" & RowData.Item(IdxColumnName.TD007.ToString).ToString.Trim
        Dim TD008 = IIf(IsNumeric(RowData.Item(IdxColumnName.TD008.ToString)), RowData.Item(IdxColumnName.TD008.ToString), 0)
        Dim TD009 = "" & RowData.Item(IdxColumnName.TD009.ToString).ToString.Trim
        Dim TD010 = IIf(IsNumeric(RowData.Item(IdxColumnName.TD010.ToString)), RowData.Item(IdxColumnName.TD010.ToString), 0)
        Dim TD011 = IIf(IsNumeric(RowData.Item(IdxColumnName.TD011.ToString)), RowData.Item(IdxColumnName.TD011.ToString), 0)
        Dim TD012 = "" & RowData.Item(IdxColumnName.TD012.ToString).ToString.Trim
        Dim TD021 = "" & RowData.Item(IdxColumnName.TD021.ToString).ToString.Trim
        Dim TD022 = "" & RowData.Item(IdxColumnName.TD022.ToString).ToString.Trim
        Dim TD023 = "" & RowData.Item(IdxColumnName.TD023.ToString).ToString.Trim
        Dim TD024 = "" & RowData.Item(IdxColumnName.TD024.ToString).ToString.Trim
        Info = New clsPURTD(TD001, TD002, TD003, TD004, TD005, TD006, TD007, TD008, TD009, TD010, TD011, TD012, TD021, TD022, TD023, TD024)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
