Imports System.Collections.Concurrent
Partial Class MOCTOManagement
  Public Shared TableName As String = "MOCTO"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    TO001
    TO002
    TO003
    TO004
    TO005
    TO006
    TO007
    TO008
    TO009
    TO010
    TO017
    TO034
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMOCTO) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}')",
      strSQL,
      TableName,
     IdxColumnName.TO001.ToString, Info.TO001,
     IdxColumnName.TO002.ToString, Info.TO002,
     IdxColumnName.TO003.ToString, Info.TO003,
     IdxColumnName.TO004.ToString, Info.TO004,
     IdxColumnName.TO005.ToString, Info.TO005,
     IdxColumnName.TO006.ToString, Info.TO006,
     IdxColumnName.TO007.ToString, Info.TO007,
     IdxColumnName.TO008.ToString, Info.TO008,
     IdxColumnName.TO009.ToString, Info.TO009,
     IdxColumnName.TO010.ToString, Info.TO010,
     IdxColumnName.TO017.ToString, Info.TO017,
     IdxColumnName.TO034.ToString, Info.TO034
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMOCTO) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.TO001.ToString, Info.TO001,
     IdxColumnName.TO002.ToString, Info.TO002,
     IdxColumnName.TO003.ToString, Info.TO003,
     IdxColumnName.TO004.ToString, Info.TO004,
     IdxColumnName.TO005.ToString, Info.TO005,
     IdxColumnName.TO006.ToString, Info.TO006,
     IdxColumnName.TO007.ToString, Info.TO007,
     IdxColumnName.TO008.ToString, Info.TO008,
     IdxColumnName.TO009.ToString, Info.TO009,
     IdxColumnName.TO010.ToString, Info.TO010,
     IdxColumnName.TO017.ToString, Info.TO017,
     IdxColumnName.TO034.ToString, Info.TO034
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMOCTO) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.TO001.ToString, Info.TO001,
     IdxColumnName.TO002.ToString, Info.TO002,
     IdxColumnName.TO003.ToString, Info.TO003
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

  Public Shared Function GetDataDictionaryByKEY(ByVal TO001 As String, ByVal TO002 As String, ByVal TO003 As String) As Dictionary(Of String, clsMOCTO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCTO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If TO001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TO001.ToString, TO001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TO001.ToString, TO001)
            End If
          End If
          If TO002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TO002.ToString, TO002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TO002.ToString, TO002)
            End If
          End If
          If TO003 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TO003.ToString, TO003)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TO003.ToString, TO003)
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
              Dim Info As clsMOCTO = Nothing
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

  Public Shared Function GetDataDictionaryByPO_ID(ByVal TO001 As String, ByVal TO002 As String) As Dictionary(Of String, clsMOCTO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCTO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""

          If TO001 <> "" AndAlso TO002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}'", IdxColumnName.TO001.ToString, TO001, IdxColumnName.TO002.ToString, TO002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4}'", strWhere, IdxColumnName.TO001.ToString, TO001, IdxColumnName.TO002.ToString, TO002)
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
              Dim Info As clsMOCTO = Nothing
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
  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsMOCTO, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim TO001 = "" & RowData.Item(IdxColumnName.TO001.ToString).ToString.Trim
        Dim TO002 = "" & RowData.Item(IdxColumnName.TO002.ToString).ToString.Trim
        Dim TO003 = "" & RowData.Item(IdxColumnName.TO003.ToString).ToString.Trim
        Dim TO004 = "" & RowData.Item(IdxColumnName.TO004.ToString).ToString.Trim
        Dim TO005 = "" & RowData.Item(IdxColumnName.TO005.ToString).ToString.Trim
        Dim TO006 = "" & RowData.Item(IdxColumnName.TO006.ToString).ToString.Trim
        Dim TO007 = "" & RowData.Item(IdxColumnName.TO007.ToString).ToString.Trim
        Dim TO008 = "" & RowData.Item(IdxColumnName.TO008.ToString).ToString.Trim
        Dim TO009 = "" & RowData.Item(IdxColumnName.TO009.ToString).ToString.Trim
        Dim TO010 = "" & RowData.Item(IdxColumnName.TO010.ToString).ToString.Trim
        Dim TO017 = IIf(IsNumeric(RowData.Item(IdxColumnName.TO017.ToString)), RowData.Item(IdxColumnName.TO017.ToString), 0)
        Dim TO034 = "" & RowData.Item(IdxColumnName.TO034.ToString).ToString.Trim
        Info = New clsMOCTO(TO001, TO002, TO003, TO004, TO005, TO006, TO007, TO008, TO009, TO010, TO017, TO034)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
