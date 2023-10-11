Imports System.Collections.Concurrent
Partial Class PURTFManagement
  Public Shared TableName As String = "PURTF"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    TF001
    TF002
    TF003
    TF004
    TF005
    TF006
    TF007
    TF008
    TF009
    TF010
    TF011
    TF012
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPURTF) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}')",
      strSQL,
      TableName,
     IdxColumnName.TF001.ToString, Info.TF001,
     IdxColumnName.TF002.ToString, Info.TF002,
     IdxColumnName.TF003.ToString, Info.TF003,
     IdxColumnName.TF004.ToString, Info.TF004,
     IdxColumnName.TF005.ToString, Info.TF005,
     IdxColumnName.TF006.ToString, Info.TF006,
     IdxColumnName.TF007.ToString, Info.TF007,
     IdxColumnName.TF008.ToString, Info.TF008,
     IdxColumnName.TF009.ToString, CInt(Info.TF009),
     IdxColumnName.TF010.ToString, Info.TF010,
     IdxColumnName.TF011.ToString, CInt(Info.TF011),
     IdxColumnName.TF012.ToString, CInt(Info.TF012)
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsPURTF) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.TF001.ToString, Info.TF001,
     IdxColumnName.TF002.ToString, Info.TF002,
     IdxColumnName.TF003.ToString, Info.TF003,
     IdxColumnName.TF004.ToString, Info.TF004,
     IdxColumnName.TF005.ToString, Info.TF005,
     IdxColumnName.TF006.ToString, Info.TF006,
     IdxColumnName.TF007.ToString, Info.TF007,
     IdxColumnName.TF008.ToString, Info.TF008,
     IdxColumnName.TF009.ToString, CInt(Info.TF009),
     IdxColumnName.TF010.ToString, Info.TF010,
     IdxColumnName.TF011.ToString, CInt(Info.TF011),
     IdxColumnName.TF012.ToString, CInt(Info.TF012)
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsPURTF) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.TF001.ToString, Info.TF001,
     IdxColumnName.TF002.ToString, Info.TF002,
     IdxColumnName.TF003.ToString, Info.TF003
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

  Public Shared Function GetDataDictionaryByKEY(ByVal TF001 As String, ByVal TF002 As String, ByVal TF003 As String) As Dictionary(Of String, clsPURTF)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPURTF)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If TF001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TF001.ToString, TF001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TF001.ToString, TF001)
            End If
          End If
          If TF002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TF002.ToString, TF002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TF002.ToString, TF002)
            End If
          End If
          If TF003 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.TF003.ToString, TF003)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.TF003.ToString, TF003)
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
              Dim Info As clsPURTF = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPURTF Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPURTF Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByTF002(ByVal TF001 As String, ByVal TF002 As String, ByVal TF004 As String) As Dictionary(Of String, clsPURTF)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPURTF)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""

          If TF002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}' AND {4} = '{5}' ", IdxColumnName.TF001.ToString, TF001, IdxColumnName.TF002.ToString, TF002, IdxColumnName.TF004.ToString, TF004)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' AND {3} = '{4}' AND {5} = '{6}' ", strWhere, IdxColumnName.TF001.ToString, TF001, IdxColumnName.TF002.ToString, TF002, IdxColumnName.TF004.ToString, TF004)
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
              Dim Info As clsPURTF = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPURTF Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPURTF Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsPURTF, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim TF001 = "" & RowData.Item(IdxColumnName.TF001.ToString).ToString.Trim
        Dim TF002 = "" & RowData.Item(IdxColumnName.TF002.ToString).ToString.Trim
        Dim TF003 = "" & RowData.Item(IdxColumnName.TF003.ToString).ToString.Trim
        Dim TF004 = "" & RowData.Item(IdxColumnName.TF004.ToString).ToString.Trim
        Dim TF005 = "" & RowData.Item(IdxColumnName.TF005.ToString).ToString.Trim
        Dim TF006 = "" & RowData.Item(IdxColumnName.TF006.ToString).ToString.Trim
        Dim TF007 = "" & RowData.Item(IdxColumnName.TF007.ToString).ToString.Trim
        Dim TF008 = "" & RowData.Item(IdxColumnName.TF008.ToString).ToString.Trim
        Dim TF009 = IIf(IsNumeric(RowData.Item(IdxColumnName.TF009.ToString)), RowData.Item(IdxColumnName.TF009.ToString), 0)
        Dim TF010 = "" & RowData.Item(IdxColumnName.TF010.ToString).ToString.Trim
        Dim TF011 = IIf(IsNumeric(RowData.Item(IdxColumnName.TF011.ToString)), RowData.Item(IdxColumnName.TF011.ToString), 0)
        Dim TF012 = IIf(IsNumeric(RowData.Item(IdxColumnName.TF012.ToString)), RowData.Item(IdxColumnName.TF012.ToString), 0)
        Info = New clsPURTF(TF001, TF002, TF003, TF004, TF005, TF006, TF007, TF008, TF009, TF010, TF011, TF012)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
