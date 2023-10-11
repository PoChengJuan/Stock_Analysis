Imports System.Collections.Concurrent
Partial Class INVXFManagement
  Public Shared TableName As String = "INVXF"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    XF001
    XF002
    XF003
    XF004
    XF005
    XF006
    XF007
    XF008
    XF009
    XF010
    XF011
    XF012
    XF013
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsINVXF) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}')",
      strSQL,
      TableName,
     IdxColumnName.XF001.ToString, Info.XF001,
     IdxColumnName.XF002.ToString, Info.XF002,
     IdxColumnName.XF003.ToString, Info.XF003,
     IdxColumnName.XF004.ToString, Info.XF004,
     IdxColumnName.XF005.ToString, Info.XF005,
     IdxColumnName.XF006.ToString, CInt(Info.XF006),
     IdxColumnName.XF007.ToString, Info.XF007,
     IdxColumnName.XF008.ToString, Info.XF008,
     IdxColumnName.XF009.ToString, Info.XF009,
     IdxColumnName.XF010.ToString, Info.XF010,
     IdxColumnName.XF011.ToString, Info.XF011,
     IdxColumnName.XF012.ToString, Info.XF012
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsINVXF) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}' WHERE {2}='{3}' AND {4}='{5}' AND {8}='{9}'",
      strSQL,
      TableName,
     IdxColumnName.XF001.ToString, Info.XF001,
     IdxColumnName.XF002.ToString, Info.XF002,
     IdxColumnName.XF003.ToString, Info.XF003,
     IdxColumnName.XF004.ToString, Info.XF004,
     IdxColumnName.XF005.ToString, Info.XF005,
     IdxColumnName.XF006.ToString, CInt(Info.XF006),
     IdxColumnName.XF007.ToString, Info.XF007,
     IdxColumnName.XF008.ToString, Info.XF008,
     IdxColumnName.XF009.ToString, Info.XF009,
     IdxColumnName.XF010.ToString, Info.XF010,
     IdxColumnName.XF011.ToString, Info.XF011,
     IdxColumnName.XF012.ToString, Info.XF012,
     IdxColumnName.XF013.ToString, Info.XF013
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsINVXF) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.XF001.ToString, Info.XF001,
     IdxColumnName.XF002.ToString, Info.XF002,
     IdxColumnName.XF004.ToString, Info.XF004
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

  Public Shared Function GetDataDictionaryByKEY(ByVal XF001 As String, ByVal XF002 As String, ByVal XF004 As String) As Dictionary(Of String, clsINVXF)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXF)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If XF001 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XF001.ToString, XF001)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XF001.ToString, XF001)
            End If
          End If
          If XF002 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XF002.ToString, XF002)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XF002.ToString, XF002)
            End If
          End If
          If XF004 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XF004.ToString, XF004)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XF004.ToString, XF004)
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
              Dim Info As clsINVXF = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsINVXF Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsINVXF Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByXF001_XF002(ByVal XF001 As String, ByVal XF002 As String) As Dictionary(Of String, clsINVXF)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXF)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""

          'If PO_ID <> "" Then
          If strWhere = "" Then
            strWhere = String.Format("WHERE {0} = '{1}' AND {2} = '{3}' ", IdxColumnName.XF001.ToString, XF001, IdxColumnName.XF002.ToString, XF002)
          Else
            strWhere = String.Format("{0} AND {1} = '{2}' AND {4} = '{5}'", strWhere, IdxColumnName.XF001.ToString, XF001, IdxColumnName.XF002.ToString, XF002)
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
              Dim Info As clsINVXF = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsINVXF Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsINVXF Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByXF009_IS_ZERO() As Dictionary(Of String, clsINVXF)
    Try
      Dim ret_dic As New Dictionary(Of String, clsINVXF)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = "WHERE XF009 IN ('0','7','9') AND XF010 like '1'"

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
              Dim Info As clsINVXF = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsINVXF Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsINVXF Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsINVXF, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim XF001 = "" & RowData.Item(IdxColumnName.XF001.ToString).ToString.Trim
        Dim XF002 = "" & RowData.Item(IdxColumnName.XF002.ToString).ToString.Trim
        Dim XF003 = "" & RowData.Item(IdxColumnName.XF003.ToString).ToString.Trim
        Dim XF004 = "" & RowData.Item(IdxColumnName.XF004.ToString).ToString.Trim
        Dim XF005 = "" & RowData.Item(IdxColumnName.XF005.ToString).ToString.Trim
        Dim XF006 = IIf(IsNumeric(RowData.Item(IdxColumnName.XF006.ToString)), RowData.Item(IdxColumnName.XF006.ToString), 0)
        Dim XF007 = "" & RowData.Item(IdxColumnName.XF007.ToString).ToString.Trim
        Dim XF008 = "" & RowData.Item(IdxColumnName.XF008.ToString).ToString.Trim
        Dim XF009 = "" & RowData.Item(IdxColumnName.XF009.ToString).ToString.Trim
        Dim XF010 = "" & RowData.Item(IdxColumnName.XF010.ToString).ToString.Trim
        Dim XF011 = "" & RowData.Item(IdxColumnName.XF011.ToString).ToString.Trim
        Dim XF012 = "" & RowData.Item(IdxColumnName.XF012.ToString).ToString.Trim
        Dim XF013 = "" & RowData.Item(IdxColumnName.XF013.ToString).ToString.Trim
        Info = New clsINVXF(XF001, XF002, XF003, XF004, XF005, XF006, XF007, XF008, XF009, XF010, XF011, XF012, XF013)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
