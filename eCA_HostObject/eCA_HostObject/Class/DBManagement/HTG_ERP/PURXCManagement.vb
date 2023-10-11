Imports System.Collections.Concurrent
Partial Class PURXCManagement
  Public Shared TableName As String = "PURXC"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    XC001
    XC002
    XC003
    XC004
    XC005
    XC006
    XC007
    XC008
    XC009
    XC010
    XC011
    XC012
    XC013
    XC014
    XC015
    XC016
    XC017
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPURXC) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}')",
      strSQL,
      TableName,
     IdxColumnName.XC001.ToString, Info.XC001,
     IdxColumnName.XC002.ToString, Info.XC002,
     IdxColumnName.XC003.ToString, Info.XC003,
     IdxColumnName.XC004.ToString, Info.XC004,
     IdxColumnName.XC005.ToString, CInt(Info.XC005),
     IdxColumnName.XC006.ToString, Info.XC006,
     IdxColumnName.XC007.ToString, Info.XC007,
     IdxColumnName.XC008.ToString, Info.XC008,
     IdxColumnName.XC009.ToString, Info.XC009,
     IdxColumnName.XC010.ToString, Info.XC010,
     IdxColumnName.XC011.ToString, CInt(Info.XC011),
     IdxColumnName.XC012.ToString, CInt(Info.XC012),
     IdxColumnName.XC013.ToString, CInt(Info.XC013),
     IdxColumnName.XC014.ToString, Info.XC014,
     IdxColumnName.XC015.ToString, Info.XC015,
     IdxColumnName.XC016.ToString, Info.XC016,
     IdxColumnName.XC017.ToString, Info.XC017
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsPURXC) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}'",
      strSQL,
      TableName,
     IdxColumnName.XC008.ToString, Info.XC008,
     IdxColumnName.XC009.ToString, Info.XC009,
     IdxColumnName.XC010.ToString, Info.XC010,
     IdxColumnName.XC016.ToString, Info.XC016,
     IdxColumnName.XC001.ToString, Info.XC001,
     IdxColumnName.XC002.ToString, Info.XC002,
     IdxColumnName.XC003.ToString, Info.XC003,
     IdxColumnName.XC004.ToString, Info.XC004,
     IdxColumnName.XC005.ToString, CInt(Info.XC005),
     IdxColumnName.XC006.ToString, Info.XC006,
     IdxColumnName.XC007.ToString, Info.XC007,
     IdxColumnName.XC011.ToString, CInt(Info.XC011),
     IdxColumnName.XC012.ToString, CInt(Info.XC012),
     IdxColumnName.XC013.ToString, CInt(Info.XC013),
     IdxColumnName.XC014.ToString, Info.XC014,
     IdxColumnName.XC015.ToString, Info.XC015,
     IdxColumnName.XC017.ToString, Info.XC017
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsPURXC) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}'",
      strSQL,
      TableName,
     IdxColumnName.XC008.ToString, Info.XC008,
     IdxColumnName.XC009.ToString, Info.XC009,
     IdxColumnName.XC010.ToString, Info.XC010,
     IdxColumnName.XC016.ToString, Info.XC016
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

  Public Shared Function GetDataDictionaryByKEY(ByVal XC008 As String, ByVal XC009 As String, ByVal XC010 As String, ByVal XC016 As String) As Dictionary(Of String, clsPURXC)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPURXC)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If XC008 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XC008.ToString, XC008)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XC008.ToString, XC008)
            End If
          End If
          If XC009 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XC009.ToString, XC009)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XC009.ToString, XC009)
            End If
          End If
          If XC010 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XC010.ToString, XC010)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XC010.ToString, XC010)
            End If
          End If
          If XC016 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XC016.ToString, XC016)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XC016.ToString, XC016)
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
              Dim Info As clsPURXC = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPURXC Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPURXC Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsPURXC, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim XC001 = "" & RowData.Item(IdxColumnName.XC001.ToString)
        Dim XC002 = "" & RowData.Item(IdxColumnName.XC002.ToString)
        Dim XC003 = "" & RowData.Item(IdxColumnName.XC003.ToString)
        Dim XC004 = "" & RowData.Item(IdxColumnName.XC004.ToString)
        Dim XC005 = IIf(IsNumeric(RowData.Item(IdxColumnName.XC005.ToString)), RowData.Item(IdxColumnName.XC005.ToString), 0)
        Dim XC006 = "" & RowData.Item(IdxColumnName.XC006.ToString)
        Dim XC007 = "" & RowData.Item(IdxColumnName.XC007.ToString)
        Dim XC008 = "" & RowData.Item(IdxColumnName.XC008.ToString)
        Dim XC009 = "" & RowData.Item(IdxColumnName.XC009.ToString)
        Dim XC010 = "" & RowData.Item(IdxColumnName.XC010.ToString)
        Dim XC011 = IIf(IsNumeric(RowData.Item(IdxColumnName.XC011.ToString)), RowData.Item(IdxColumnName.XC011.ToString), 0)
        Dim XC012 = IIf(IsNumeric(RowData.Item(IdxColumnName.XC012.ToString)), RowData.Item(IdxColumnName.XC012.ToString), 0)
        Dim XC013 = IIf(IsNumeric(RowData.Item(IdxColumnName.XC013.ToString)), RowData.Item(IdxColumnName.XC013.ToString), 0)
        Dim XC014 = "" & RowData.Item(IdxColumnName.XC014.ToString)
        Dim XC015 = "" & RowData.Item(IdxColumnName.XC015.ToString)
        Dim XC016 = "" & RowData.Item(IdxColumnName.XC016.ToString)
        Dim XC017 = "" & RowData.Item(IdxColumnName.XC017.ToString)
        Info = New clsPURXC(XC001, XC002, XC003, XC004, XC005, XC006, XC007, XC008, XC009, XC010, XC011, XC012, XC013, XC014, XC015, XC016, XC017)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
