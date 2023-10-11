﻿Imports System.Collections.Concurrent
Partial Class MOCXBManagement
  Public Shared TableName As String = "MOCXB"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    XB001
    XB002
    XB003
    XB004
    XB005
    XB006
    XB007
    XB008
    XB009
    XB010
    XB011
    XB012
    XB013
    XB014
    XB015
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMOCXB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}')",
      strSQL,
      TableName,
     IdxColumnName.XB001.ToString, Info.XB001,
     IdxColumnName.XB002.ToString, Info.XB002,
     IdxColumnName.XB003.ToString, Info.XB003,
     IdxColumnName.XB004.ToString, Info.XB004,
     IdxColumnName.XB005.ToString, Info.XB005,
     IdxColumnName.XB006.ToString, Info.XB006,
     IdxColumnName.XB007.ToString, Info.XB007,
     IdxColumnName.XB008.ToString, CInt(Info.XB008),
     IdxColumnName.XB009.ToString, Info.XB009,
     IdxColumnName.XB010.ToString, Info.XB010,
     IdxColumnName.XB011.ToString, Info.XB011,
     IdxColumnName.XB012.ToString, Info.XB012,
     IdxColumnName.XB013.ToString, Info.XB013,
     IdxColumnName.XB014.ToString, Info.XB014,
     IdxColumnName.XB015.ToString, Info.XB015
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMOCXB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.XB010.ToString, Info.XB010,
     IdxColumnName.XB011.ToString, Info.XB011,
     IdxColumnName.XB013.ToString, Info.XB013,
     IdxColumnName.XB001.ToString, Info.XB001,
     IdxColumnName.XB002.ToString, Info.XB002,
     IdxColumnName.XB003.ToString, Info.XB003,
     IdxColumnName.XB004.ToString, Info.XB004,
     IdxColumnName.XB005.ToString, Info.XB005,
     IdxColumnName.XB006.ToString, Info.XB006,
     IdxColumnName.XB007.ToString, Info.XB007,
     IdxColumnName.XB008.ToString, CInt(Info.XB008),
     IdxColumnName.XB009.ToString, Info.XB009,
     IdxColumnName.XB012.ToString, Info.XB012,
     IdxColumnName.XB014.ToString, Info.XB014,
     IdxColumnName.XB015.ToString, Info.XB015
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMOCXB) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.XB010.ToString, Info.XB010,
     IdxColumnName.XB011.ToString, Info.XB011,
     IdxColumnName.XB013.ToString, Info.XB013
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

  Public Shared Function GetDataDictionaryByKEY(ByVal XB010 As String, ByVal XB011 As String, ByVal XB013 As String) As Dictionary(Of String, clsMOCXB)
    Try
      Dim ret_dic As New Dictionary(Of String, clsMOCXB)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If XB010 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XB010.ToString, XB010)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XB010.ToString, XB010)
            End If
          End If
          If XB011 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XB011.ToString, XB011)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XB011.ToString, XB011)
            End If
          End If
          If XB013 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.XB013.ToString, XB013)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.XB013.ToString, XB013)
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
              Dim Info As clsMOCXB = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsMOCXB Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsMOCXB Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsMOCXB, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim XB001 = "" & RowData.Item(IdxColumnName.XB001.ToString)
        Dim XB002 = "" & RowData.Item(IdxColumnName.XB002.ToString)
        Dim XB003 = "" & RowData.Item(IdxColumnName.XB003.ToString)
        Dim XB004 = "" & RowData.Item(IdxColumnName.XB004.ToString)
        Dim XB005 = "" & RowData.Item(IdxColumnName.XB005.ToString)
        Dim XB006 = "" & RowData.Item(IdxColumnName.XB006.ToString)
        Dim XB007 = "" & RowData.Item(IdxColumnName.XB007.ToString)
        Dim XB008 = IIf(IsNumeric(RowData.Item(IdxColumnName.XB008.ToString)), RowData.Item(IdxColumnName.XB008.ToString), 0)
        Dim XB009 = "" & RowData.Item(IdxColumnName.XB009.ToString)
        Dim XB010 = "" & RowData.Item(IdxColumnName.XB010.ToString)
        Dim XB011 = "" & RowData.Item(IdxColumnName.XB011.ToString)
        Dim XB012 = "" & RowData.Item(IdxColumnName.XB012.ToString)
        Dim XB013 = "" & RowData.Item(IdxColumnName.XB013.ToString)
        Dim XB014 = "" & RowData.Item(IdxColumnName.XB014.ToString)
        Dim XB015 = "" & RowData.Item(IdxColumnName.XB015.ToString)
        Info = New clsMOCXB(XB001, XB002, XB003, XB004, XB005, XB006, XB007, XB008, XB009, XB010, XB011, XB012, XB013, XB014, XB015)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
