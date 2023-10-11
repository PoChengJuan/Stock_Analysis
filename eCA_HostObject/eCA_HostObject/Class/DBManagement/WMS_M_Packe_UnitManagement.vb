Partial Class WMS_M_Packe_UnitManagement
  Public Shared TableName As String = "WMS_M_Packe_Unit"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    PACKE_UNIT
    PACKE_UNIT_NAME
    PACKE_UNIT_COMMON1
    PACKE_UNIT_COMMON2
    PACKE_UNIT_COMMON3
    PACKE_UNIT_COMMON4
    PACKE_UNIT_COMMON5
    COMMENTS
    CREATE_TIME
    UPDATE_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMPackeUnit) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}')",
      strSQL,
      TableName,
      IdxColumnName.PACKE_UNIT.ToString, Info.PACKE_UNIT,
      IdxColumnName.PACKE_UNIT_NAME.ToString, Info.PACKE_UNIT_NAME,
      IdxColumnName.PACKE_UNIT_COMMON1.ToString, Info.PACKE_UNIT_COMMON1,
      IdxColumnName.PACKE_UNIT_COMMON2.ToString, Info.PACKE_UNIT_COMMON2,
      IdxColumnName.PACKE_UNIT_COMMON3.ToString, Info.PACKE_UNIT_COMMON3,
      IdxColumnName.PACKE_UNIT_COMMON4.ToString, Info.PACKE_UNIT_COMMON4,
      IdxColumnName.PACKE_UNIT_COMMON5.ToString, Info.PACKE_UNIT_COMMON5,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME
     )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsMPackeUnit) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.PACKE_UNIT.ToString, Info.PACKE_UNIT,
      IdxColumnName.PACKE_UNIT_NAME.ToString, Info.PACKE_UNIT_NAME,
      IdxColumnName.PACKE_UNIT_COMMON1.ToString, Info.PACKE_UNIT_COMMON1,
      IdxColumnName.PACKE_UNIT_COMMON2.ToString, Info.PACKE_UNIT_COMMON2,
      IdxColumnName.PACKE_UNIT_COMMON3.ToString, Info.PACKE_UNIT_COMMON3,
      IdxColumnName.PACKE_UNIT_COMMON4.ToString, Info.PACKE_UNIT_COMMON4,
      IdxColumnName.PACKE_UNIT_COMMON5.ToString, Info.PACKE_UNIT_COMMON5,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME
      )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsMPackeUnit) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.PACKE_UNIT.ToString, Info.PACKE_UNIT,
      IdxColumnName.PACKE_UNIT_NAME.ToString, Info.PACKE_UNIT_NAME,
      IdxColumnName.PACKE_UNIT_COMMON1.ToString, Info.PACKE_UNIT_COMMON1,
      IdxColumnName.PACKE_UNIT_COMMON2.ToString, Info.PACKE_UNIT_COMMON2,
      IdxColumnName.PACKE_UNIT_COMMON3.ToString, Info.PACKE_UNIT_COMMON3,
      IdxColumnName.PACKE_UNIT_COMMON4.ToString, Info.PACKE_UNIT_COMMON4,
      IdxColumnName.PACKE_UNIT_COMMON5.ToString, Info.PACKE_UNIT_COMMON5,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME
      )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Private Shared Function SetInfoFromDB(ByRef Info As clsMPackeUnit, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PACKE_UNIT = "" & RowData.Item(IdxColumnName.PACKE_UNIT.ToString)
        Dim PACKE_UNIT_NAME = "" & RowData.Item(IdxColumnName.PACKE_UNIT_NAME.ToString)
        Dim PACKE_UNIT_COMMON1 = "" & RowData.Item(IdxColumnName.PACKE_UNIT_COMMON1.ToString)
        Dim PACKE_UNIT_COMMON2 = "" & RowData.Item(IdxColumnName.PACKE_UNIT_COMMON2.ToString)
        Dim PACKE_UNIT_COMMON3 = "" & RowData.Item(IdxColumnName.PACKE_UNIT_COMMON3.ToString)
        Dim PACKE_UNIT_COMMON4 = "" & RowData.Item(IdxColumnName.PACKE_UNIT_COMMON4.ToString)
        Dim PACKE_UNIT_COMMON5 = "" & RowData.Item(IdxColumnName.PACKE_UNIT_COMMON5.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Info = New clsMPackeUnit(PACKE_UNIT, PACKE_UNIT_NAME, PACKE_UNIT_COMMON1, PACKE_UNIT_COMMON2, PACKE_UNIT_COMMON3, PACKE_UNIT_COMMON4, PACKE_UNIT_COMMON5, COMMENTS, CREATE_TIME, UPDATE_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetWMS_M_Packe_UnitListByALL() As List(Of clsMPackeUnit)
    Try
      Dim _lstReturn As New List(Of clsMPackeUnit)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} ",
        strSQL,
        TableName
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsMPackeUnit = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info)
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetdicPackeUnitByPackeUnit(ByVal PACKE_UNIT As String) As Dictionary(Of String, clsMPackeUnit)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsMPackeUnit)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strUniqueIDList As String = ""


        If strWhere = "" Then
          strWhere = String.Format("WHERE {0} IN ('{1}') ", IdxColumnName.PACKE_UNIT.ToString, PACKE_UNIT)
        Else
          strWhere = String.Format("{0} AND {1} = ('{2}') ", strWhere, IdxColumnName.PACKE_UNIT.ToString, PACKE_UNIT)
        End If
        strSQL = String.Format("Select * from {1} {2} ",
            strSQL,
            TableName,
            strWhere
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsMPackeUnit = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
