Partial Class WMS_CT_GUID_LabelManagement
  Public Shared TableName As String = "WMS_CT_GUID_Label"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    GUID
    UniqueID
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsCTGUIDLabel) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4}) values ('{3}','{5}')",
      strSQL,
      TableName,
      IdxColumnName.GUID.ToString, Info.GUID,
      IdxColumnName.UniqueID.ToString, Info.UniqueID
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsCTGUIDLabel) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.GUID.ToString, Info.GUID,
      IdxColumnName.UniqueID.ToString, Info.UniqueID
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsCTGUIDLabel) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.GUID.ToString, Info.GUID,
      IdxColumnName.UniqueID.ToString, Info.UniqueID
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsCTGUIDLabel, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim GUID = "" & RowData.Item(IdxColumnName.GUID.ToString)
        Dim UniqueID = If(IsNumeric(RowData.Item(IdxColumnName.UniqueID.ToString)), RowData.Item(IdxColumnName.UniqueID.ToString), 0 & RowData.Item(IdxColumnName.UniqueID.ToString))
        Info = New clsCTGUIDLabel(GUID, UniqueID)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetWMS_CT_GUID_LabelListByALL() As List(Of clsCTGUIDLabel)
    Try
      Dim _lstReturn As New List(Of clsCTGUIDLabel)
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
            Dim Info As clsCTGUIDLabel = Nothing
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
  Public Shared Function GetdicGuid_LabelByUniqueID(ByVal UniqueID As String) As Dictionary(Of String, clsCTGUIDLabel)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsCTGUIDLabel)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strUniqueIDList As String = ""

        If strUniqueIDList = "" Then
          strUniqueIDList = "'" & UniqueID & "'"
        Else
          strUniqueIDList = strUniqueIDList & ",'" & UniqueID & "'"
        End If
        If strWhere = "" Then
          strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.UniqueID.ToString, strUniqueIDList)
        Else
          strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.UniqueID.ToString, strUniqueIDList)
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
            Dim Info As clsCTGUIDLabel = Nothing
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
