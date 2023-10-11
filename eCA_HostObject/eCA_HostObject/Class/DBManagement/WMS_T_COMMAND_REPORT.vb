Imports System.Collections.Concurrent


Partial Class WMS_T_COMMAND_REPORT
  Public Shared TableName As String = "WMS_T_COMMAND_REPORT"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    UUID
    REPORT_SYSTEM_TYPE
    REPORT_SYSTEM_UUID
    CREATE_TIME
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsCommandReport) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8}) values ('{3}','{5}','{7}','{9}')",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, Info.UUID,
      IdxColumnName.REPORT_SYSTEM_TYPE.ToString, CInt(Info.REPORT_SYSTEM_TYPE),
      IdxColumnName.REPORT_SYSTEM_UUID.ToString, Info.REPORT_SYSTEM_UUID,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsCommandReport) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' ",
      strSQL,
      TableName,
      IdxColumnName.UUID.ToString, Info.UUID,
      IdxColumnName.REPORT_SYSTEM_TYPE.ToString, CInt(Info.REPORT_SYSTEM_TYPE),
      IdxColumnName.REPORT_SYSTEM_UUID.ToString, Info.REPORT_SYSTEM_UUID
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
  Public Shared Function GetWMS_T_COMMAND_REPORTDataListByALL() As List(Of clsCommandReport)
    Try
      Dim _lstReturn As New List(Of clsCommandReport)
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
            Dim Info As clsCommandReport = Nothing
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsCommandReport, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim UUID = "" & RowData.Item(IdxColumnName.UUID.ToString)
        Dim REPORT_SYSTEM_TYPE = IIf(IsNumeric(RowData.Item(IdxColumnName.REPORT_SYSTEM_TYPE.ToString)), RowData.Item(IdxColumnName.REPORT_SYSTEM_TYPE.ToString), 0)
        Dim REPORT_SYSTEM_UUID = "" & RowData.Item(IdxColumnName.REPORT_SYSTEM_UUID.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Info = New clsCommandReport(UUID, REPORT_SYSTEM_TYPE, REPORT_SYSTEM_UUID, CREATE_TIME)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
