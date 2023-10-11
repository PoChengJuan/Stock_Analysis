Partial Class WMS_CM_CLASS_ATTENDANCEManagement
  Public Shared TableName As String = "WMS_CM_CLASS_ATTENDANCE"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    CLASS_NO
    ATTENDANCE_COUNT
    UPDATE_USER
    UPDATE_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsCLASS_ATTENDANCE) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8}) values ('{3}',{5},'{7}','{9}')",
      strSQL,
      TableName,
      IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
      IdxColumnName.ATTENDANCE_COUNT.ToString, Info.ATTENDANCE_COUNT,
      IdxColumnName.UPDATE_USER.ToString, Info.UPDATE_USER,
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsCLASS_ATTENDANCE) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}={5},{6}='{7}',{8}='{9}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
      IdxColumnName.ATTENDANCE_COUNT.ToString, Info.ATTENDANCE_COUNT,
      IdxColumnName.UPDATE_USER.ToString, Info.UPDATE_USER,
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsCLASS_ATTENDANCE) As Integer
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
      IdxColumnName.ATTENDANCE_COUNT.ToString, Info.ATTENDANCE_COUNT,
      IdxColumnName.UPDATE_USER.ToString, Info.UPDATE_USER,
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
  '- GET
  Public Shared Function GetWMS_CM_ClassAttendanceDataListByALL() As List(Of clsCLASS_ATTENDANCE)
    Try
      Dim _lstReturn As New List(Of clsCLASS_ATTENDANCE)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCLASS_ATTENDANCE = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info)
          Next
        End If
        'End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  Private Shared Function SetInfoFromDB(ByRef Info As clsCLASS_ATTENDANCE, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim CLASS_NO = "" & RowData.Item(IdxColumnName.CLASS_NO.ToString)
        Dim ATTENDANCE_COUNT = If(IsNumeric(RowData.Item(IdxColumnName.ATTENDANCE_COUNT.ToString)), RowData.Item(IdxColumnName.ATTENDANCE_COUNT.ToString), 0 & RowData.Item(IdxColumnName.ATTENDANCE_COUNT.ToString))
        Dim UPDATE_USER = "" & RowData.Item(IdxColumnName.UPDATE_USER.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Info = New clsCLASS_ATTENDANCE(CLASS_NO, ATTENDANCE_COUNT, UPDATE_USER, UPDATE_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
