Partial Class WMS_CM_CLASS_ASSIGNATIONManagement
  Public Shared TableName As String = "WMS_CM_CLASS_ASSIGNATION"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    CLASS_NO
    ASSIGNATION_RATE
    UPDATE_USER
    UPDATE_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsCLASS_ASSIGNATION) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12}) values ('{3}','{5}','{7}',{9},'{11}','{13}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
      IdxColumnName.ASSIGNATION_RATE.ToString, Info.ASSIGNATION_RATE,
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsCLASS_ASSIGNATION) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}={9},{10}='{11}',{12}='{13}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
      IdxColumnName.ASSIGNATION_RATE.ToString, Info.ASSIGNATION_RATE,
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsCLASS_ASSIGNATION) As Integer
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, Info.AREA_NO,
      IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
      IdxColumnName.ASSIGNATION_RATE.ToString, Info.ASSIGNATION_RATE,
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
  Public Shared Function GetWMS_CM_ClassAssignationDataListByALL() As List(Of clsCLASS_ASSIGNATION)
    Try
      Dim _lstReturn As New List(Of clsCLASS_ASSIGNATION)
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
            Dim Info As clsCLASS_ASSIGNATION = Nothing
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

  Private Shared Function SetInfoFromDB(ByRef Info As clsCLASS_ASSIGNATION, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim FACTORY_NO = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim AREA_NO = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim CLASS_NO = "" & RowData.Item(IdxColumnName.CLASS_NO.ToString)
        Dim ASSIGNATION_RATE = If(IsNumeric(RowData.Item(IdxColumnName.ASSIGNATION_RATE.ToString)), RowData.Item(IdxColumnName.ASSIGNATION_RATE.ToString), 0 & RowData.Item(IdxColumnName.ASSIGNATION_RATE.ToString))
        Dim UPDATE_USER = "" & RowData.Item(IdxColumnName.UPDATE_USER.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Info = New clsCLASS_ASSIGNATION(FACTORY_NO, AREA_NO, CLASS_NO, ASSIGNATION_RATE, UPDATE_USER, UPDATE_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function




End Class
