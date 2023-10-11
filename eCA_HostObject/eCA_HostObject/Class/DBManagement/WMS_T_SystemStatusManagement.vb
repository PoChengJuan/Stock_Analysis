
Partial Class WMS_T_SystemStatusManagement
  Public Shared TableName As String = "WMS_T_SYSTEM_STATUS"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    STATUS_NO
    STATUS_NAME
    STATUS_VALUE
    UPDATE_TIME
    STATUS_MODE
    STATUS_TYPE1
    STATUS_TYPE2
    STATUS_TYPE3
    STATUS_DESC
  End Enum

  '- GetSQL
  '-請將 clsSystemStatus 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsSystemStatus) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18}) values ('{3}','{5}','{7}','{9}',{11},'{13}','{15}','{17}','{19}')",
      strSQL,
      TableName,
      IdxColumnName.STATUS_NO.ToString, CInt(Info.STATUS_NO),
      IdxColumnName.STATUS_NAME.ToString, Info.STATUS_NAME,
      IdxColumnName.STATUS_VALUE.ToString, Info.STATUS_VALUE,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.STATUS_MODE.ToString, CInt(Info.STATUS_MODE),
      IdxColumnName.STATUS_TYPE1.ToString, Info.STATUS_TYPE1,
      IdxColumnName.STATUS_TYPE2.ToString, Info.STATUS_TYPE2,
      IdxColumnName.STATUS_TYPE3.ToString, Info.STATUS_TYPE3,
      IdxColumnName.STATUS_DESC.ToString, Info.STATUS_DESC
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsSystemStatus) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.STATUS_NO.ToString, Info.STATUS_NO
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsSystemStatus) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}={11},{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.STATUS_NO.ToString, CInt(Info.STATUS_NO),
      IdxColumnName.STATUS_NAME.ToString, Info.STATUS_NAME,
      IdxColumnName.STATUS_VALUE.ToString, Info.STATUS_VALUE,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.STATUS_MODE.ToString, CInt(Info.STATUS_MODE),
      IdxColumnName.STATUS_TYPE1.ToString, Info.STATUS_TYPE1,
      IdxColumnName.STATUS_TYPE2.ToString, Info.STATUS_TYPE2,
      IdxColumnName.STATUS_TYPE3.ToString, Info.STATUS_TYPE3,
      IdxColumnName.STATUS_DESC.ToString, Info.STATUS_DESC
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- GET
  Public Shared Function GetWMS_T_System_StatusDataListByALL() As List(Of clsSystemStatus)
    Try
      Dim _lstReturn As New List(Of clsSystemStatus)
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
            Dim Info As clsSystemStatus = Nothing
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

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsSystemStatus, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim STATUS_NO = 0 & RowData.Item(IdxColumnName.STATUS_NO.ToString)
        Dim STATUS_NAME = "" & RowData.Item(IdxColumnName.STATUS_NAME.ToString)
        Dim STATUS_VALUE = "" & RowData.Item(IdxColumnName.STATUS_VALUE.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)

        Dim STATUS_MODE = 0 & RowData.Item(IdxColumnName.STATUS_MODE.ToString)
        Dim STATUS_TYPE1 = "" & RowData.Item(IdxColumnName.STATUS_TYPE1.ToString)
        Dim STATUS_TYPE2 = "" & RowData.Item(IdxColumnName.STATUS_TYPE2.ToString)
        Dim STATUS_TYPE3 = "" & RowData.Item(IdxColumnName.STATUS_TYPE3.ToString)
        Dim STATUS_DESC = "" & RowData.Item(IdxColumnName.STATUS_DESC.ToString)

        Info = New clsSystemStatus(STATUS_NO, STATUS_NAME, STATUS_VALUE, UPDATE_TIME, STATUS_MODE, STATUS_TYPE1, STATUS_TYPE2, STATUS_TYPE3, STATUS_DESC)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
