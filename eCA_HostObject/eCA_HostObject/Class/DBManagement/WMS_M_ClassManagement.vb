
Partial Class WMS_M_ClassManagement
  Public Shared TableName As String = "WMS_M_Class"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    CLASS_NO
    CLASS_ID
    CLASS_ALIS
    CLASS_DESC
    CLASS_MANAGER
    PHONE
    CLASS_START_TIME
    CLASS_END_TIME
  End Enum

  '- GetSQL
  '-請將 clsClass 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsClass) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}')",
 strSQL,
 TableName,
 IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
 IdxColumnName.CLASS_ID.ToString, Info.CLASS_ID,
 IdxColumnName.CLASS_ALIS.ToString, Info.CLASS_ALIS,
 IdxColumnName.CLASS_DESC.ToString, Info.CLASS_DESC,
 IdxColumnName.CLASS_MANAGER.ToString, Info.CLASS_MANAGER,
 IdxColumnName.PHONE.ToString, Info.PHONE,
 IdxColumnName.CLASS_START_TIME.ToString, Info.CLASS_START_TIME,
 IdxColumnName.CLASS_END_TIME.ToString, Info.CLASS_END_TIME
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsClass) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
 strSQL,
 TableName,
 IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
 IdxColumnName.CLASS_ID.ToString, Info.CLASS_ID,
 IdxColumnName.CLASS_ALIS.ToString, Info.CLASS_ALIS,
 IdxColumnName.CLASS_DESC.ToString, Info.CLASS_DESC,
 IdxColumnName.CLASS_MANAGER.ToString, Info.CLASS_MANAGER,
 IdxColumnName.PHONE.ToString, Info.PHONE,
 IdxColumnName.CLASS_START_TIME.ToString, Info.CLASS_START_TIME,
 IdxColumnName.CLASS_END_TIME.ToString, Info.CLASS_END_TIME
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsClass) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}' WHERE {2}='{3}'",
 strSQL,
 TableName,
 IdxColumnName.CLASS_NO.ToString, Info.CLASS_NO,
 IdxColumnName.CLASS_ID.ToString, Info.CLASS_ID,
 IdxColumnName.CLASS_ALIS.ToString, Info.CLASS_ALIS,
 IdxColumnName.CLASS_DESC.ToString, Info.CLASS_DESC,
 IdxColumnName.CLASS_MANAGER.ToString, Info.CLASS_MANAGER,
 IdxColumnName.PHONE.ToString, Info.PHONE,
 IdxColumnName.CLASS_START_TIME.ToString, Info.CLASS_START_TIME,
 IdxColumnName.CLASS_END_TIME.ToString, Info.CLASS_END_TIME
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
  Public Shared Function GetWMS_M_ClassDataListByALL() As List(Of clsClass)
    Try
      Dim _lstReturn As New List(Of clsClass)
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
            Dim Info As clsClass = Nothing
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
  Public Shared Function GetclsClassListByCLASS_NO(ByVal class_no As String) As List(Of clsClass)
    Try
      Dim _lstReturn As New List(Of clsClass)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
strSQL,
TableName,
IdxColumnName.CLASS_NO.ToString, class_no
)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsClass = Nothing
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsClass, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim CLASS_NO = "" & RowData.Item(IdxColumnName.CLASS_NO.ToString)
        Dim CLASS_ID = "" & RowData.Item(IdxColumnName.CLASS_ID.ToString)
        Dim CLASS_ALIS = "" & RowData.Item(IdxColumnName.CLASS_ALIS.ToString)
        Dim CLASS_DESC = "" & RowData.Item(IdxColumnName.CLASS_DESC.ToString)
        Dim CLASS_MANAGER = "" & RowData.Item(IdxColumnName.CLASS_MANAGER.ToString)
        Dim PHONE = "" & RowData.Item(IdxColumnName.PHONE.ToString)
        Dim CLASS_START_TIME = "" & RowData.Item(IdxColumnName.CLASS_START_TIME.ToString)
        Dim CLASS_END_TIME = "" & RowData.Item(IdxColumnName.CLASS_END_TIME.ToString)
        Info = New clsClass(CLASS_NO, CLASS_ID, CLASS_ALIS, CLASS_DESC, CLASS_MANAGER, PHONE, CLASS_START_TIME, CLASS_END_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
