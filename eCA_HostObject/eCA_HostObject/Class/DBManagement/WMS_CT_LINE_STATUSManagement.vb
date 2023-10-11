Public Class WMS_CT_LINE_STATUSManagement
  Public Shared TableName As String = "WMS_CT_LINE_STATUS"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    STATUS
    UPDATE_TIME
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsLine_Status) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12}) values ('{3}','{5}','{7}','{9}',{11},'{13}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
      IdxColumnName.STATUS.ToString, CInt(Info.Status),
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time
     )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsLine_Status) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
      IdxColumnName.STATUS.ToString, CInt(Info.Status),
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsLine_Status) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}={11},{12}='{13}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
      IdxColumnName.STATUS.ToString, CInt(Info.Status),
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '- GET
  Public Shared Function GetWMS_CT_LINE_STATUSDataListByALL() As List(Of clsLine_Status)
    Try
      Dim _lstReturn As New List(Of clsLine_Status)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsLine_Status = Nothing
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

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsLine_Status, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim Factory_No = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim Area_No = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim Device_No = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim Unit_ID = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
        Dim Status = If(IsNumeric(RowData.Item(IdxColumnName.STATUS.ToString)), RowData.Item(IdxColumnName.STATUS.ToString), 0 & RowData.Item(IdxColumnName.STATUS.ToString))
        Dim Update_Time = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Info = New clsLine_Status(Factory_No, Area_No, Device_No, Unit_ID, Status, Update_Time)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
