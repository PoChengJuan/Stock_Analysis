
Partial Class WMS_T_ALARMManagement
  Public Shared TableName As String = "WMS_T_ALARM"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    OCCUR_TIME
    ALARM_CODE
    ALARM_TYPE
    CMD_ID
    SEND_STATUS
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsALARM) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18}) values ('{3}','{5}','{7}','{9}','{11}','{13}',{15},'{17}',{19})",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
      IdxColumnName.DEVICE_NO.ToString, CI.DEVICE_NO,
      IdxColumnName.UNIT_ID.ToString, CI.UNIT_ID,
      IdxColumnName.OCCUR_TIME.ToString, CI.OCCUR_TIME,
      IdxColumnName.ALARM_CODE.ToString, CI.ALARM_CODE,
      IdxColumnName.ALARM_TYPE.ToString, CInt(CI.ALARM_TYPE),
      IdxColumnName.CMD_ID.ToString, CI.CMD_ID,
      IdxColumnName.SEND_STATUS.ToString, CInt(CI.SEND_STATUS)
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
  Public Shared Function GetDeleteSQL(ByRef CI As clsALARM) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' AND {10}='{11}' ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
      IdxColumnName.DEVICE_NO.ToString, CI.DEVICE_NO,
      IdxColumnName.UNIT_ID.ToString, CI.UNIT_ID,
      IdxColumnName.ALARM_CODE.ToString, CI.ALARM_CODE
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
  Public Shared Function GetUpdateSQL(ByRef CI As clsALARM) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {2}='{3}',{4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}={15},{16}='{17}',{18}={19} WHERE ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, CI.FACTORY_NO,
      IdxColumnName.AREA_NO.ToString, CI.AREA_NO,
      IdxColumnName.DEVICE_NO.ToString, CI.DEVICE_NO,
      IdxColumnName.UNIT_ID.ToString, CI.UNIT_ID,
      IdxColumnName.OCCUR_TIME.ToString, CI.OCCUR_TIME,
      IdxColumnName.ALARM_CODE.ToString, CI.ALARM_CODE,
      IdxColumnName.ALARM_TYPE.ToString, CInt(CI.ALARM_TYPE),
      IdxColumnName.CMD_ID.ToString, CI.CMD_ID,
      IdxColumnName.SEND_STATUS.ToString, CInt(CI.SEND_STATUS)
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

  Public Shared Function GetWMS_T_ALARMDataListByALL() As List(Of clsALARM)
    Try
      Dim _lstReturn As New List(Of clsALARM)
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
            Dim Info As clsALARM = Nothing
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsALARM, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim FACTORY_NO = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim AREA_NO = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim DEVICE_NO = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim UNIT_ID = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
        Dim OCCUR_TIME = "" & RowData.Item(IdxColumnName.OCCUR_TIME.ToString)
        Dim ALARM_CODE = "" & RowData.Item(IdxColumnName.ALARM_CODE.ToString)
        Dim ALARM_TYPE = 0 & RowData.Item(IdxColumnName.ALARM_TYPE.ToString)
        Dim CMD_ID = "" & RowData.Item(IdxColumnName.CMD_ID.ToString)
        Dim SEND_STATUS = 0 & RowData.Item(IdxColumnName.SEND_STATUS.ToString)
        Info = New clsALARM(FACTORY_NO, AREA_NO, DEVICE_NO, UNIT_ID, OCCUR_TIME, ALARM_CODE, ALARM_TYPE, CMD_ID, SEND_STATUS)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
