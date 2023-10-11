Partial Class WMS_M_RETURN_SUPPLIER_SETTINGManagement
  Public Shared TableName As String = "WMS_M_RETURN_SUPPLIER_SETTING"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    LOCATION_NO
    SUPPLIER_NO
    HIGH_WATER
    LOW_WATER
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsRETURNSUPPLIERSETTING) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8}) values ('{3}','{5}',{7},{9})",
 strSQL,
 TableName,
 IdxColumnName.LOCATION_NO.ToString, Info.LOCATION_NO,
 IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
 IdxColumnName.HIGH_WATER.ToString, Info.HIGH_WATER,
 IdxColumnName.LOW_WATER.ToString, Info.LOW_WATER
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsRETURNSUPPLIERSETTING) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}={7},{8}={9} WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.LOCATION_NO.ToString, Info.LOCATION_NO,
      IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
      IdxColumnName.HIGH_WATER.ToString, Info.HIGH_WATER,
      IdxColumnName.LOW_WATER.ToString, Info.LOW_WATER
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsRETURNSUPPLIERSETTING) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.LOCATION_NO.ToString, Info.LOCATION_NO,
      IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
      IdxColumnName.HIGH_WATER.ToString, Info.HIGH_WATER,
      IdxColumnName.LOW_WATER.ToString, Info.LOW_WATER
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsRETURNSUPPLIERSETTING, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim LOCATION_NO = "" & RowData.Item(IdxColumnName.LOCATION_NO.ToString)
        Dim SUPPLIER_NO = "" & RowData.Item(IdxColumnName.SUPPLIER_NO.ToString)
        Dim HIGH_WATER = If(IsNumeric(RowData.Item(IdxColumnName.HIGH_WATER.ToString)), RowData.Item(IdxColumnName.HIGH_WATER.ToString), 0 & RowData.Item(IdxColumnName.HIGH_WATER.ToString))
        Dim LOW_WATER = If(IsNumeric(RowData.Item(IdxColumnName.LOW_WATER.ToString)), RowData.Item(IdxColumnName.LOW_WATER.ToString), 0 & RowData.Item(IdxColumnName.LOW_WATER.ToString))
        Info = New clsRETURNSUPPLIERSETTING(LOCATION_NO, SUPPLIER_NO, HIGH_WATER, LOW_WATER)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetdicWMS_M_RETURN_SUPPLIER_SETTINGByALL() As Dictionary(Of String, clsRETURNSUPPLIERSETTING)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsRETURNSUPPLIERSETTING)
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
            Dim Info As clsRETURNSUPPLIERSETTING = Nothing
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
