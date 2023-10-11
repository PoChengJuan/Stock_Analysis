Public Class WMS_CH_LINE_PRODUCTION_HISTManagement
  Public Shared TableName As String = "WMS_CH_LINE_PRODUCTION_HIST"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    DEVICE_NO
    UNIT_ID
    QTY_PROCESS
    QTY_MODIFY
    QTY_NG
    HIST_TIME
    QTY_TOTAL
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsLineProduction_Hist) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18}) values ('{3}','{5}','{7}','{9}',{11},{13},{15},'{17}',{19})",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.UNIT_ID.ToString, Info.Unit_ID,
      IdxColumnName.QTY_PROCESS.ToString, Info.Qty_Process,
      IdxColumnName.QTY_MODIFY.ToString, Info.Qty_Modify,
      IdxColumnName.QTY_NG.ToString, Info.Qty_NG,
      IdxColumnName.HIST_TIME.ToString, Info.Hist_Time,
      IdxColumnName.QTY_TOTAL.ToString, Info.QTY_TOTAL
     )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- Select
  Public Shared Function GetclsLineProductionHISTByAll() As List(Of clsLineProduction_Hist)
    Try
      Dim _lstReturn As New List(Of clsLineProduction_Hist)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4} = '{5}' AND {6} = '{7}' AND {8} = '{9}' ",
          strSQL,
          TableName,
          IdxColumnName.FACTORY_NO.ToString, "",
          IdxColumnName.AREA_NO.ToString, "",
          IdxColumnName.DEVICE_NO.ToString, "",
          IdxColumnName.UNIT_ID.ToString, "",
          IdxColumnName.QTY_PROCESS.ToString, "",
          IdxColumnName.QTY_MODIFY.ToString, "",
          IdxColumnName.QTY_NG.ToString, "",
          IdxColumnName.HIST_TIME.ToString, "",
          IdxColumnName.QTY_TOTAL.ToString, ""
          )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsLineProduction_Hist = Nothing
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
  '- Select
  Public Shared Function GetclsLineProductionHISTByHistTime(ByVal StartTime As String, ByVal EndTime As String) As List(Of clsLineProduction_Hist)
    Try
      Dim _lstReturn As New List(Of clsLineProduction_Hist)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE ",
          strSQL,
          TableName
          )
        strSQL += IdxColumnName.HIST_TIME.ToString & "  BETWEEN '" & StartTime & "' and '" & EndTime & "'"

        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsLineProduction_Hist = Nothing
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsLineProduction_Hist, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim Factory_No As String = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim Area_No As String = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim Device_No As String = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim Unit_ID As String = "" & RowData.Item(IdxColumnName.UNIT_ID.ToString)
        Dim Qty_Process As Double = "" & RowData.Item(IdxColumnName.QTY_PROCESS.ToString)
        Dim Qty_Modify As Double = "" & RowData.Item(IdxColumnName.QTY_MODIFY.ToString)
        Dim Qty_NG As Double = "" & RowData.Item(IdxColumnName.QTY_NG.ToString)
        Dim Hist_Time As String = "" & RowData.Item(IdxColumnName.HIST_TIME.ToString)
        Dim QTY_TOTAL As Double = "" & RowData.Item(IdxColumnName.QTY_TOTAL.ToString)

        Info = New clsLineProduction_Hist(Factory_No, Area_No, Device_No, Unit_ID, Qty_Process, Qty_Modify, Qty_NG, Hist_Time, QTY_TOTAL)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Class
