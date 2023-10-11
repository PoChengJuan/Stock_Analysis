Public Class WMS_CT_PRODUCE_INFOManagement
  Public Shared TableName As String = "WMS_CT_PRODUCE_INFO"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    FACTORY_NO
    AREA_NO
    PO_ID
    SKU_NO
    STATUS
    QTY
    QTY_PROCESS
    PREVIOUS_QTY_PROCESS
    QTY_NG
    PREVIOUS_QTY_NG
    PREVIOUS_AREA_NO
    CREATE_TIME
    START_TIME
    UPDATE_TIME
    FINISH_TIME
    PO_Info1
    PO_Info2
    PO_Info3
    PO_Info4
    PO_Info5
    PO_Info6
    PO_Info7
    PO_Info8
    PO_Info9
    PO_Info10
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsProduce_Info) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50}) values ('{3}','{5}','{7}','{9}',{11},{13},{15},{17},{19},{21},'{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}')",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.STATUS.ToString, CInt(Info.Status),
      IdxColumnName.QTY.ToString, Info.Qty,
      IdxColumnName.QTY_PROCESS.ToString, Info.Qty_Process,
      IdxColumnName.PREVIOUS_QTY_PROCESS.ToString, Info.Previous_Qty_Process,
      IdxColumnName.QTY_NG.ToString, Info.Qty_NG,
      IdxColumnName.PREVIOUS_QTY_NG.ToString, Info.Previous_Qty_NG,
      IdxColumnName.PREVIOUS_AREA_NO.ToString, Info.Previous_Area_No,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.START_TIME.ToString, Info.Start_Time,
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.Finish_Time,
      IdxColumnName.PO_Info1.ToString, Info.PO_Info1,
      IdxColumnName.PO_Info2.ToString, Info.PO_Info2,
      IdxColumnName.PO_Info3.ToString, Info.PO_Info3,
      IdxColumnName.PO_Info4.ToString, Info.PO_Info4,
      IdxColumnName.PO_Info5.ToString, Info.PO_Info5,
      IdxColumnName.PO_Info6.ToString, Info.PO_Info6,
      IdxColumnName.PO_Info7.ToString, Info.PO_Info7,
      IdxColumnName.PO_Info8.ToString, Info.PO_Info8,
      IdxColumnName.PO_Info9.ToString, Info.PO_Info9,
      IdxColumnName.PO_Info10.ToString, Info.PO_Info10
     )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsProduce_Info) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.STATUS.ToString, CInt(Info.Status),
      IdxColumnName.QTY.ToString, Info.Qty,
      IdxColumnName.QTY_PROCESS.ToString, Info.Qty_Process,
      IdxColumnName.PREVIOUS_QTY_PROCESS.ToString, Info.Previous_Qty_Process,
      IdxColumnName.QTY_NG.ToString, Info.Qty_NG,
      IdxColumnName.PREVIOUS_QTY_NG.ToString, Info.Previous_Qty_NG,
      IdxColumnName.PREVIOUS_AREA_NO.ToString, Info.Previous_Area_No,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.START_TIME.ToString, Info.Start_Time,
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.Finish_Time
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsProduce_Info) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}={11},{12}={13},{14}={15},{16}={17},{18}={19},{20}={21},{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.STATUS.ToString, CInt(Info.Status),
      IdxColumnName.QTY.ToString, Info.Qty,
      IdxColumnName.QTY_PROCESS.ToString, Info.Qty_Process,
      IdxColumnName.PREVIOUS_QTY_PROCESS.ToString, Info.Previous_Qty_Process,
      IdxColumnName.QTY_NG.ToString, Info.Qty_NG,
      IdxColumnName.PREVIOUS_QTY_NG.ToString, Info.Previous_Qty_NG,
      IdxColumnName.PREVIOUS_AREA_NO.ToString, Info.Previous_Area_No,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.START_TIME.ToString, Info.Start_Time,
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.Finish_Time,
      IdxColumnName.PO_Info1.ToString, Info.PO_Info1,
      IdxColumnName.PO_Info2.ToString, Info.PO_Info2,
      IdxColumnName.PO_Info3.ToString, Info.PO_Info3,
      IdxColumnName.PO_Info4.ToString, Info.PO_Info4,
      IdxColumnName.PO_Info5.ToString, Info.PO_Info5,
      IdxColumnName.PO_Info6.ToString, Info.PO_Info6,
      IdxColumnName.PO_Info7.ToString, Info.PO_Info7,
      IdxColumnName.PO_Info8.ToString, Info.PO_Info8,
      IdxColumnName.PO_Info9.ToString, Info.PO_Info9,
      IdxColumnName.PO_Info10.ToString, Info.PO_Info10
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- GET
  Public Shared Function GetWMS_C_PRODUCE_INFODataListByALL() As List(Of clsProduce_Info)
    Try
      Dim _lstReturn As New List(Of clsProduce_Info)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsProduce_Info = Nothing
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsProduce_Info, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim Factory_No = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim Area_No = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim Status = If(IsNumeric(RowData.Item(IdxColumnName.STATUS.ToString)), RowData.Item(IdxColumnName.STATUS.ToString), 0 & RowData.Item(IdxColumnName.STATUS.ToString))
        Dim Qty = If(IsNumeric(RowData.Item(IdxColumnName.QTY.ToString)), RowData.Item(IdxColumnName.QTY.ToString), 0 & RowData.Item(IdxColumnName.QTY.ToString))
        Dim Qty_Process = If(IsNumeric(RowData.Item(IdxColumnName.QTY_PROCESS.ToString)), RowData.Item(IdxColumnName.QTY_PROCESS.ToString), 0 & RowData.Item(IdxColumnName.QTY_PROCESS.ToString))
        Dim Previous_Qty_Process = If(IsNumeric(RowData.Item(IdxColumnName.PREVIOUS_QTY_PROCESS.ToString)), RowData.Item(IdxColumnName.PREVIOUS_QTY_PROCESS.ToString), 0 & RowData.Item(IdxColumnName.PREVIOUS_QTY_PROCESS.ToString))
        Dim Qty_NG = If(IsNumeric(RowData.Item(IdxColumnName.QTY_NG.ToString)), RowData.Item(IdxColumnName.QTY_NG.ToString), 0 & RowData.Item(IdxColumnName.QTY_NG.ToString))
        Dim Previous_Qty_NG = If(IsNumeric(RowData.Item(IdxColumnName.PREVIOUS_QTY_NG.ToString)), RowData.Item(IdxColumnName.PREVIOUS_QTY_NG.ToString), 0 & RowData.Item(IdxColumnName.PREVIOUS_QTY_NG.ToString))
        Dim Previous_Area_No = "" & RowData.Item(IdxColumnName.PREVIOUS_AREA_NO.ToString)
        Dim Create_Time = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim Start_Time = "" & RowData.Item(IdxColumnName.START_TIME.ToString)
        Dim Update_Time = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim Finish_Time = "" & RowData.Item(IdxColumnName.FINISH_TIME.ToString)
        Dim PO_Info1 = "" & RowData.Item(IdxColumnName.PO_Info1.ToString)
        Dim PO_Info2 = "" & RowData.Item(IdxColumnName.PO_Info2.ToString)
        Dim PO_Info3 = "" & RowData.Item(IdxColumnName.PO_Info3.ToString)
        Dim PO_Info4 = "" & RowData.Item(IdxColumnName.PO_Info4.ToString)
        Dim PO_Info5 = "" & RowData.Item(IdxColumnName.PO_Info5.ToString)
        Dim PO_Info6 = "" & RowData.Item(IdxColumnName.PO_Info6.ToString)
        Dim PO_Info7 = "" & RowData.Item(IdxColumnName.PO_Info7.ToString)
        Dim PO_Info8 = "" & RowData.Item(IdxColumnName.PO_Info8.ToString)
        Dim PO_Info9 = "" & RowData.Item(IdxColumnName.PO_Info9.ToString)
        Dim PO_Info10 = "" & RowData.Item(IdxColumnName.PO_Info10.ToString)
        Info = New clsProduce_Info(Factory_No, Area_No, PO_ID, SKU_NO, Status, Qty, Qty_Process, Previous_Qty_Process,
                                   Qty_NG, Previous_Qty_NG, Previous_Area_No, Create_Time, Start_Time, Update_Time, Finish_Time, PO_Info1,
                                   PO_Info2, PO_Info3, PO_Info4, PO_Info5, PO_Info6, PO_Info7, PO_Info8, PO_Info9, PO_Info10)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
