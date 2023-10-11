Partial Class WMS_M_SKU_Packe_StructureManagement
  Public Shared TableName As String = "WMS_M_SKU_Packe_Structure"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    SKU_NO
    PACKE_LV
    PACKE_UNIT
    SUB_PACKE_UNIT
    PACKE_WEIGHT
    PACKE_VOLUME
    PACKE_BCR
    OUT_MAX_UNIT
    IN_MAX_UNIT
    QTY
    COMMENTS
    CREATE_TIME
    UPDATE_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsMSKUPackeStructure) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26}) values ('{3}',{5},'{7}','{9}',{11},{13},'{15}',{17},{19},{21},'{23}','{25}','{27}')",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.PACKE_LV.ToString, Info.PACKE_LV,
      IdxColumnName.PACKE_UNIT.ToString, Info.PACKE_UNIT,
      IdxColumnName.SUB_PACKE_UNIT.ToString, Info.SUB_PACKE_UNIT,
      IdxColumnName.PACKE_WEIGHT.ToString, Info.PACKE_WEIGHT,
      IdxColumnName.PACKE_VOLUME.ToString, Info.PACKE_VOLUME,
      IdxColumnName.PACKE_BCR.ToString, Info.PACKE_BCR,
      IdxColumnName.OUT_MAX_UNIT.ToString, Info.OUT_MAX_UNIT,
      IdxColumnName.IN_MAX_UNIT.ToString, Info.IN_MAX_UNIT,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsMSKUPackeStructure) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}={11},{12}={13},{14}='{15}',{16}={17},{18}={19},{20}={21},{22}='{23}',{24}='{25}',{26}='{27}' WHERE {2}='{3}' And {4}={5} And {6}='{7}' And {8}='{9}'",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.PACKE_LV.ToString, Info.PACKE_LV,
      IdxColumnName.PACKE_UNIT.ToString, Info.PACKE_UNIT,
      IdxColumnName.SUB_PACKE_UNIT.ToString, Info.SUB_PACKE_UNIT,
      IdxColumnName.PACKE_WEIGHT.ToString, Info.PACKE_WEIGHT,
      IdxColumnName.PACKE_VOLUME.ToString, Info.PACKE_VOLUME,
      IdxColumnName.PACKE_BCR.ToString, Info.PACKE_BCR,
      IdxColumnName.OUT_MAX_UNIT.ToString, Info.OUT_MAX_UNIT,
      IdxColumnName.IN_MAX_UNIT.ToString, Info.IN_MAX_UNIT,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsMSKUPackeStructure) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}={5} AND {6}='{7}' AND {8}='{9}' ",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.PACKE_LV.ToString, Info.PACKE_LV,
      IdxColumnName.PACKE_UNIT.ToString, Info.PACKE_UNIT,
      IdxColumnName.SUB_PACKE_UNIT.ToString, Info.SUB_PACKE_UNIT,
      IdxColumnName.PACKE_WEIGHT.ToString, Info.PACKE_WEIGHT,
      IdxColumnName.PACKE_VOLUME.ToString, Info.PACKE_VOLUME,
      IdxColumnName.PACKE_BCR.ToString, Info.PACKE_BCR,
      IdxColumnName.OUT_MAX_UNIT.ToString, Info.OUT_MAX_UNIT,
      IdxColumnName.IN_MAX_UNIT.ToString, Info.IN_MAX_UNIT,
      IdxColumnName.QTY.ToString, Info.QTY,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsMSKUPackeStructure, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim PACKE_LV = If(IsNumeric(RowData.Item(IdxColumnName.PACKE_LV.ToString)), RowData.Item(IdxColumnName.PACKE_LV.ToString), 0 & RowData.Item(IdxColumnName.PACKE_LV.ToString))
        Dim PACKE_UNIT = "" & RowData.Item(IdxColumnName.PACKE_UNIT.ToString)
        Dim SUB_PACKE_UNIT = "" & RowData.Item(IdxColumnName.SUB_PACKE_UNIT.ToString)
        Dim PACKE_WEIGHT = If(IsNumeric(RowData.Item(IdxColumnName.PACKE_WEIGHT.ToString)), RowData.Item(IdxColumnName.PACKE_WEIGHT.ToString), 0 & RowData.Item(IdxColumnName.PACKE_WEIGHT.ToString))
        Dim PACKE_VOLUME = If(IsNumeric(RowData.Item(IdxColumnName.PACKE_VOLUME.ToString)), RowData.Item(IdxColumnName.PACKE_VOLUME.ToString), 0 & RowData.Item(IdxColumnName.PACKE_VOLUME.ToString))
        Dim PACKE_BCR = "" & RowData.Item(IdxColumnName.PACKE_BCR.ToString)
        Dim OUT_MAX_UNIT = If(IsNumeric(RowData.Item(IdxColumnName.OUT_MAX_UNIT.ToString)), RowData.Item(IdxColumnName.OUT_MAX_UNIT.ToString), 0 & RowData.Item(IdxColumnName.OUT_MAX_UNIT.ToString))
        Dim IN_MAX_UNIT = If(IsNumeric(RowData.Item(IdxColumnName.IN_MAX_UNIT.ToString)), RowData.Item(IdxColumnName.IN_MAX_UNIT.ToString), 0 & RowData.Item(IdxColumnName.IN_MAX_UNIT.ToString))
        Dim QTY = If(IsNumeric(RowData.Item(IdxColumnName.QTY.ToString)), RowData.Item(IdxColumnName.QTY.ToString), 0 & RowData.Item(IdxColumnName.QTY.ToString))
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Info = New clsMSKUPackeStructure(SKU_NO, PACKE_LV, PACKE_UNIT, SUB_PACKE_UNIT, PACKE_WEIGHT, PACKE_VOLUME, PACKE_BCR, OUT_MAX_UNIT, IN_MAX_UNIT, QTY, COMMENTS, CREATE_TIME, UPDATE_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetWMS_M_SKU_Packe_StructureListByALL() As List(Of clsMSKUPackeStructure)
    Try
      Dim _lstReturn As New List(Of clsMSKUPackeStructure)
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
            Dim Info As clsMSKUPackeStructure = Nothing
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
  Public Shared Function GetdicSKUPackeStructureBySKU(ByVal SKU_NO As String) As Dictionary(Of String, clsMSKUPackeStructure)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsMSKUPackeStructure)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strUniqueIDList As String = ""


        If strWhere = "" Then
          strWhere = String.Format("WHERE {0} IN ('{1}') ", IdxColumnName.SKU_NO.ToString, SKU_NO)
        Else
          strWhere = String.Format("{0} AND {1} = ('{2}') ", strWhere, IdxColumnName.SKU_NO.ToString, SKU_NO)
        End If
        strSQL = String.Format("Select * from {1} {2} ",
            strSQL,
            TableName,
            strWhere
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsMSKUPackeStructure = Nothing
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
  Public Shared Function GetdicSKUPackeStructrueBySKUNo(ByVal SKUNo As String) As Dictionary(Of String, clsMSKUPackeStructure)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsMSKUPackeStructure)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strSKUNoList As String = ""

        If strSKUNoList = "" Then
          strSKUNoList = "'" & SKUNo & "'"
        Else
          strSKUNoList = strSKUNoList & ",'" & SKUNo & "'"
        End If
        If strWhere = "" Then
          strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.SKU_NO.ToString, strSKUNoList)
        Else
          strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.SKU_NO.ToString, strSKUNoList)
        End If
        strSQL = String.Format("Select * from {1} {2} ",
            strSQL,
            TableName,
            strWhere
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsMSKUPackeStructure = Nothing
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
