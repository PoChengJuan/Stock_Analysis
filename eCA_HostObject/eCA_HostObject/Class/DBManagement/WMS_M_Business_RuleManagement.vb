
Partial Class WMS_M_Business_RuleManagement
  Public Shared TableName As String = "WMS_M_Business_Rule"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    RULE_NO
    RULE_NAME
    RULE_TYPE1
    RULE_TYPE2
    RULE_VALUE
    ENABLE
    UPDATE_TIME
    RULE_DESC
    USER_SET_ENABLE
  End Enum

  '- GetSQL
  '-請將 clsBusiness_Rule 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsBusiness_Rule) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{15},{16}) values ({3},'{5}','{7}','{9}','{11}',{13},'{15}','{17}',{19})",
      strSQL,
      TableName,
      IdxColumnName.RULE_NO.ToString, Info.Rule_No,
      IdxColumnName.RULE_NAME.ToString, Info.Rule_Name,
      IdxColumnName.RULE_TYPE1.ToString, Info.Rule_Type1,
      IdxColumnName.RULE_TYPE2.ToString, Info.Rule_Type2,
      IdxColumnName.RULE_VALUE.ToString, Info.Rule_Value,
      IdxColumnName.ENABLE.ToString, BooleanConvertToInteger(Info.Enable),
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time,
      IdxColumnName.RULE_DESC.ToString, Info.Rule_Desc,
      IdxColumnName.USER_SET_ENABLE.ToString, BooleanConvertToInteger(Info.User_Set_Enable)
)
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsBusiness_Rule) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}={3} ",
      strSQL,
      TableName,
      IdxColumnName.RULE_NO.ToString, Info.Rule_No
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsBusiness_Rule) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}={13},{14}='{15}',{16}='{17}',{18}={19} WHERE {2}={3}",
      strSQL,
      TableName,
      IdxColumnName.RULE_NO.ToString, Info.Rule_No,
      IdxColumnName.RULE_NAME.ToString, Info.Rule_Name,
      IdxColumnName.RULE_TYPE1.ToString, Info.Rule_Type1,
      IdxColumnName.RULE_TYPE2.ToString, Info.Rule_Type2,
      IdxColumnName.RULE_VALUE.ToString, Info.Rule_Value,
      IdxColumnName.ENABLE.ToString, BooleanConvertToInteger(Info.Enable),
      IdxColumnName.UPDATE_TIME.ToString, Info.Update_Time,
      IdxColumnName.RULE_DESC.ToString, Info.Rule_Desc,
      IdxColumnName.USER_SET_ENABLE.ToString, BooleanConvertToInteger(Info.User_Set_Enable)
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- GET
  Public Shared Function GetWMS_M_Business_RuleDataListByALL() As List(Of clsBusiness_Rule)
    Try
      Dim _lstReturn As New List(Of clsBusiness_Rule)
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
            Dim Info As clsBusiness_Rule = Nothing
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
  Public Shared Function GetclsBusiness_RuleListByRULE_NO(ByVal rule_no As Double) As List(Of clsBusiness_Rule)
    Try
      Dim _lstReturn As New List(Of clsBusiness_Rule)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE  {2} = {3} ",
          strSQL,
          TableName,
          IdxColumnName.RULE_NO.ToString, rule_no
          )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsBusiness_Rule = Nothing
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsBusiness_Rule, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim RULE_NO = 0 & RowData.Item(IdxColumnName.RULE_NO.ToString)
        Dim RULE_NAME = "" & RowData.Item(IdxColumnName.RULE_NAME.ToString)
        Dim RULE_TYPE1 = "" & RowData.Item(IdxColumnName.RULE_TYPE1.ToString)
        Dim RULE_TYPE2 = "" & RowData.Item(IdxColumnName.RULE_TYPE2.ToString)
        Dim RULE_VALUE = "" & RowData.Item(IdxColumnName.RULE_VALUE.ToString)
        Dim ENABLE = IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.ENABLE.ToString))
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim RULE_DESC = "" & RowData.Item(IdxColumnName.RULE_DESC.ToString)
        Dim USER_SET_ENABLE = IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.USER_SET_ENABLE.ToString))
        Info = New clsBusiness_Rule(RULE_NO, RULE_NAME, RULE_TYPE1, RULE_TYPE2, RULE_VALUE, ENABLE, UPDATE_TIME, RULE_DESC, USER_SET_ENABLE)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
