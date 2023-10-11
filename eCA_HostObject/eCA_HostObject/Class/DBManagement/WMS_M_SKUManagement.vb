
Partial Class WMS_M_SKUManagement
  Public Shared TableName As String = "WMS_M_SKU"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    SKU_NO
    SKU_ID1
    SKU_ID2
    SKU_ID3
    SKU_ALIS1
    SKU_ALIS2
    SKU_DESC
    SKU_CATALOG
    SKU_TYPE1
    SKU_TYPE2
    SKU_TYPE3
    SKU_COMMON1
    SKU_COMMON2
    SKU_COMMON3
    SKU_COMMON4
    SKU_COMMON5
    SKU_COMMON6
    SKU_COMMON7
    SKU_COMMON8
    SKU_COMMON9
    SKU_COMMON10
    SKU_L
    SKU_W
    SKU_H
    SKU_WEIGHT
    SKU_VALUE
    SKU_UNIT
    INBOUND_UNIT
    OUTBOUND_UNIT
    HIGH_WATER
    LOW_WATER
    AVAILABLE_DAYS
    SAVE_DAYS
    CREATE_TIME
    UPDATE_TIME
    WEIGHT_DIFFERENCE
    ENABLE
    EFFECTIVE_DATE
    FAILURE_DATE
    QC_METHOD
    COMMENTS
  End Enum

  '- GetSQL
  '-請將 clsSKU 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsSKU) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}')",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.SKU_ID1.ToString, Info.SKU_ID1,
      IdxColumnName.SKU_ID2.ToString, Info.SKU_ID2,
      IdxColumnName.SKU_ID3.ToString, Info.SKU_ID3,
      IdxColumnName.SKU_ALIS1.ToString, Info.SKU_ALIS1,
      IdxColumnName.SKU_ALIS2.ToString, Info.SKU_ALIS2,
      IdxColumnName.SKU_DESC.ToString, Info.SKU_DESC,
      IdxColumnName.SKU_CATALOG.ToString, CInt(Info.SKU_CATALOG),
      IdxColumnName.SKU_TYPE1.ToString, Info.SKU_TYPE1,
      IdxColumnName.SKU_TYPE2.ToString, Info.SKU_TYPE2,
      IdxColumnName.SKU_TYPE3.ToString, Info.SKU_TYPE3,
      IdxColumnName.SKU_COMMON1.ToString, Info.SKU_COMMON1,
      IdxColumnName.SKU_COMMON2.ToString, Info.SKU_COMMON2,
      IdxColumnName.SKU_COMMON3.ToString, Info.SKU_COMMON3,
      IdxColumnName.SKU_COMMON4.ToString, Info.SKU_COMMON4,
      IdxColumnName.SKU_COMMON5.ToString, Info.SKU_COMMON5,
      IdxColumnName.SKU_COMMON6.ToString, Info.SKU_COMMON6,
      IdxColumnName.SKU_COMMON7.ToString, Info.SKU_COMMON7,
      IdxColumnName.SKU_COMMON8.ToString, Info.SKU_COMMON8,
      IdxColumnName.SKU_COMMON9.ToString, Info.SKU_COMMON9,
      IdxColumnName.SKU_COMMON10.ToString, Info.SKU_COMMON10,
      IdxColumnName.SKU_L.ToString, Info.SKU_L,
      IdxColumnName.SKU_W.ToString, Info.SKU_W,
      IdxColumnName.SKU_H.ToString, Info.SKU_H,
      IdxColumnName.SKU_WEIGHT.ToString, Info.SKU_WEIGHT,
      IdxColumnName.SKU_VALUE.ToString, Info.SKU_VALUE,
      IdxColumnName.SKU_UNIT.ToString, Info.SKU_UNIT,
      IdxColumnName.INBOUND_UNIT.ToString, Info.INBOUND_UNIT,
      IdxColumnName.OUTBOUND_UNIT.ToString, Info.OUTBOUND_UNIT,
      IdxColumnName.HIGH_WATER.ToString, Info.HIGH_WATER,
      IdxColumnName.LOW_WATER.ToString, Info.LOW_WATER,
      IdxColumnName.AVAILABLE_DAYS.ToString, Info.AVAILABLE_DAYS,
      IdxColumnName.SAVE_DAYS.ToString, Info.SAVE_DAYS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.WEIGHT_DIFFERENCE.ToString, Info.WEIGHT_DIFFERENCE,
      IdxColumnName.ENABLE.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.ENABLE),
      IdxColumnName.EFFECTIVE_DATE.ToString, Info.EFFECTIVE_DATE,
      IdxColumnName.FAILURE_DATE.ToString, Info.FAILURE_DATE,
      IdxColumnName.QC_METHOD.ToString, Info.QC_METHOD,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsSKU) As String
    Try

      Dim strSQL As String = ""
      'strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.SKU_ID1.ToString, Info.SKU_ID1,
      IdxColumnName.SKU_ID2.ToString, Info.SKU_ID2,
      IdxColumnName.SKU_ID3.ToString, Info.SKU_ID3
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsSKU) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
      strSQL,
      TableName,
      IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
      IdxColumnName.SKU_ID1.ToString, Info.SKU_ID1,
      IdxColumnName.SKU_ID2.ToString, Info.SKU_ID2,
      IdxColumnName.SKU_ID3.ToString, Info.SKU_ID3,
      IdxColumnName.SKU_ALIS1.ToString, Info.SKU_ALIS1,
      IdxColumnName.SKU_ALIS2.ToString, Info.SKU_ALIS2,
      IdxColumnName.SKU_DESC.ToString, Info.SKU_DESC,
      IdxColumnName.SKU_CATALOG.ToString, CInt(Info.SKU_CATALOG),
      IdxColumnName.SKU_TYPE1.ToString, Info.SKU_TYPE1,
      IdxColumnName.SKU_TYPE2.ToString, Info.SKU_TYPE2,
      IdxColumnName.SKU_TYPE3.ToString, Info.SKU_TYPE3,
      IdxColumnName.SKU_COMMON1.ToString, Info.SKU_COMMON1,
      IdxColumnName.SKU_COMMON2.ToString, Info.SKU_COMMON2,
      IdxColumnName.SKU_COMMON3.ToString, Info.SKU_COMMON3,
      IdxColumnName.SKU_COMMON4.ToString, Info.SKU_COMMON4,
      IdxColumnName.SKU_COMMON5.ToString, Info.SKU_COMMON5,
      IdxColumnName.SKU_COMMON6.ToString, Info.SKU_COMMON6,
      IdxColumnName.SKU_COMMON7.ToString, Info.SKU_COMMON7,
      IdxColumnName.SKU_COMMON8.ToString, Info.SKU_COMMON8,
      IdxColumnName.SKU_COMMON9.ToString, Info.SKU_COMMON9,
      IdxColumnName.SKU_COMMON10.ToString, Info.SKU_COMMON10,
      IdxColumnName.SKU_L.ToString, Info.SKU_L,
      IdxColumnName.SKU_W.ToString, Info.SKU_W,
      IdxColumnName.SKU_H.ToString, Info.SKU_H,
      IdxColumnName.SKU_WEIGHT.ToString, Info.SKU_WEIGHT,
      IdxColumnName.SKU_VALUE.ToString, Info.SKU_VALUE,
      IdxColumnName.SKU_UNIT.ToString, Info.SKU_UNIT,
      IdxColumnName.INBOUND_UNIT.ToString, Info.INBOUND_UNIT,
      IdxColumnName.OUTBOUND_UNIT.ToString, Info.OUTBOUND_UNIT,
      IdxColumnName.HIGH_WATER.ToString, Info.HIGH_WATER,
      IdxColumnName.LOW_WATER.ToString, Info.LOW_WATER,
      IdxColumnName.AVAILABLE_DAYS.ToString, Info.AVAILABLE_DAYS,
      IdxColumnName.SAVE_DAYS.ToString, Info.SAVE_DAYS,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.WEIGHT_DIFFERENCE.ToString, Info.WEIGHT_DIFFERENCE,
      IdxColumnName.ENABLE.ToString, ModuleHelpFunc.BooleanConvertToInteger(Info.ENABLE),
      IdxColumnName.EFFECTIVE_DATE.ToString, Info.EFFECTIVE_DATE,
      IdxColumnName.FAILURE_DATE.ToString, Info.FAILURE_DATE,
      IdxColumnName.QC_METHOD.ToString, Info.QC_METHOD,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS
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
  Public Shared Function GetWMS_M_SKUDataListByALL() As Dictionary(Of String, clsSKU)
    Try
      Dim dicReturn As New Dictionary(Of String, clsSKU)
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
            Dim Info As clsSKU = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If dicReturn.ContainsKey(Info.gid) = False Then
              dicReturn.Add(Info.gid, Info)
            End If
          Next
        End If
        'End If
      End If
      Return dicReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetclsSKUListBySKU_NO_SKU_ID1_SKU_ID2(ByVal sku_id1 As String, sku_id2 As String) As Dictionary(Of String, clsSKU)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsSKU)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4} = '{5}'  ",
          strSQL,
          TableName,
          IdxColumnName.SKU_ID1.ToString, sku_id1,
          IdxColumnName.SKU_ID2.ToString, sku_id2
          )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsSKU = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info.gid, Info)
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
  Public Shared Function GetdicSKUListBydicSKUNo(ByVal dicSKUNo As Dictionary(Of String, String)) As Dictionary(Of String, clsSKU)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsSKU)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strSKUNoList As String = ""

        Dim count_flag = 0
        For i = 0 To dicSKUNo.Count - 1
          If strSKUNoList = "" Then
            strSKUNoList = "'" & dicSKUNo.Keys(i) & "'"
          Else
            strSKUNoList = strSKUNoList & ",'" & dicSKUNo.Keys(i) & "'"
          End If
          If i - count_flag > 800 OrElse i = (dicSKUNo.Count - 1) Then
            count_flag = i
            strWhere = ""
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
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsSKU = Nothing
                SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                If _lstReturn.ContainsKey(Info.gid) = False Then
                  _lstReturn.Add(Info.gid, Info)
                End If
              Next
            End If
            strSKUNoList = ""
          End If
        Next

      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetdicSKUBySKUNo(ByVal SKUNo As String) As Dictionary(Of String, clsSKU)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsSKU)
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
            Dim Info As clsSKU = Nothing
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

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsSKU, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim SKU_ID1 = "" & RowData.Item(IdxColumnName.SKU_ID1.ToString)
        Dim SKU_ID2 = "" & RowData.Item(IdxColumnName.SKU_ID2.ToString)
        Dim SKU_ID3 = "" & RowData.Item(IdxColumnName.SKU_ID3.ToString)
        Dim SKU_ALIS1 = "" & RowData.Item(IdxColumnName.SKU_ALIS1.ToString)
        Dim SKU_ALIS2 = "" & RowData.Item(IdxColumnName.SKU_ALIS2.ToString)
        Dim SKU_DESC = "" & RowData.Item(IdxColumnName.SKU_DESC.ToString)
        Dim SKU_CATALOG = 0 & RowData.Item(IdxColumnName.SKU_CATALOG.ToString)
        Dim SKU_TYPE1 = "" & RowData.Item(IdxColumnName.SKU_TYPE1.ToString)
        Dim SKU_TYPE2 = "" & RowData.Item(IdxColumnName.SKU_TYPE2.ToString)
        Dim SKU_TYPE3 = "" & RowData.Item(IdxColumnName.SKU_TYPE3.ToString)
        Dim SKU_COMMON1 = "" & RowData.Item(IdxColumnName.SKU_COMMON1.ToString)
        Dim SKU_COMMON2 = "" & RowData.Item(IdxColumnName.SKU_COMMON2.ToString)
        Dim SKU_COMMON3 = "" & RowData.Item(IdxColumnName.SKU_COMMON3.ToString)
        Dim SKU_COMMON4 = "" & RowData.Item(IdxColumnName.SKU_COMMON4.ToString)
        Dim SKU_COMMON5 = "" & RowData.Item(IdxColumnName.SKU_COMMON5.ToString)
        Dim SKU_COMMON6 = "" & RowData.Item(IdxColumnName.SKU_COMMON6.ToString)
        Dim SKU_COMMON7 = "" & RowData.Item(IdxColumnName.SKU_COMMON7.ToString)
        Dim SKU_COMMON8 = "" & RowData.Item(IdxColumnName.SKU_COMMON8.ToString)
        Dim SKU_COMMON9 = "" & RowData.Item(IdxColumnName.SKU_COMMON9.ToString)
        Dim SKU_COMMON10 = "" & RowData.Item(IdxColumnName.SKU_COMMON10.ToString)
        Dim SKU_L = 0 & RowData.Item(IdxColumnName.SKU_L.ToString)
        Dim SKU_W = 0 & RowData.Item(IdxColumnName.SKU_W.ToString)
        Dim SKU_H = 0 & RowData.Item(IdxColumnName.SKU_H.ToString)
        Dim SKU_WEIGHT = 0 & RowData.Item(IdxColumnName.SKU_WEIGHT.ToString)
        Dim SKU_VALUE = 0 & RowData.Item(IdxColumnName.SKU_VALUE.ToString)
        Dim SKU_UNIT = "" & RowData.Item(IdxColumnName.SKU_UNIT.ToString)
        Dim INBOUND_UNIT = "" & RowData.Item(IdxColumnName.INBOUND_UNIT.ToString)
        Dim OUTBOUND_UNIT = "" & RowData.Item(IdxColumnName.OUTBOUND_UNIT.ToString)
        Dim HIGH_WATER = 0 & RowData.Item(IdxColumnName.HIGH_WATER.ToString)
        Dim LOW_WATER = 0 & RowData.Item(IdxColumnName.LOW_WATER.ToString)
        Dim AVAILABLE_DAYS = 0 & RowData.Item(IdxColumnName.AVAILABLE_DAYS.ToString)
        Dim SAVE_DAYS = 0 & RowData.Item(IdxColumnName.SAVE_DAYS.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim WEIGHT_DIFFERENCE = 0 & RowData.Item(IdxColumnName.WEIGHT_DIFFERENCE.ToString)
        Dim ENABLE = BooleanConvertToInteger(RowData.Item(IdxColumnName.ENABLE.ToString))
        Dim EFFECTIVE_DATE = "" & RowData.Item(IdxColumnName.EFFECTIVE_DATE.ToString)
        Dim FAILURE_DATE = "" & RowData.Item(IdxColumnName.FAILURE_DATE.ToString)
        Dim QC_METHOD = "" & RowData.Item(IdxColumnName.QC_METHOD.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Info = New clsSKU(SKU_NO, SKU_ID1, SKU_ID2, SKU_ID3, SKU_ALIS1, SKU_ALIS2, SKU_DESC, SKU_CATALOG, SKU_TYPE1, SKU_TYPE2, SKU_TYPE3, SKU_COMMON1, SKU_COMMON2, SKU_COMMON3, SKU_COMMON4, SKU_COMMON5, SKU_COMMON6, SKU_COMMON7, SKU_COMMON8, SKU_COMMON9, SKU_COMMON10, SKU_L, SKU_W, SKU_H, SKU_WEIGHT, SKU_VALUE, SKU_UNIT, INBOUND_UNIT, OUTBOUND_UNIT, HIGH_WATER, LOW_WATER, AVAILABLE_DAYS, SAVE_DAYS, CREATE_TIME, UPDATE_TIME, WEIGHT_DIFFERENCE, ENABLE, EFFECTIVE_DATE, FAILURE_DATE, QC_METHOD, COMMENTS)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
