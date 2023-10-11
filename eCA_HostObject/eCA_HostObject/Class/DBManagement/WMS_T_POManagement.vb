Partial Class WMS_T_POManagement
  Public Shared TableName As String = "WMS_T_PO"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    PO_ID
    PO_TYPE1
    PO_TYPE2
    PO_TYPE3
    WO_TYPE
    PRIORITY
    CREATE_TIME
    START_TIME
    FINISH_TIME
    USER_ID
    CUSTOMER_NO
    CLASS_NO
    SHIPPING_NO
    PO_STATUS
    WRITE_OFF_NO
    AUTO_BOUND
    H_PO_CREATE_TIME
    H_PO_FINISH_TIME
    H_PO_STEP_NO
    H_PO_ORDER_TYPE
    H_PO1
    H_PO2
    H_PO3
    H_PO4
    H_PO5
    H_PO6
    H_PO7
    H_PO8
    H_PO9
    H_PO10
    H_PO11
    H_PO12
    H_PO13
    H_PO14
    H_PO15
    H_PO16
    H_PO17
    H_PO18
    H_PO19
    H_PO20
    SUPPLIER_NO
    PO_KEY1
    PO_KEY2
    PO_KEY3
    PO_KEY4
    PO_KEY5
  End Enum

  '- GetSQL
  '-請將 clsPO 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsPO) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92}) values ('{3}','{5}','{7}','{9}',{11},{13},'{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}',{31},{33},'{35}','{37}',{39},'{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}')",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_TYPE1.ToString, Info.PO_Type1,
      IdxColumnName.PO_TYPE2.ToString, Info.PO_Type2,
      IdxColumnName.PO_TYPE3.ToString, Info.PO_Type3,
      IdxColumnName.WO_TYPE.ToString, CInt(Info.WO_Type),
      IdxColumnName.PRIORITY.ToString, Info.Priority,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.START_TIME.ToString, Info.Start_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.Finish_Time,
      IdxColumnName.USER_ID.ToString, Info.User_ID,
      IdxColumnName.CUSTOMER_NO.ToString, Info.Customer_No,
      IdxColumnName.CLASS_NO.ToString, Info.Class_No,
      IdxColumnName.SHIPPING_NO.ToString, Info.Shipping_No,
      IdxColumnName.WRITE_OFF_NO.ToString, Info.Write_Off_No,
      IdxColumnName.PO_STATUS.ToString, CInt(Info.PO_Status),
      IdxColumnName.AUTO_BOUND.ToString, BooleanConvertToInteger(Info.Auto_Bound),
      IdxColumnName.H_PO_CREATE_TIME.ToString, Info.H_PO_CREATE_TIME,
      IdxColumnName.H_PO_FINISH_TIME.ToString, Info.H_PO_FINISH_TIME,
      IdxColumnName.H_PO_STEP_NO.ToString, Info.H_PO_STEP_NO,
      IdxColumnName.H_PO_ORDER_TYPE.ToString, Info.H_PO_ORDER_TYPE,
      IdxColumnName.H_PO1.ToString, Info.H_PO1,
      IdxColumnName.H_PO2.ToString, Info.H_PO2,
      IdxColumnName.H_PO3.ToString, Info.H_PO3,
      IdxColumnName.H_PO4.ToString, Info.H_PO4,
      IdxColumnName.H_PO5.ToString, Info.H_PO5,
      IdxColumnName.H_PO6.ToString, Info.H_PO6,
      IdxColumnName.H_PO7.ToString, Info.H_PO7,
      IdxColumnName.H_PO8.ToString, Info.H_PO8,
      IdxColumnName.H_PO9.ToString, Info.H_PO9,
      IdxColumnName.H_PO10.ToString, Info.H_PO10,
      IdxColumnName.H_PO11.ToString, Info.H_PO11,
      IdxColumnName.H_PO12.ToString, Info.H_PO12,
      IdxColumnName.H_PO13.ToString, Info.H_PO13,
      IdxColumnName.H_PO14.ToString, Info.H_PO14,
      IdxColumnName.H_PO15.ToString, Info.H_PO15,
      IdxColumnName.H_PO16.ToString, Info.H_PO16,
      IdxColumnName.H_PO17.ToString, Info.H_PO17,
      IdxColumnName.H_PO18.ToString, Info.H_PO18,
      IdxColumnName.H_PO19.ToString, Info.H_PO19,
      IdxColumnName.H_PO20.ToString, Info.H_PO20,
      IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
      IdxColumnName.PO_KEY1.ToString, Info.PO_KEY1,
      IdxColumnName.PO_KEY2.ToString, Info.PO_KEY2,
      IdxColumnName.PO_KEY3.ToString, Info.PO_KEY3,
      IdxColumnName.PO_KEY4.ToString, Info.PO_KEY4,
      IdxColumnName.PO_KEY5.ToString, Info.PO_KEY5
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsPO) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsPO) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}={11},{12}={13},{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}={31},{32}={33},{34}='{35}',{36}='{37}',{38}={39},{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.PO_TYPE1.ToString, Info.PO_Type1,
      IdxColumnName.PO_TYPE2.ToString, Info.PO_Type2,
      IdxColumnName.PO_TYPE3.ToString, Info.PO_Type3,
      IdxColumnName.WO_TYPE.ToString, CInt(Info.WO_Type),
      IdxColumnName.PRIORITY.ToString, Info.Priority,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.START_TIME.ToString, Info.Start_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.Finish_Time,
      IdxColumnName.USER_ID.ToString, Info.User_ID,
      IdxColumnName.CUSTOMER_NO.ToString, Info.Customer_No,
      IdxColumnName.CLASS_NO.ToString, Info.Class_No,
      IdxColumnName.SHIPPING_NO.ToString, Info.Shipping_No,
      IdxColumnName.WRITE_OFF_NO.ToString, Info.Write_Off_No,
      IdxColumnName.PO_STATUS.ToString, CInt(Info.PO_Status),
      IdxColumnName.AUTO_BOUND.ToString, BooleanConvertToInteger(Info.Auto_Bound),
      IdxColumnName.H_PO_CREATE_TIME.ToString, Info.H_PO_CREATE_TIME,
      IdxColumnName.H_PO_FINISH_TIME.ToString, Info.H_PO_FINISH_TIME,
      IdxColumnName.H_PO_STEP_NO.ToString, Info.H_PO_STEP_NO,
      IdxColumnName.H_PO_ORDER_TYPE.ToString, Info.H_PO_ORDER_TYPE,
      IdxColumnName.H_PO1.ToString, Info.H_PO1,
      IdxColumnName.H_PO2.ToString, Info.H_PO2,
      IdxColumnName.H_PO3.ToString, Info.H_PO3,
      IdxColumnName.H_PO4.ToString, Info.H_PO4,
      IdxColumnName.H_PO5.ToString, Info.H_PO5,
      IdxColumnName.H_PO6.ToString, Info.H_PO6,
      IdxColumnName.H_PO7.ToString, Info.H_PO7,
      IdxColumnName.H_PO8.ToString, Info.H_PO8,
      IdxColumnName.H_PO9.ToString, Info.H_PO9,
      IdxColumnName.H_PO10.ToString, Info.H_PO10,
      IdxColumnName.H_PO11.ToString, Info.H_PO11,
      IdxColumnName.H_PO12.ToString, Info.H_PO12,
      IdxColumnName.H_PO13.ToString, Info.H_PO13,
      IdxColumnName.H_PO14.ToString, Info.H_PO14,
      IdxColumnName.H_PO15.ToString, Info.H_PO15,
      IdxColumnName.H_PO16.ToString, Info.H_PO16,
      IdxColumnName.H_PO17.ToString, Info.H_PO17,
      IdxColumnName.H_PO18.ToString, Info.H_PO18,
      IdxColumnName.H_PO19.ToString, Info.H_PO19,
      IdxColumnName.H_PO20.ToString, Info.H_PO20,
      IdxColumnName.SUPPLIER_NO.ToString, Info.SUPPLIER_NO,
      IdxColumnName.PO_KEY1.ToString, Info.PO_KEY1,
      IdxColumnName.PO_KEY2.ToString, Info.PO_KEY2,
      IdxColumnName.PO_KEY3.ToString, Info.PO_KEY3,
      IdxColumnName.PO_KEY4.ToString, Info.PO_KEY4,
      IdxColumnName.PO_KEY5.ToString, Info.PO_KEY5
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
  Public Shared Function GetWMS_T_PODataListByALL() As List(Of clsPO)
    Try
      Dim _lstReturn As New List(Of clsPO)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsPO = Nothing
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
  Public Shared Function GetclsPOListByPO_ID(ByVal po_id As String) As List(Of clsPO)
    Try
      Dim _lstReturn As New List(Of clsPO)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
              strSQL,
              TableName,
              IdxColumnName.PO_ID.ToString, po_id
              )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsPO = Nothing
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

  Public Shared Function GetclsWMS_T_POListByPO_ID_OrderType(ByVal po_id As String, ByVal OrderType As enuOrderType) As Dictionary(Of String, clsPO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet

          strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' and {4}={5} ",
          strSQL,
          TableName,
          IdxColumnName.PO_ID.ToString, po_id,
          IdxColumnName.H_PO_ORDER_TYPE, CInt(OrderType)
          )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPO
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從資料庫抓取PO的資料
  Public Shared Function GetPODictionaryByPOID(ByVal PO_ID As String) As Dictionary(Of String, clsPO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If PO_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_ID.ToString, PO_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_ID.ToString, PO_ID)
            End If
          End If
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select * from {1} {2} ",
    strSQL,
  TableName,
  strWhere
  )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPO = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從資料庫抓取PO的資料
  Public Shared Function GetPODictionaryByALL() As Dictionary(Of String, clsPO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
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
              Dim Info As clsPO = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetPODictionaryByWO_TYPE(ByVal WO_TYPE As enuWOType) As Dictionary(Of String, clsPO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select * from {1} WHERE {2} = '{3}'",
    strSQL,
  TableName,
  IdxColumnName.WO_TYPE, WO_TYPE
  )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPO = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從資料庫抓取PO的資料
  Public Shared Function GetPODictionaryByPOID_ORDERTYPE(ByVal PO_ID As String, ByVal ORDERTYPE As String) As Dictionary(Of String, clsPO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If PO_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PO_ID.ToString, PO_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PO_ID.ToString, PO_ID)
            End If
          End If
          If ORDERTYPE <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.H_PO_ORDER_TYPE.ToString, ORDERTYPE)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.H_PO_ORDER_TYPE.ToString, ORDERTYPE)
            End If
          End If


          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select * from {1} {2} ",
              strSQL,
            TableName,
            strWhere
            )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPO = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '從資料庫抓取PO的資料
  Public Shared Function GetPODictionaryBydicPOID(ByVal dicPOID As Dictionary(Of String, String)) As Dictionary(Of String, clsPO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          Dim strPOList As String = ""
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          'For Each PO_ID As String In dicPOID.Values
          '  If strPOList = "" Then
          '    strPOList = "'" & PO_ID & "'"
          '  Else
          '    strPOList = strPOList & ",'" & PO_ID & "'"
          '  End If
          'Next
          'If strWhere = "" Then
          '  strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.PO_ID.ToString, strPOList)
          'Else
          '  strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.PO_ID.ToString, strPOList)
          'End If
          'Dim strSQL As String = String.Empty
          'Dim DatasetMessage As New DataSet
          'strSQL = String.Format("Select * from {1} {2} ",
          '    strSQL,
          '  TableName,
          '  strWhere
          '  )
          'SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          'DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          'If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          '  For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
          '    Dim Info As clsPO = Nothing
          '    If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
          '      If Info IsNot Nothing Then
          '        If ret_dic.ContainsKey(Info.gid) = False Then
          '          ret_dic.Add(Info.gid, Info)
          '        End If
          '      Else
          '        SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '      End If
          '    Else
          '      SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '    End If

          '  Next
          'End If

          Dim count_flag = 0
          For i = 0 To dicPOID.Count - 1
            If strPOList = "" Then
              strPOList = "'" & dicPOID.Keys(i) & "'"
            Else
              strPOList = strPOList & ",'" & dicPOID.Keys(i) & "'"
            End If
            If i - count_flag > 800 OrElse i = (dicPOID.Count - 1) Then
              count_flag = i
              strWhere = ""
              If strWhere = "" Then
                strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.PO_ID.ToString, strPOList)
              Else
                strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.PO_ID.ToString, strPOList)
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
                  Dim Info As clsPO = Nothing
                  SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Next
              End If
              strPOList = ""
            End If
          Next

        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetHPODictionaryBydicPOID(ByVal dicPOID As Dictionary(Of String, String)) As Dictionary(Of String, clsPO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          Dim strPOList As String = ""
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet
          'For Each PO_ID As String In dicPOID.Values
          '  If strPOList = "" Then
          '    strPOList = "'" & PO_ID & "'"
          '  Else
          '    strPOList = strPOList & ",'" & PO_ID & "'"
          '  End If
          'Next
          'If strWhere = "" Then
          '  strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.PO_ID.ToString, strPOList)
          'Else
          '  strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.PO_ID.ToString, strPOList)
          'End If
          'Dim strSQL As String = String.Empty
          'Dim DatasetMessage As New DataSet
          'strSQL = String.Format("Select * from {1} {2} ",
          '    strSQL,
          '  TableName,
          '  strWhere
          '  )
          'SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          'DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          'If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          '  For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
          '    Dim Info As clsPO = Nothing
          '    If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
          '      If Info IsNot Nothing Then
          '        If ret_dic.ContainsKey(Info.gid) = False Then
          '          ret_dic.Add(Info.gid, Info)
          '        End If
          '      Else
          '        SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '      End If
          '    Else
          '      SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '    End If

          '  Next
          'End If

          Dim count_flag = 0
          For i = 0 To dicPOID.Count - 1
            If strPOList = "" Then
              strPOList = "'" & dicPOID.Keys(i) & "'"
            Else
              strPOList = strPOList & ",'" & dicPOID.Keys(i) & "'"
            End If
            If i - count_flag > 800 OrElse i = (dicPOID.Count - 1) Then
              count_flag = i
              strWhere = ""
              If strWhere = "" Then
                strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.PO_ID.ToString, strPOList)
              Else
                strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.PO_ID.ToString, strPOList)
              End If
              strSQL = String.Format("Select * from WMS_H_PO {2} ",
                  strSQL,
                  TableName,
                  strWhere
              )
              SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
              If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
                For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                  Dim Info As clsPO = Nothing
                  SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Next
              End If
              strPOList = ""
            End If
          Next

        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function




  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsPO, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_TYPE1 = "" & RowData.Item(IdxColumnName.PO_TYPE1.ToString)
        Dim PO_TYPE2 = "" & RowData.Item(IdxColumnName.PO_TYPE2.ToString)
        Dim PO_TYPE3 = "" & RowData.Item(IdxColumnName.PO_TYPE3.ToString)
        Dim WO_TYPE = 0 & RowData.Item(IdxColumnName.WO_TYPE.ToString)
        Dim PRIORITY = 0 & RowData.Item(IdxColumnName.PRIORITY.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim START_TIME = "" & RowData.Item(IdxColumnName.START_TIME.ToString)
        Dim FINISH_TIME = "" & RowData.Item(IdxColumnName.FINISH_TIME.ToString)
        Dim USER_ID = "" & RowData.Item(IdxColumnName.USER_ID.ToString)
        Dim CUSTOMER_NO = "" & RowData.Item(IdxColumnName.CUSTOMER_NO.ToString)
        Dim CLASS_NO = "" & RowData.Item(IdxColumnName.CLASS_NO.ToString)
        Dim SHIPPING_NO = "" & RowData.Item(IdxColumnName.SHIPPING_NO.ToString)
        Dim PO_STATUS = 0 & RowData.Item(IdxColumnName.PO_STATUS.ToString)
        Dim WRITE_OFF_NO = "" & RowData.Item(IdxColumnName.WRITE_OFF_NO.ToString)
        Dim AUTO_FLAG = IntegerConvertToBoolean(0 & RowData.Item(IdxColumnName.AUTO_BOUND.ToString))
        Dim H_PO_CREATE_TIME = "" & RowData.Item(IdxColumnName.H_PO_CREATE_TIME.ToString)
        Dim H_PO_FINISH_TIME = "" & RowData.Item(IdxColumnName.H_PO_FINISH_TIME.ToString)
        Dim H_PO_STEP_NO = 0 & RowData.Item(IdxColumnName.H_PO_STEP_NO.ToString)
        Dim H_PO_ORDER_TYPE = "" & RowData.Item(IdxColumnName.H_PO_ORDER_TYPE.ToString)
        Dim H_PO1 = "" & RowData.Item(IdxColumnName.H_PO1.ToString)
        Dim H_PO2 = "" & RowData.Item(IdxColumnName.H_PO2.ToString)
        Dim H_PO3 = "" & RowData.Item(IdxColumnName.H_PO3.ToString)
        Dim H_PO4 = "" & RowData.Item(IdxColumnName.H_PO4.ToString)
        Dim H_PO5 = "" & RowData.Item(IdxColumnName.H_PO5.ToString)
        Dim H_PO6 = "" & RowData.Item(IdxColumnName.H_PO6.ToString)
        Dim H_PO7 = "" & RowData.Item(IdxColumnName.H_PO7.ToString)
        Dim H_PO8 = "" & RowData.Item(IdxColumnName.H_PO8.ToString)
        Dim H_PO9 = "" & RowData.Item(IdxColumnName.H_PO9.ToString)
        Dim H_PO10 = "" & RowData.Item(IdxColumnName.H_PO10.ToString)
        Dim H_PO11 = "" & RowData.Item(IdxColumnName.H_PO11.ToString)
        Dim H_PO12 = "" & RowData.Item(IdxColumnName.H_PO12.ToString)
        Dim H_PO13 = "" & RowData.Item(IdxColumnName.H_PO13.ToString)
        Dim H_PO14 = "" & RowData.Item(IdxColumnName.H_PO14.ToString)
        Dim H_PO15 = "" & RowData.Item(IdxColumnName.H_PO15.ToString)
        Dim H_PO16 = "" & RowData.Item(IdxColumnName.H_PO16.ToString)
        Dim H_PO17 = "" & RowData.Item(IdxColumnName.H_PO17.ToString)
        Dim H_PO18 = "" & RowData.Item(IdxColumnName.H_PO18.ToString)
        Dim H_PO19 = "" & RowData.Item(IdxColumnName.H_PO19.ToString)
        Dim H_PO20 = "" & RowData.Item(IdxColumnName.H_PO20.ToString)
        Dim SUPPLIER_NO = "" & RowData.Item(IdxColumnName.SUPPLIER_NO.ToString)
        Dim PO_KEY1 = "" & RowData.Item(IdxColumnName.PO_KEY1.ToString)       'Vito_19b16
        Dim PO_KEY2 = "" & RowData.Item(IdxColumnName.PO_KEY2.ToString)       'Vito_19b16
        Dim PO_KEY3 = "" & RowData.Item(IdxColumnName.PO_KEY3.ToString)       'Vito_19b16
        Dim PO_KEY4 = "" & RowData.Item(IdxColumnName.PO_KEY4.ToString)       'Vito_19b16
        Dim PO_KEY5 = "" & RowData.Item(IdxColumnName.PO_KEY5.ToString)       'Vito_19b16
        Info = New clsPO(PO_ID, PO_TYPE1, PO_TYPE2, PO_TYPE3, WO_TYPE, PRIORITY, CREATE_TIME, START_TIME, FINISH_TIME, USER_ID, CUSTOMER_NO, CLASS_NO, SHIPPING_NO, PO_STATUS, WRITE_OFF_NO, AUTO_FLAG,
                                 H_PO_CREATE_TIME, H_PO_FINISH_TIME, H_PO_STEP_NO, H_PO_ORDER_TYPE, H_PO1, H_PO2, H_PO3, H_PO4, H_PO5, H_PO6, H_PO7, H_PO8, H_PO9, H_PO10, H_PO11, H_PO12, H_PO13, H_PO14, H_PO15, H_PO16, H_PO17, H_PO18, H_PO19, H_PO20, SUPPLIER_NO,
                                 PO_KEY1, PO_KEY2, PO_KEY3, PO_KEY4, PO_KEY5) 'Vito_19b16
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Shared Function GetclsWMS_T_POListByPO_ID(ByVal po_id As String) As Dictionary(Of String, clsPO)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPO)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strSQL As String = String.Empty
          Dim DatasetMessage As New DataSet

          strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
          strSQL,
          TableName,
          IdxColumnName.PO_ID.ToString, po_id
          )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPO
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPO Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPO Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
