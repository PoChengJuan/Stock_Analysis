Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_T_POManagement_BackUp
  Public Shared TableName As String = "WMS_T_PO"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsPO_Back)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    PO_ID
    PO_TYPE_1
    PO_TYPE_2
    PO_TYPE_3
    WO_TYPE
    PRIORITY
    CREATE_TIME
    START_TIME
    FINISH_TIME
    USER_ID
    CUSTOMER_NO
    CLASS_NO
    SHIPPING_NO
    OWNER_NO
    FACTORY_NO
    DEST_AREA_NO
    WRITE_OFF_NO
    PO_STATUS
    HOST_CREATE_TIME
    HOST_FINISH_TIME
    HOST_STEP_NO
    HOST_OrderType
    HOST_CUSTOMER_NO
    HOST_CUSTOMER_ID
    HOST_CUSTOMER_COMMON1
    HOST_CUSTOMER_COMMON2
    HOST_CUSTOMER_COMMON3
    HOST_CUSTOMER_COMMON4
    HOST_CUSTOMER_COMMON5
    HOST_OWNER_NO
    HOST_OWNER_ID
    HOST_OWNER_COMMON1
    HOST_OWNER_COMMON2
    HOST_OWNER_COMMON3
    HOST_OWNER_COMMON4
    HOST_OWNER_COMMON5
    HOST_COMMON1
    HOST_COMMON2
    HOST_COMMON3
    HOST_COMMON4
    HOST_COMMON5
    HOST_COMMON6
    HOST_COMMON7
    HOST_COMMON8
    HOST_COMMON9
    HOST_COMMON10
    HOST_COMMON11
    HOST_COMMON12
    HOST_COMMON13
    HOST_COMMON14
    HOST_COMMON15
    HOST_COMMON16
    HOST_COMMON17
    HOST_COMMON18
    HOST_COMMON19
    HOST_COMMON20
    HOST_COMMENTS
  End Enum

  '- GetSQL
  '-請將 clsPO 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsPO_Back) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92},{94},{96},{98},{100},{102},{104},{106},{108},{110}) values ('{3}','{5}','{7}','{9}',{11},{13},'{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}',{37},'{39}','{41}',{43},{45},'{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}','{95}','{97}','{99}','{101}','{103}','{105}','{107}','{109}','{111}')",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.get_PO_ID,
      IdxColumnName.PO_TYPE_1.ToString, Info.get_PO_Type_1,
      IdxColumnName.PO_TYPE_2.ToString, Info.get_PO_Type_2,
      IdxColumnName.PO_TYPE_3.ToString, Info.get_PO_Type_3,
      IdxColumnName.WO_TYPE.ToString, CInt(Info.get_WO_Type),
      IdxColumnName.PRIORITY.ToString, Info.get_Priority,
      IdxColumnName.CREATE_TIME.ToString, Info.get_Create_Time,
      IdxColumnName.START_TIME.ToString, Info.get_Start_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.get_Finish_Time,
      IdxColumnName.USER_ID.ToString, Info.get_User_ID,
      IdxColumnName.CUSTOMER_NO.ToString, Info.get_Customer_No,
      IdxColumnName.CLASS_NO.ToString, Info.get_Class_NO,
      IdxColumnName.SHIPPING_NO.ToString, Info.get_Shipping_No,
      IdxColumnName.OWNER_NO.ToString, Info.get_OWNER_NO,
      IdxColumnName.FACTORY_NO.ToString, Info.get_Factory_NO,
      IdxColumnName.DEST_AREA_NO.ToString, Info.get_Dest_Area_NO,
      IdxColumnName.WRITE_OFF_NO.ToString, Info.get_WRITE_OFF_NO,
      IdxColumnName.PO_STATUS.ToString, CInt(Info.get_PO_STATUS),
      IdxColumnName.HOST_CREATE_TIME.ToString, Info.get_HOST_CREATE_TIME,
      IdxColumnName.HOST_FINISH_TIME.ToString, Info.get_HOST_FINISH_TIME,
      IdxColumnName.HOST_STEP_NO.ToString, Info.get_HOST_STEP_NO,
      IdxColumnName.HOST_OrderType.ToString, CInt(Info.get_HOST_OrderType),
      IdxColumnName.HOST_CUSTOMER_ID.ToString, Info.get_HOST_CUSTOMER_ID,
      IdxColumnName.HOST_CUSTOMER_COMMON1.ToString, Info.get_HOST_CUSTOMER_COMMON1,
      IdxColumnName.HOST_CUSTOMER_COMMON2.ToString, Info.get_HOST_CUSTOMER_COMMON2,
      IdxColumnName.HOST_CUSTOMER_COMMON3.ToString, Info.get_HOST_CUSTOMER_COMMON3,
      IdxColumnName.HOST_CUSTOMER_COMMON4.ToString, Info.get_HOST_CUSTOMER_COMMON4,
      IdxColumnName.HOST_CUSTOMER_COMMON5.ToString, Info.get_HOST_CUSTOMER_COMMON5,
      IdxColumnName.HOST_OWNER_ID.ToString, Info.get_HOST_OWNER_ID,
      IdxColumnName.HOST_OWNER_COMMON1.ToString, Info.get_HOST_OWNER_COMMON1,
      IdxColumnName.HOST_OWNER_COMMON2.ToString, Info.get_HOST_OWNER_COMMON2,
      IdxColumnName.HOST_OWNER_COMMON3.ToString, Info.get_HOST_OWNER_COMMON3,
      IdxColumnName.HOST_OWNER_COMMON4.ToString, Info.get_HOST_OWNER_COMMON4,
      IdxColumnName.HOST_OWNER_COMMON5.ToString, Info.get_HOST_OWNER_COMMON5,
      IdxColumnName.HOST_COMMON1.ToString, Info.get_HOST_COMMON1,
      IdxColumnName.HOST_COMMON2.ToString, Info.get_HOST_COMMON2,
      IdxColumnName.HOST_COMMON3.ToString, Info.get_HOST_COMMON3,
      IdxColumnName.HOST_COMMON4.ToString, Info.get_HOST_COMMON4,
      IdxColumnName.HOST_COMMON5.ToString, Info.get_HOST_COMMON5,
      IdxColumnName.HOST_COMMON6.ToString, Info.get_HOST_COMMON6,
      IdxColumnName.HOST_COMMON7.ToString, Info.get_HOST_COMMON7,
      IdxColumnName.HOST_COMMON8.ToString, Info.get_HOST_COMMON8,
      IdxColumnName.HOST_COMMON9.ToString, Info.get_HOST_COMMON9,
      IdxColumnName.HOST_COMMON10.ToString, Info.get_HOST_COMMON10,
      IdxColumnName.HOST_COMMON11.ToString, Info.get_HOST_COMMON11,
      IdxColumnName.HOST_COMMON12.ToString, Info.get_HOST_COMMON12,
      IdxColumnName.HOST_COMMON13.ToString, Info.get_HOST_COMMON13,
      IdxColumnName.HOST_COMMON14.ToString, Info.get_HOST_COMMON14,
      IdxColumnName.HOST_COMMON15.ToString, Info.get_HOST_COMMON15,
      IdxColumnName.HOST_COMMON16.ToString, Info.get_HOST_COMMON16,
      IdxColumnName.HOST_COMMON17.ToString, Info.get_HOST_COMMON17,
      IdxColumnName.HOST_COMMON18.ToString, Info.get_HOST_COMMON18,
      IdxColumnName.HOST_COMMON19.ToString, Info.get_HOST_COMMON19,
      IdxColumnName.HOST_COMMON20.ToString, Info.get_HOST_COMMON20,
      IdxColumnName.HOST_COMMENTS.ToString, Info.get_HOST_COMMENTS
     )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsPO_Back) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.get_PO_ID
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsPO_Back) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}={11},{12}={13},{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}={37},{38}='{39}',{40}='{41}',{42}={43},{44}={45},{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}',{94}='{95}',{96}='{97}',{98}='{99}',{100}='{101}',{102}='{103}',{104}='{105}',{106}='{107}',{108}='{109}',{110}='{111}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.get_PO_ID,
      IdxColumnName.PO_TYPE_1.ToString, Info.get_PO_Type_1,
      IdxColumnName.PO_TYPE_2.ToString, Info.get_PO_Type_2,
      IdxColumnName.PO_TYPE_3.ToString, Info.get_PO_Type_3,
      IdxColumnName.WO_TYPE.ToString, Info.get_WO_Type,
      IdxColumnName.PRIORITY.ToString, Info.get_Priority,
      IdxColumnName.CREATE_TIME.ToString, Info.get_Create_Time,
      IdxColumnName.START_TIME.ToString, Info.get_Start_Time,
      IdxColumnName.FINISH_TIME.ToString, Info.get_Finish_Time,
      IdxColumnName.USER_ID.ToString, Info.get_User_ID,
      IdxColumnName.CUSTOMER_NO.ToString, Info.get_Customer_No,
      IdxColumnName.CLASS_NO.ToString, Info.get_Class_NO,
      IdxColumnName.SHIPPING_NO.ToString, Info.get_Shipping_No,
      IdxColumnName.OWNER_NO.ToString, Info.get_OWNER_NO,
      IdxColumnName.FACTORY_NO.ToString, Info.get_Factory_NO,
      IdxColumnName.DEST_AREA_NO.ToString, Info.get_Dest_Area_NO,
      IdxColumnName.WRITE_OFF_NO.ToString, Info.get_WRITE_OFF_NO,
      IdxColumnName.PO_STATUS.ToString, Info.get_PO_STATUS,
      IdxColumnName.HOST_CREATE_TIME.ToString, Info.get_HOST_CREATE_TIME,
      IdxColumnName.HOST_FINISH_TIME.ToString, Info.get_HOST_FINISH_TIME,
      IdxColumnName.HOST_STEP_NO.ToString, Info.get_HOST_STEP_NO,
      IdxColumnName.HOST_OrderType.ToString, Info.get_HOST_OrderType,
      IdxColumnName.HOST_CUSTOMER_ID.ToString, Info.get_HOST_CUSTOMER_ID,
      IdxColumnName.HOST_CUSTOMER_COMMON1.ToString, Info.get_HOST_CUSTOMER_COMMON1,
      IdxColumnName.HOST_CUSTOMER_COMMON2.ToString, Info.get_HOST_CUSTOMER_COMMON2,
      IdxColumnName.HOST_CUSTOMER_COMMON3.ToString, Info.get_HOST_CUSTOMER_COMMON3,
      IdxColumnName.HOST_CUSTOMER_COMMON4.ToString, Info.get_HOST_CUSTOMER_COMMON4,
      IdxColumnName.HOST_CUSTOMER_COMMON5.ToString, Info.get_HOST_CUSTOMER_COMMON5,
      IdxColumnName.HOST_OWNER_ID.ToString, Info.get_HOST_OWNER_ID,
      IdxColumnName.HOST_OWNER_COMMON1.ToString, Info.get_HOST_OWNER_COMMON1,
      IdxColumnName.HOST_OWNER_COMMON2.ToString, Info.get_HOST_OWNER_COMMON2,
      IdxColumnName.HOST_OWNER_COMMON3.ToString, Info.get_HOST_OWNER_COMMON3,
      IdxColumnName.HOST_OWNER_COMMON4.ToString, Info.get_HOST_OWNER_COMMON4,
      IdxColumnName.HOST_OWNER_COMMON5.ToString, Info.get_HOST_OWNER_COMMON5,
      IdxColumnName.HOST_COMMON1.ToString, Info.get_HOST_COMMON1,
      IdxColumnName.HOST_COMMON2.ToString, Info.get_HOST_COMMON2,
      IdxColumnName.HOST_COMMON3.ToString, Info.get_HOST_COMMON3,
      IdxColumnName.HOST_COMMON4.ToString, Info.get_HOST_COMMON4,
      IdxColumnName.HOST_COMMON5.ToString, Info.get_HOST_COMMON5,
      IdxColumnName.HOST_COMMON6.ToString, Info.get_HOST_COMMON6,
      IdxColumnName.HOST_COMMON7.ToString, Info.get_HOST_COMMON7,
      IdxColumnName.HOST_COMMON8.ToString, Info.get_HOST_COMMON8,
      IdxColumnName.HOST_COMMON9.ToString, Info.get_HOST_COMMON9,
      IdxColumnName.HOST_COMMON10.ToString, Info.get_HOST_COMMON10,
      IdxColumnName.HOST_COMMON11.ToString, Info.get_HOST_COMMON11,
      IdxColumnName.HOST_COMMON12.ToString, Info.get_HOST_COMMON12,
      IdxColumnName.HOST_COMMON13.ToString, Info.get_HOST_COMMON13,
      IdxColumnName.HOST_COMMON14.ToString, Info.get_HOST_COMMON14,
      IdxColumnName.HOST_COMMON15.ToString, Info.get_HOST_COMMON15,
      IdxColumnName.HOST_COMMON16.ToString, Info.get_HOST_COMMON16,
      IdxColumnName.HOST_COMMON17.ToString, Info.get_HOST_COMMON17,
      IdxColumnName.HOST_COMMON18.ToString, Info.get_HOST_COMMON18,
      IdxColumnName.HOST_COMMON19.ToString, Info.get_HOST_COMMON19,
      IdxColumnName.HOST_COMMON20.ToString, Info.get_HOST_COMMON20,
      IdxColumnName.HOST_COMMENTS.ToString, Info.get_HOST_COMMENTS
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '- Add & Insert
  Public Shared Function AddWMS_T_POData(ByVal Info As clsPO_Back, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If AddlstWMS_T_POData(New List(Of clsPO_Back)({Info}), SendToDB) = True Then
          Return True
        End If '-載不載入記憶體都是呼叫同一個function
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function AddlstWMS_T_POData(ByVal Info As List(Of clsPO_Back), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True

        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).get_gid()
            If dicData.ContainsKey(key) = True Then
              SendMessageToLog("Add the same key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next

          If SendToDB Then
            If InsertWMS_T_PODataToDB(Info) Then
              If AddOrUpdateWMS_T_PODataToDictionary(Info) Then
                SendMessageToLog("InsertDic WMS_T_POData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("InsertDic WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("InsertDB WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If AddOrUpdateWMS_T_PODataToDictionary(Info) Then
              SendMessageToLog("InsertDic WMS_T_POData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("InsertDic WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If InsertWMS_T_PODataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("InsertDic WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function

  '- Update
  Public Shared Function UpdateWMS_T_POData(ByVal Info As clsPO_Back, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If UpdatelstWMS_T_POData(New List(Of clsPO_Back)({Info}), SendToDB) = True Then
          Return True
        End If
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function UpdatelstWMS_T_POData(ByVal Info As List(Of clsPO_Back), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True

        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).get_gid()
            If dicData.ContainsKey(key) = False Then
              SendMessageToLog("There is no key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next

          If SendToDB Then
            If UpdateWMS_T_PODataToDB(Info) Then
              If AddOrUpdateWMS_T_PODataToDictionary(Info) Then
                SendMessageToLog("UpdateDic WMS_T_POData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("UpdateDic WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("UpdateDB WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If AddOrUpdateWMS_T_PODataToDictionary(Info) Then
              SendMessageToLog("UpdateDic WMS_T_POData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("UpdateDic WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If UpdateWMS_T_PODataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("UpdateDB WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function

  '- Delete
  Public Shared Function DeleteWMS_T_POData(ByVal Info As clsPO_Back, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If DeletelstWMS_T_POData(New List(Of clsPO_Back)({Info}), SendToDB) = True Then
          Return True
        End If '-載不載入記憶體都是呼叫同一個function
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function DeletelstWMS_T_POData(ByVal Info As List(Of clsPO_Back), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True

        If DictionaryNeeded = 1 Then '-載入記憶體
          For i = 0 To Info.Count - 1
            Dim key As String = Info(i).get_gid()
            If dicData.ContainsKey(key) = False Then
              SendMessageToLog("There is no key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Next

          If SendToDB Then
            If DeleteWMS_T_PODataToDB(Info) Then
              If DeleteWMS_T_PODataToDictionary(Info) Then
                SendMessageToLog("DeleteDic WMS_T_POData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("DeleteDic WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("DeleteDB WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If DeleteWMS_T_PODataToDB(Info) Then
              SendMessageToLog("DeleteDic WMS_T_POData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("DeleteDB WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
          Return True
        Else
          If SendToDB Then
            If DeleteWMS_T_PODataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("DeleteDB WMS_T_POData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            SendMessageToLog("Do nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return True
          End If
        End If
        Return True
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function

  '- GET
  Public Shared Function GetWMS_T_PODataListByALL() As List(Of clsPO_Back)
    SyncLock objLock
      Try
        Dim _lstReturn As New List(Of clsPO_Back)
        If DictionaryNeeded = 1 Then '-載入記憶體
          Dim LinqFind As IEnumerable(Of clsPO_Back) = From TC In dicData Select TC.Value
          '- From TC In dicData Where TC.Value.xxx = xxx AND TC.Value.xxx = xxx AND TC.Value.xxx = xxx Select TC.Value '-範例
          For Each objTC As clsPO_Back In LinqFind
            _lstReturn.Add(objTC)
          Next

          Return _lstReturn
        Else
          If DBTool IsNot Nothing Then
            If DBTool.isConnection(DBTool.m_CN) = True Then
              Dim strSQL As String = String.Empty
              Dim rs As ADODB.Recordset = Nothing

              strSQL = String.Format("Select * from {0}", TableName)
              SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
              DBTool.SQLExcute(strSQL, rs)
              Dim DatasetMessage As New DataSet
              Dim OLEDBAdapter As New OleDbDataAdapter
              OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

              If DatasetMessage.Tables(TableName).Rows.Count > 0 Then
                For RowIndex = 0 To DatasetMessage.Tables(TableName).Rows.Count - 1
                  Dim Info As clsPO_Back = Nothing
                  SetInfoFromDB(Info, DatasetMessage.Tables(TableName).Rows(RowIndex))
                  _lstReturn.Add(Info)
                Next
              End If
            End If
          End If
          Return _lstReturn
        End If
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function
  Public Shared Function GetclsPOListByPO_ID(ByVal po_id As String) As List(Of clsPO_Back)
    SyncLock objLock
      Try
        Dim _lstReturn As New List(Of clsPO_Back)
        If DictionaryNeeded = 1 Then '-載入記憶體
          Dim strSQL As String = String.Empty
          Dim rs As ADODB.Recordset = Nothing

          Dim LinqFind As IEnumerable(Of clsPO_Back) = From TC In dicData Where TC.Value.get_PO_ID = po_id Select TC.Value
          '- From TC In dicData Where TC.Value.xxx = xxx AND TC.Value.xxx = xxx AND TC.Value.xxx = xxx Select TC.Value '-範例
          For Each objTC As clsPO_Back In LinqFind
            _lstReturn.Add(objTC)
          Next

          Return _lstReturn
        Else
          If DBTool IsNot Nothing Then
            If DBTool.isConnection(DBTool.m_CN) = True Then
              Dim strSQL As String = String.Empty
              Dim rs As ADODB.Recordset = Nothing

              strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
              strSQL,
              TableName,
              IdxColumnName.PO_ID.ToString, po_id
              )
              SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
              DBTool.SQLExcute(strSQL, rs)
              Dim DatasetMessage As New DataSet
              Dim OLEDBAdapter As New OleDbDataAdapter
              OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

              If DatasetMessage.Tables(TableName).Rows.Count > 0 Then
                For RowIndex = 0 To DatasetMessage.Tables(TableName).Rows.Count - 1
                  Dim Info As clsPO_Back = Nothing
                  SetInfoFromDB(Info, DatasetMessage.Tables(TableName).Rows(RowIndex))
                  _lstReturn.Add(Info)
                Next
              End If
            End If
          End If
          Return _lstReturn
        End If
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function

  '-Function


  '-以下為內部私人用
  Private Shared Function InsertWMS_T_PODataToDB(ByRef Info As List(Of clsPO_Back)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92},{94},{96},{98},{100},{102},{104},{106},{108},{110}) values ('{3}','{5}','{7}','{9}',{11},{13},'{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}',{37},'{39}','{41}',{43},{45},'{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}','{95}','{97}','{99}','{101}','{103}','{105}','{107}','{109}','{111}')",
        strSQL,
        TableName,
        IdxColumnName.PO_ID.ToString, CI.get_PO_ID,
        IdxColumnName.PO_TYPE_1.ToString, CI.get_PO_Type_1,
        IdxColumnName.PO_TYPE_2.ToString, CI.get_PO_Type_2,
        IdxColumnName.PO_TYPE_3.ToString, CI.get_PO_Type_3,
        IdxColumnName.WO_TYPE.ToString, CI.get_WO_Type,
        IdxColumnName.PRIORITY.ToString, CI.get_Priority,
        IdxColumnName.CREATE_TIME.ToString, CI.get_Create_Time,
        IdxColumnName.START_TIME.ToString, CI.get_Start_Time,
        IdxColumnName.FINISH_TIME.ToString, CI.get_Finish_Time,
        IdxColumnName.USER_ID.ToString, CI.get_User_ID,
        IdxColumnName.CUSTOMER_NO.ToString, CI.get_Customer_No,
        IdxColumnName.CLASS_NO.ToString, CI.get_Class_NO,
        IdxColumnName.SHIPPING_NO.ToString, CI.get_Shipping_No,
        IdxColumnName.OWNER_NO.ToString, CI.get_OWNER_NO,
        IdxColumnName.FACTORY_NO.ToString, CI.get_Factory_NO,
        IdxColumnName.DEST_AREA_NO.ToString, CI.get_Dest_Area_NO,
        IdxColumnName.WRITE_OFF_NO.ToString, CI.get_WRITE_OFF_NO,
        IdxColumnName.PO_STATUS.ToString, CI.get_PO_STATUS,
        IdxColumnName.HOST_CREATE_TIME.ToString, CI.get_HOST_CREATE_TIME,
        IdxColumnName.HOST_FINISH_TIME.ToString, CI.get_HOST_FINISH_TIME,
        IdxColumnName.HOST_STEP_NO.ToString, CI.get_HOST_STEP_NO,
        IdxColumnName.HOST_OrderType.ToString, CI.get_HOST_OrderType,
        IdxColumnName.HOST_CUSTOMER_ID.ToString, CI.get_HOST_CUSTOMER_ID,
        IdxColumnName.HOST_CUSTOMER_COMMON1.ToString, CI.get_HOST_CUSTOMER_COMMON1,
        IdxColumnName.HOST_CUSTOMER_COMMON2.ToString, CI.get_HOST_CUSTOMER_COMMON2,
        IdxColumnName.HOST_CUSTOMER_COMMON3.ToString, CI.get_HOST_CUSTOMER_COMMON3,
        IdxColumnName.HOST_CUSTOMER_COMMON4.ToString, CI.get_HOST_CUSTOMER_COMMON4,
        IdxColumnName.HOST_CUSTOMER_COMMON5.ToString, CI.get_HOST_CUSTOMER_COMMON5,
        IdxColumnName.HOST_OWNER_ID.ToString, CI.get_HOST_OWNER_ID,
        IdxColumnName.HOST_OWNER_COMMON1.ToString, CI.get_HOST_OWNER_COMMON1,
        IdxColumnName.HOST_OWNER_COMMON2.ToString, CI.get_HOST_OWNER_COMMON2,
        IdxColumnName.HOST_OWNER_COMMON3.ToString, CI.get_HOST_OWNER_COMMON3,
        IdxColumnName.HOST_OWNER_COMMON4.ToString, CI.get_HOST_OWNER_COMMON4,
        IdxColumnName.HOST_OWNER_COMMON5.ToString, CI.get_HOST_OWNER_COMMON5,
        IdxColumnName.HOST_COMMON1.ToString, CI.get_HOST_COMMON1,
        IdxColumnName.HOST_COMMON2.ToString, CI.get_HOST_COMMON2,
        IdxColumnName.HOST_COMMON3.ToString, CI.get_HOST_COMMON3,
        IdxColumnName.HOST_COMMON4.ToString, CI.get_HOST_COMMON4,
        IdxColumnName.HOST_COMMON5.ToString, CI.get_HOST_COMMON5,
        IdxColumnName.HOST_COMMON6.ToString, CI.get_HOST_COMMON6,
        IdxColumnName.HOST_COMMON7.ToString, CI.get_HOST_COMMON7,
        IdxColumnName.HOST_COMMON8.ToString, CI.get_HOST_COMMON8,
        IdxColumnName.HOST_COMMON9.ToString, CI.get_HOST_COMMON9,
        IdxColumnName.HOST_COMMON10.ToString, CI.get_HOST_COMMON10,
        IdxColumnName.HOST_COMMON11.ToString, CI.get_HOST_COMMON11,
        IdxColumnName.HOST_COMMON12.ToString, CI.get_HOST_COMMON12,
        IdxColumnName.HOST_COMMON13.ToString, CI.get_HOST_COMMON13,
        IdxColumnName.HOST_COMMON14.ToString, CI.get_HOST_COMMON14,
        IdxColumnName.HOST_COMMON15.ToString, CI.get_HOST_COMMON15,
        IdxColumnName.HOST_COMMON16.ToString, CI.get_HOST_COMMON16,
        IdxColumnName.HOST_COMMON17.ToString, CI.get_HOST_COMMON17,
        IdxColumnName.HOST_COMMON18.ToString, CI.get_HOST_COMMON18,
        IdxColumnName.HOST_COMMON19.ToString, CI.get_HOST_COMMON19,
        IdxColumnName.HOST_COMMON20.ToString, CI.get_HOST_COMMON20,
        IdxColumnName.HOST_COMMENTS.ToString, CI.get_HOST_COMMENTS
        )
        lstSql.Add(strSQL)
      Next
      If SendSQLToDB(lstSql) = True Then
        Return True
      Else
        SendMessageToLog("Insert to WMS_T_POData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateWMS_T_PODataToDB(ByRef Info As List(Of clsPO_Back)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}={11},{12}={13},{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}={37},{38}='{39}',{40}='{41}',{42}={43},{44}={45},{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}',{94}='{95}',{96}='{97}',{98}='{99}',{100}='{101}',{102}='{103}',{104}='{105}',{106}='{107}',{108}='{109}',{110}='{111}' WHERE {2}='{3}'",
        strSQL,
        TableName,
        IdxColumnName.PO_ID.ToString, CI.get_PO_ID,
        IdxColumnName.PO_TYPE_1.ToString, CI.get_PO_Type_1,
        IdxColumnName.PO_TYPE_2.ToString, CI.get_PO_Type_2,
        IdxColumnName.PO_TYPE_3.ToString, CI.get_PO_Type_3,
        IdxColumnName.WO_TYPE.ToString, CI.get_WO_Type,
        IdxColumnName.PRIORITY.ToString, CI.get_Priority,
        IdxColumnName.CREATE_TIME.ToString, CI.get_Create_Time,
        IdxColumnName.START_TIME.ToString, CI.get_Start_Time,
        IdxColumnName.FINISH_TIME.ToString, CI.get_Finish_Time,
        IdxColumnName.USER_ID.ToString, CI.get_User_ID,
        IdxColumnName.CUSTOMER_NO.ToString, CI.get_Customer_No,
        IdxColumnName.CLASS_NO.ToString, CI.get_Class_NO,
        IdxColumnName.SHIPPING_NO.ToString, CI.get_Shipping_No,
        IdxColumnName.OWNER_NO.ToString, CI.get_OWNER_NO,
        IdxColumnName.FACTORY_NO.ToString, CI.get_Factory_NO,
        IdxColumnName.DEST_AREA_NO.ToString, CI.get_Dest_Area_NO,
        IdxColumnName.WRITE_OFF_NO.ToString, CI.get_WRITE_OFF_NO,
        IdxColumnName.PO_STATUS.ToString, CI.get_PO_STATUS,
        IdxColumnName.HOST_CREATE_TIME.ToString, CI.get_HOST_CREATE_TIME,
        IdxColumnName.HOST_FINISH_TIME.ToString, CI.get_HOST_FINISH_TIME,
        IdxColumnName.HOST_STEP_NO.ToString, CI.get_HOST_STEP_NO,
        IdxColumnName.HOST_OrderType.ToString, CI.get_HOST_OrderType,
        IdxColumnName.HOST_CUSTOMER_ID.ToString, CI.get_HOST_CUSTOMER_ID,
        IdxColumnName.HOST_CUSTOMER_COMMON1.ToString, CI.get_HOST_CUSTOMER_COMMON1,
        IdxColumnName.HOST_CUSTOMER_COMMON2.ToString, CI.get_HOST_CUSTOMER_COMMON2,
        IdxColumnName.HOST_CUSTOMER_COMMON3.ToString, CI.get_HOST_CUSTOMER_COMMON3,
        IdxColumnName.HOST_CUSTOMER_COMMON4.ToString, CI.get_HOST_CUSTOMER_COMMON4,
        IdxColumnName.HOST_CUSTOMER_COMMON5.ToString, CI.get_HOST_CUSTOMER_COMMON5,
        IdxColumnName.HOST_OWNER_ID.ToString, CI.get_HOST_OWNER_ID,
        IdxColumnName.HOST_OWNER_COMMON1.ToString, CI.get_HOST_OWNER_COMMON1,
        IdxColumnName.HOST_OWNER_COMMON2.ToString, CI.get_HOST_OWNER_COMMON2,
        IdxColumnName.HOST_OWNER_COMMON3.ToString, CI.get_HOST_OWNER_COMMON3,
        IdxColumnName.HOST_OWNER_COMMON4.ToString, CI.get_HOST_OWNER_COMMON4,
        IdxColumnName.HOST_OWNER_COMMON5.ToString, CI.get_HOST_OWNER_COMMON5,
        IdxColumnName.HOST_COMMON1.ToString, CI.get_HOST_COMMON1,
        IdxColumnName.HOST_COMMON2.ToString, CI.get_HOST_COMMON2,
        IdxColumnName.HOST_COMMON3.ToString, CI.get_HOST_COMMON3,
        IdxColumnName.HOST_COMMON4.ToString, CI.get_HOST_COMMON4,
        IdxColumnName.HOST_COMMON5.ToString, CI.get_HOST_COMMON5,
        IdxColumnName.HOST_COMMON6.ToString, CI.get_HOST_COMMON6,
        IdxColumnName.HOST_COMMON7.ToString, CI.get_HOST_COMMON7,
        IdxColumnName.HOST_COMMON8.ToString, CI.get_HOST_COMMON8,
        IdxColumnName.HOST_COMMON9.ToString, CI.get_HOST_COMMON9,
        IdxColumnName.HOST_COMMON10.ToString, CI.get_HOST_COMMON10,
        IdxColumnName.HOST_COMMON11.ToString, CI.get_HOST_COMMON11,
        IdxColumnName.HOST_COMMON12.ToString, CI.get_HOST_COMMON12,
        IdxColumnName.HOST_COMMON13.ToString, CI.get_HOST_COMMON13,
        IdxColumnName.HOST_COMMON14.ToString, CI.get_HOST_COMMON14,
        IdxColumnName.HOST_COMMON15.ToString, CI.get_HOST_COMMON15,
        IdxColumnName.HOST_COMMON16.ToString, CI.get_HOST_COMMON16,
        IdxColumnName.HOST_COMMON17.ToString, CI.get_HOST_COMMON17,
        IdxColumnName.HOST_COMMON18.ToString, CI.get_HOST_COMMON18,
        IdxColumnName.HOST_COMMON19.ToString, CI.get_HOST_COMMON19,
        IdxColumnName.HOST_COMMON20.ToString, CI.get_HOST_COMMON20,
        IdxColumnName.HOST_COMMENTS.ToString, CI.get_HOST_COMMENTS
        )
        lstSql.Add(strSQL)
      Next

      If SendSQLToDB(lstSql) = True Then
        Return True
      Else
        SendMessageToLog("Update to WMS_T_POData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function DeleteWMS_T_PODataToDB(ByRef Info As List(Of clsPO_Back)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
        strSQL,
        TableName,
        IdxColumnName.PO_ID.ToString, CI.get_PO_ID
        )
        lstSql.Add(strSQL)
      Next

      If SendSQLToDB(lstSql) = True Then
        Return True
      Else
        SendMessageToLog("Delete WMS_T_POData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '-內部記憶體增刪修
  Private Shared Function AddOrUpdateWMS_T_PODataToDictionary(ByRef Info As List(Of clsPO_Back)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      For Each CI In Info
        Dim _Data As clsPO_Back = CI
        Dim key As String = _Data.get_gid()
        dicData.AddOrUpdate(key,
        _Data,
        Function(dicKey, ExistVal)
          UpdateInfo(dicKey, ExistVal, _Data)
          Return ExistVal
        End Function)
      Next

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  Private Shared Function DeleteWMS_T_PODataToDictionary(ByRef Info As List(Of clsPO_Back)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      For i = 0 To Info.Count - 1
        Dim key As String = Info(i).get_gid()
        If dicData.TryRemove(key, Nothing) = False Then

          SendMessageToLog("dicData.TryRemove Failed -WMS_T_POData", eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
      Next

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsPO_Back, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_TYPE_1 = "" & RowData.Item(IdxColumnName.PO_TYPE_1.ToString)
        Dim PO_TYPE_2 = "" & RowData.Item(IdxColumnName.PO_TYPE_2.ToString)
        Dim PO_TYPE_3 = "" & RowData.Item(IdxColumnName.PO_TYPE_3.ToString)
        Dim WO_TYPE = 0 & RowData.Item(IdxColumnName.WO_TYPE.ToString)
        Dim PRIORITY = 0 & RowData.Item(IdxColumnName.PRIORITY.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim START_TIME = "" & RowData.Item(IdxColumnName.START_TIME.ToString)
        Dim FINISH_TIME = "" & RowData.Item(IdxColumnName.FINISH_TIME.ToString)
        Dim USER_ID = "" & RowData.Item(IdxColumnName.USER_ID.ToString)
        Dim CUSTOMER_NO = "" & RowData.Item(IdxColumnName.CUSTOMER_NO.ToString)
        Dim CLASS_NO = "" & RowData.Item(IdxColumnName.CLASS_NO.ToString)
        Dim SHIPPING_NO = "" & RowData.Item(IdxColumnName.SHIPPING_NO.ToString)
        Dim OWNER_NO = "" & RowData.Item(IdxColumnName.OWNER_NO.ToString)
        Dim FACTORY_NO = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim DEST_AREA_NO = "" & RowData.Item(IdxColumnName.DEST_AREA_NO.ToString)
        Dim WRITE_OFF_NO = "" & RowData.Item(IdxColumnName.WRITE_OFF_NO.ToString)
        Dim PO_STATUS = 0 & RowData.Item(IdxColumnName.PO_STATUS.ToString)
        Dim HOST_CREATE_TIME = "" & RowData.Item(IdxColumnName.HOST_CREATE_TIME.ToString)
        Dim HOST_FINISH_TIME = "" & RowData.Item(IdxColumnName.HOST_FINISH_TIME.ToString)
        Dim HOST_STEP_NO = 0 & RowData.Item(IdxColumnName.HOST_STEP_NO.ToString)
        Dim HOST_OrderType = 0 & RowData.Item(IdxColumnName.HOST_OrderType.ToString)
        Dim HOST_CUSTOMER_NO = "" & RowData.Item(IdxColumnName.HOST_CUSTOMER_NO.ToString)
        Dim HOST_CUSTOMER_ID = "" & RowData.Item(IdxColumnName.HOST_CUSTOMER_ID.ToString)
        Dim HOST_CUSTOMER_COMMON1 = "" & RowData.Item(IdxColumnName.HOST_CUSTOMER_COMMON1.ToString)
        Dim HOST_CUSTOMER_COMMON2 = "" & RowData.Item(IdxColumnName.HOST_CUSTOMER_COMMON2.ToString)
        Dim HOST_CUSTOMER_COMMON3 = "" & RowData.Item(IdxColumnName.HOST_CUSTOMER_COMMON3.ToString)
        Dim HOST_CUSTOMER_COMMON4 = "" & RowData.Item(IdxColumnName.HOST_CUSTOMER_COMMON4.ToString)
        Dim HOST_CUSTOMER_COMMON5 = "" & RowData.Item(IdxColumnName.HOST_CUSTOMER_COMMON5.ToString)
        Dim HOST_OWNER_NO = "" & RowData.Item(IdxColumnName.HOST_OWNER_NO.ToString)
        Dim HOST_OWNER_ID = "" & RowData.Item(IdxColumnName.HOST_OWNER_ID.ToString)
        Dim HOST_OWNER_COMMON1 = "" & RowData.Item(IdxColumnName.HOST_OWNER_COMMON1.ToString)
        Dim HOST_OWNER_COMMON2 = "" & RowData.Item(IdxColumnName.HOST_OWNER_COMMON2.ToString)
        Dim HOST_OWNER_COMMON3 = "" & RowData.Item(IdxColumnName.HOST_OWNER_COMMON3.ToString)
        Dim HOST_OWNER_COMMON4 = "" & RowData.Item(IdxColumnName.HOST_OWNER_COMMON4.ToString)
        Dim HOST_OWNER_COMMON5 = "" & RowData.Item(IdxColumnName.HOST_OWNER_COMMON5.ToString)
        Dim HOST_COMMON1 = "" & RowData.Item(IdxColumnName.HOST_COMMON1.ToString)
        Dim HOST_COMMON2 = "" & RowData.Item(IdxColumnName.HOST_COMMON2.ToString)
        Dim HOST_COMMON3 = "" & RowData.Item(IdxColumnName.HOST_COMMON3.ToString)
        Dim HOST_COMMON4 = "" & RowData.Item(IdxColumnName.HOST_COMMON4.ToString)
        Dim HOST_COMMON5 = "" & RowData.Item(IdxColumnName.HOST_COMMON5.ToString)
        Dim HOST_COMMON6 = "" & RowData.Item(IdxColumnName.HOST_COMMON6.ToString)
        Dim HOST_COMMON7 = "" & RowData.Item(IdxColumnName.HOST_COMMON7.ToString)
        Dim HOST_COMMON8 = "" & RowData.Item(IdxColumnName.HOST_COMMON8.ToString)
        Dim HOST_COMMON9 = "" & RowData.Item(IdxColumnName.HOST_COMMON9.ToString)
        Dim HOST_COMMON10 = "" & RowData.Item(IdxColumnName.HOST_COMMON10.ToString)
        Dim HOST_COMMON11 = "" & RowData.Item(IdxColumnName.HOST_COMMON11.ToString)
        Dim HOST_COMMON12 = "" & RowData.Item(IdxColumnName.HOST_COMMON12.ToString)
        Dim HOST_COMMON13 = "" & RowData.Item(IdxColumnName.HOST_COMMON13.ToString)
        Dim HOST_COMMON14 = "" & RowData.Item(IdxColumnName.HOST_COMMON14.ToString)
        Dim HOST_COMMON15 = "" & RowData.Item(IdxColumnName.HOST_COMMON15.ToString)
        Dim HOST_COMMON16 = "" & RowData.Item(IdxColumnName.HOST_COMMON16.ToString)
        Dim HOST_COMMON17 = "" & RowData.Item(IdxColumnName.HOST_COMMON17.ToString)
        Dim HOST_COMMON18 = "" & RowData.Item(IdxColumnName.HOST_COMMON18.ToString)
        Dim HOST_COMMON19 = "" & RowData.Item(IdxColumnName.HOST_COMMON19.ToString)
        Dim HOST_COMMON20 = "" & RowData.Item(IdxColumnName.HOST_COMMON20.ToString)
        Dim HOST_COMMENTS = "" & RowData.Item(IdxColumnName.HOST_COMMENTS.ToString)
        Info = New clsPO_Back(PO_ID, PO_TYPE_1, PO_TYPE_2, PO_TYPE_3, PRIORITY, CREATE_TIME, START_TIME, FINISH_TIME, USER_ID, CUSTOMER_NO, CLASS_NO, SHIPPING_NO, OWNER_NO, FACTORY_NO, DEST_AREA_NO, PO_STATUS, WO_TYPE, WRITE_OFF_NO, HOST_CREATE_TIME, HOST_FINISH_TIME, HOST_STEP_NO, HOST_OrderType, HOST_CUSTOMER_NO, HOST_CUSTOMER_ID, HOST_CUSTOMER_COMMON1, HOST_CUSTOMER_COMMON2, HOST_CUSTOMER_COMMON3, HOST_CUSTOMER_COMMON4, HOST_CUSTOMER_COMMON5, HOST_OWNER_NO, HOST_OWNER_ID, HOST_OWNER_COMMON1, HOST_OWNER_COMMON2, HOST_OWNER_COMMON3, HOST_OWNER_COMMON4, HOST_OWNER_COMMON5, HOST_COMMON1, HOST_COMMON2, HOST_COMMON3, HOST_COMMON4, HOST_COMMON5, HOST_COMMON6, HOST_COMMON7, HOST_COMMON8, HOST_COMMON9, HOST_COMMON10, HOST_COMMON11, HOST_COMMON12, HOST_COMMON13, HOST_COMMON14, HOST_COMMON15, HOST_COMMON16, HOST_COMMON17, HOST_COMMON18, HOST_COMMON19, HOST_COMMON20, HOST_COMMENTS)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateInfo(ByRef Key As String, ByRef Info As clsPO_Back, ByRef objNewTC As clsPO_Back) As clsPO_Back
    Try
      If Key = Info.get_gid() Then
        Info.Update_To_Memory(objNewTC)

      Else
        SendMessageToLog("Dictionary has the different key", eCALogTool.ILogTool.enuTrcLevel.lvError)
      End If

    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
    Return Info
  End Function
  Private Shared Function SendSQLToDB(ByRef lstSQL As List(Of String)) As Boolean
    Try
      If lstSQL Is Nothing Then Return False
      If lstSQL.Count = 0 Then Return True
      For i = 0 To lstSQL.Count - 1
        SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next
      If fUseBatchUpdate = 0 Then
        For i = 0 To lstSQL.Count - 1
          DBTool.O_AddSQLQueue(TableName, lstSQL(i))
        Next
      Else
        Dim rtnMsg As String = DBTool.BatchUpdate(lstSQL)
        If rtnMsg.StartsWith("OK") Then
          SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        End If
      End If
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
End Class
