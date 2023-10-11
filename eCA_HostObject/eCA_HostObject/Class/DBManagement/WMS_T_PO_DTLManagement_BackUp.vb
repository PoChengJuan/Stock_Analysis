Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_T_PO_DTLManagement_BackUp
  Public Shared TableName As String = "WMS_T_PO_DTL"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsPO_DTL_Bak)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    PO_ID
    PO_LINE_NO
    PO_SERIAL_NO
    SKU_NO
    LOT_NO
    QTY
    QTY_PROCESS
    QTY_FINISH
    COMMENTS
    PACKAGE_ID
    ITEM_COMMON1
    ITEM_COMMON2
    ITEM_COMMON3
    ITEM_COMMON4
    ITEM_COMMON5
    ITEM_COMMON6
    ITEM_COMMON7
    ITEM_COMMON8
    ITEM_COMMON9
    ITEM_COMMON10
    HOST_STEP_NO
    HOST_MOVE_TYPE
    HOST_FINISH_TIME
    HOST_BILLING_DATE
    HOST_CREATE_TIME
    HOST_FACTORY_NO
    HOST_AREA_NO
    HOST_OWNER_NO
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
    HOST_COMMON21
    HOST_COMMON22
    HOST_COMMON23
    HOST_COMMON24
    HOST_COMMON25
    HOST_COMMON26
    HOST_COMMON27
    HOST_COMMON28
    HOST_COMMON29
    HOST_COMMON30
    HOST_COMMON31
    HOST_COMMON32
    HOST_COMMON33
    HOST_COMMON34
    HOST_COMMON35
    HOST_COMMON36
    HOST_COMMON37
    HOST_COMMON38
    HOST_COMMON39
    HOST_COMMON40
    HOST_COMMENTS
  End Enum

  '- GetSQL
  '-請將 clsPO_DTL 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsPO_DTL_Bak) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92},{94},{96},{98},{100},{102},{104},{106},{108},{110},{112},{114},{116},{118},{120},{122},{124},{126},{128},{130},{132},{134},{136},{138}) values ('{3}','{5}','{7}','{9}','{11}',{13},{15},{17},'{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}',{43},'{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}','{95}','{97}','{99}','{101}','{103}','{105}','{107}','{109}','{111}','{113}','{115}','{117}','{119}','{121}','{123}','{125}','{127}','{129}','{131}','{133}','{135}','{137}','{139}')",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.get_PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.get_PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.get_PO_SERIAL_NO,
      IdxColumnName.SKU_NO.ToString, Info.get_SKU_NO,
      IdxColumnName.LOT_NO.ToString, Info.get_LOT_NO,
      IdxColumnName.QTY.ToString, Info.get_QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.get_QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.get_QTY_FINISH,
      IdxColumnName.COMMENTS.ToString, Info.get_COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.get_PACKAGE_ID,
      IdxColumnName.ITEM_COMMON1.ToString, Info.get_ITEM_COMMON1,
      IdxColumnName.ITEM_COMMON2.ToString, Info.get_ITEM_COMMON2,
      IdxColumnName.ITEM_COMMON3.ToString, Info.get_ITEM_COMMON3,
      IdxColumnName.ITEM_COMMON4.ToString, Info.get_ITEM_COMMON4,
      IdxColumnName.ITEM_COMMON5.ToString, Info.get_ITEM_COMMON5,
      IdxColumnName.ITEM_COMMON6.ToString, Info.get_ITEM_COMMON6,
      IdxColumnName.ITEM_COMMON7.ToString, Info.get_ITEM_COMMON7,
      IdxColumnName.ITEM_COMMON8.ToString, Info.get_ITEM_COMMON8,
      IdxColumnName.ITEM_COMMON9.ToString, Info.get_ITEM_COMMON9,
      IdxColumnName.ITEM_COMMON10.ToString, Info.get_ITEM_COMMON10,
      IdxColumnName.HOST_STEP_NO.ToString, Info.get_HOST_STEP_NO,
      IdxColumnName.HOST_MOVE_TYPE.ToString, Info.get_HOST_MOVE_TYPE,
      IdxColumnName.HOST_FINISH_TIME.ToString, Info.get_HOST_FINISH_TIME,
      IdxColumnName.HOST_BILLING_DATE.ToString, Info.get_HOST_BILLING_DATE,
      IdxColumnName.HOST_CREATE_TIME.ToString, Info.get_HOST_CREATE_TIME,
      IdxColumnName.HOST_FACTORY_NO.ToString, Info.get_HOST_FACTORY_NO,
      IdxColumnName.HOST_AREA_NO.ToString, Info.get_HOST_AREA_NO,
      IdxColumnName.HOST_OWNER_NO.ToString, Info.get_HOST_OWNER_NO,
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
      IdxColumnName.HOST_COMMON21.ToString, Info.get_HOST_COMMON21,
      IdxColumnName.HOST_COMMON22.ToString, Info.get_HOST_COMMON22,
      IdxColumnName.HOST_COMMON23.ToString, Info.get_HOST_COMMON23,
      IdxColumnName.HOST_COMMON24.ToString, Info.get_HOST_COMMON24,
      IdxColumnName.HOST_COMMON25.ToString, Info.get_HOST_COMMON25,
      IdxColumnName.HOST_COMMON26.ToString, Info.get_HOST_COMMON26,
      IdxColumnName.HOST_COMMON27.ToString, Info.get_HOST_COMMON27,
      IdxColumnName.HOST_COMMON28.ToString, Info.get_HOST_COMMON28,
      IdxColumnName.HOST_COMMON29.ToString, Info.get_HOST_COMMON29,
      IdxColumnName.HOST_COMMON30.ToString, Info.get_HOST_COMMON30,
      IdxColumnName.HOST_COMMON31.ToString, Info.get_HOST_COMMON31,
      IdxColumnName.HOST_COMMON32.ToString, Info.get_HOST_COMMON32,
      IdxColumnName.HOST_COMMON33.ToString, Info.get_HOST_COMMON33,
      IdxColumnName.HOST_COMMON34.ToString, Info.get_HOST_COMMON34,
      IdxColumnName.HOST_COMMON35.ToString, Info.get_HOST_COMMON35,
      IdxColumnName.HOST_COMMON36.ToString, Info.get_HOST_COMMON36,
      IdxColumnName.HOST_COMMON37.ToString, Info.get_HOST_COMMON37,
      IdxColumnName.HOST_COMMON38.ToString, Info.get_HOST_COMMON38,
      IdxColumnName.HOST_COMMON39.ToString, Info.get_HOST_COMMON39,
      IdxColumnName.HOST_COMMON40.ToString, Info.get_HOST_COMMON40,
      IdxColumnName.HOST_COMMENTS.ToString, Info.get_HOST_COMMENTS
     )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsPO_DTL_Bak) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {6}='{7}' ",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.get_PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.get_PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.get_PO_SERIAL_NO,
      IdxColumnName.SKU_NO.ToString, Info.get_SKU_NO,
      IdxColumnName.LOT_NO.ToString, Info.get_LOT_NO,
      IdxColumnName.QTY.ToString, Info.get_QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.get_QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.get_QTY_FINISH,
      IdxColumnName.COMMENTS.ToString, Info.get_COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.get_PACKAGE_ID,
      IdxColumnName.ITEM_COMMON1.ToString, Info.get_ITEM_COMMON1,
      IdxColumnName.ITEM_COMMON2.ToString, Info.get_ITEM_COMMON2,
      IdxColumnName.ITEM_COMMON3.ToString, Info.get_ITEM_COMMON3,
      IdxColumnName.ITEM_COMMON4.ToString, Info.get_ITEM_COMMON4,
      IdxColumnName.ITEM_COMMON5.ToString, Info.get_ITEM_COMMON5,
      IdxColumnName.ITEM_COMMON6.ToString, Info.get_ITEM_COMMON6,
      IdxColumnName.ITEM_COMMON7.ToString, Info.get_ITEM_COMMON7,
      IdxColumnName.ITEM_COMMON8.ToString, Info.get_ITEM_COMMON8,
      IdxColumnName.ITEM_COMMON9.ToString, Info.get_ITEM_COMMON9,
      IdxColumnName.ITEM_COMMON10.ToString, Info.get_ITEM_COMMON10,
      IdxColumnName.HOST_STEP_NO.ToString, Info.get_HOST_STEP_NO,
      IdxColumnName.HOST_MOVE_TYPE.ToString, Info.get_HOST_MOVE_TYPE,
      IdxColumnName.HOST_FINISH_TIME.ToString, Info.get_HOST_FINISH_TIME,
      IdxColumnName.HOST_BILLING_DATE.ToString, Info.get_HOST_BILLING_DATE,
      IdxColumnName.HOST_CREATE_TIME.ToString, Info.get_HOST_CREATE_TIME,
      IdxColumnName.HOST_FACTORY_NO.ToString, Info.get_HOST_FACTORY_NO,
      IdxColumnName.HOST_AREA_NO.ToString, Info.get_HOST_AREA_NO,
      IdxColumnName.HOST_OWNER_NO.ToString, Info.get_HOST_OWNER_NO,
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
      IdxColumnName.HOST_COMMON21.ToString, Info.get_HOST_COMMON21,
      IdxColumnName.HOST_COMMON22.ToString, Info.get_HOST_COMMON22,
      IdxColumnName.HOST_COMMON23.ToString, Info.get_HOST_COMMON23,
      IdxColumnName.HOST_COMMON24.ToString, Info.get_HOST_COMMON24,
      IdxColumnName.HOST_COMMON25.ToString, Info.get_HOST_COMMON25,
      IdxColumnName.HOST_COMMON26.ToString, Info.get_HOST_COMMON26,
      IdxColumnName.HOST_COMMON27.ToString, Info.get_HOST_COMMON27,
      IdxColumnName.HOST_COMMON28.ToString, Info.get_HOST_COMMON28,
      IdxColumnName.HOST_COMMON29.ToString, Info.get_HOST_COMMON29,
      IdxColumnName.HOST_COMMON30.ToString, Info.get_HOST_COMMON30,
      IdxColumnName.HOST_COMMON31.ToString, Info.get_HOST_COMMON31,
      IdxColumnName.HOST_COMMON32.ToString, Info.get_HOST_COMMON32,
      IdxColumnName.HOST_COMMON33.ToString, Info.get_HOST_COMMON33,
      IdxColumnName.HOST_COMMON34.ToString, Info.get_HOST_COMMON34,
      IdxColumnName.HOST_COMMON35.ToString, Info.get_HOST_COMMON35,
      IdxColumnName.HOST_COMMON36.ToString, Info.get_HOST_COMMON36,
      IdxColumnName.HOST_COMMON37.ToString, Info.get_HOST_COMMON37,
      IdxColumnName.HOST_COMMON38.ToString, Info.get_HOST_COMMON38,
      IdxColumnName.HOST_COMMON39.ToString, Info.get_HOST_COMMON39,
      IdxColumnName.HOST_COMMON40.ToString, Info.get_HOST_COMMON40,
      IdxColumnName.HOST_COMMENTS.ToString, Info.get_HOST_COMMENTS
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsPO_DTL_Bak) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{8}='{9}',{10}='{11}',{12}={13},{14}={15},{16}={17},{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}={43},{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}',{94}='{95}',{96}='{97}',{98}='{99}',{100}='{101}',{102}='{103}',{104}='{105}',{106}='{107}',{108}='{109}',{110}='{111}',{112}='{113}',{114}='{115}',{116}='{117}',{118}='{119}',{120}='{121}',{122}='{123}',{124}='{125}',{126}='{127}',{128}='{129}',{130}='{131}',{132}='{133}',{134}='{135}',{136}='{137}',{138}='{139}' WHERE {2}='{3}' And {6}='{7}'",
      strSQL,
      TableName,
      IdxColumnName.PO_ID.ToString, Info.get_PO_ID,
      IdxColumnName.PO_LINE_NO.ToString, Info.get_PO_LINE_NO,
      IdxColumnName.PO_SERIAL_NO.ToString, Info.get_PO_SERIAL_NO,
      IdxColumnName.SKU_NO.ToString, Info.get_SKU_NO,
      IdxColumnName.LOT_NO.ToString, Info.get_LOT_NO,
      IdxColumnName.QTY.ToString, Info.get_QTY,
      IdxColumnName.QTY_PROCESS.ToString, Info.get_QTY_PROCESS,
      IdxColumnName.QTY_FINISH.ToString, Info.get_QTY_FINISH,
      IdxColumnName.COMMENTS.ToString, Info.get_COMMENTS,
      IdxColumnName.PACKAGE_ID.ToString, Info.get_PACKAGE_ID,
      IdxColumnName.ITEM_COMMON1.ToString, Info.get_ITEM_COMMON1,
      IdxColumnName.ITEM_COMMON2.ToString, Info.get_ITEM_COMMON2,
      IdxColumnName.ITEM_COMMON3.ToString, Info.get_ITEM_COMMON3,
      IdxColumnName.ITEM_COMMON4.ToString, Info.get_ITEM_COMMON4,
      IdxColumnName.ITEM_COMMON5.ToString, Info.get_ITEM_COMMON5,
      IdxColumnName.ITEM_COMMON6.ToString, Info.get_ITEM_COMMON6,
      IdxColumnName.ITEM_COMMON7.ToString, Info.get_ITEM_COMMON7,
      IdxColumnName.ITEM_COMMON8.ToString, Info.get_ITEM_COMMON8,
      IdxColumnName.ITEM_COMMON9.ToString, Info.get_ITEM_COMMON9,
      IdxColumnName.ITEM_COMMON10.ToString, Info.get_ITEM_COMMON10,
      IdxColumnName.HOST_STEP_NO.ToString, Info.get_HOST_STEP_NO,
      IdxColumnName.HOST_MOVE_TYPE.ToString, Info.get_HOST_MOVE_TYPE,
      IdxColumnName.HOST_FINISH_TIME.ToString, Info.get_HOST_FINISH_TIME,
      IdxColumnName.HOST_BILLING_DATE.ToString, Info.get_HOST_BILLING_DATE,
      IdxColumnName.HOST_CREATE_TIME.ToString, Info.get_HOST_CREATE_TIME,
      IdxColumnName.HOST_FACTORY_NO.ToString, Info.get_HOST_FACTORY_NO,
      IdxColumnName.HOST_AREA_NO.ToString, Info.get_HOST_AREA_NO,
      IdxColumnName.HOST_OWNER_NO.ToString, Info.get_HOST_OWNER_NO,
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
      IdxColumnName.HOST_COMMON21.ToString, Info.get_HOST_COMMON21,
      IdxColumnName.HOST_COMMON22.ToString, Info.get_HOST_COMMON22,
      IdxColumnName.HOST_COMMON23.ToString, Info.get_HOST_COMMON23,
      IdxColumnName.HOST_COMMON24.ToString, Info.get_HOST_COMMON24,
      IdxColumnName.HOST_COMMON25.ToString, Info.get_HOST_COMMON25,
      IdxColumnName.HOST_COMMON26.ToString, Info.get_HOST_COMMON26,
      IdxColumnName.HOST_COMMON27.ToString, Info.get_HOST_COMMON27,
      IdxColumnName.HOST_COMMON28.ToString, Info.get_HOST_COMMON28,
      IdxColumnName.HOST_COMMON29.ToString, Info.get_HOST_COMMON29,
      IdxColumnName.HOST_COMMON30.ToString, Info.get_HOST_COMMON30,
      IdxColumnName.HOST_COMMON31.ToString, Info.get_HOST_COMMON31,
      IdxColumnName.HOST_COMMON32.ToString, Info.get_HOST_COMMON32,
      IdxColumnName.HOST_COMMON33.ToString, Info.get_HOST_COMMON33,
      IdxColumnName.HOST_COMMON34.ToString, Info.get_HOST_COMMON34,
      IdxColumnName.HOST_COMMON35.ToString, Info.get_HOST_COMMON35,
      IdxColumnName.HOST_COMMON36.ToString, Info.get_HOST_COMMON36,
      IdxColumnName.HOST_COMMON37.ToString, Info.get_HOST_COMMON37,
      IdxColumnName.HOST_COMMON38.ToString, Info.get_HOST_COMMON38,
      IdxColumnName.HOST_COMMON39.ToString, Info.get_HOST_COMMON39,
      IdxColumnName.HOST_COMMON40.ToString, Info.get_HOST_COMMON40,
      IdxColumnName.HOST_COMMENTS.ToString, Info.get_HOST_COMMENTS
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '- Add & Insert
  Public Shared Function AddWMS_T_PO_DTLData(ByVal Info As clsPO_DTL_Bak, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If AddlstWMS_T_PO_DTLData(New List(Of clsPO_DTL_Bak)({Info}), SendToDB) = True Then
          Return True
        End If '-載不載入記憶體都是呼叫同一個function
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function AddlstWMS_T_PO_DTLData(ByVal Info As List(Of clsPO_DTL_Bak), Optional ByVal SendToDB As Boolean = True) As Boolean
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
            If InsertWMS_T_PO_DTLDataToDB(Info) Then
              If AddOrUpdateWMS_T_PO_DTLDataToDictionary(Info) Then
                SendMessageToLog("InsertDic WMS_T_PO_DTLData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("InsertDic WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("InsertDB WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If AddOrUpdateWMS_T_PO_DTLDataToDictionary(Info) Then
              SendMessageToLog("InsertDic WMS_T_PO_DTLData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("InsertDic WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If InsertWMS_T_PO_DTLDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("InsertDic WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
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
  Public Shared Function UpdateWMS_T_PO_DTLData(ByVal Info As clsPO_DTL_Bak, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If UpdatelstWMS_T_PO_DTLData(New List(Of clsPO_DTL_Bak)({Info}), SendToDB) = True Then
          Return True
        End If
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function UpdatelstWMS_T_PO_DTLData(ByVal Info As List(Of clsPO_DTL_Bak), Optional ByVal SendToDB As Boolean = True) As Boolean
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
            If UpdateWMS_T_PO_DTLDataToDB(Info) Then
              If AddOrUpdateWMS_T_PO_DTLDataToDictionary(Info) Then
                SendMessageToLog("UpdateDic WMS_T_PO_DTLData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("UpdateDic WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("UpdateDB WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If AddOrUpdateWMS_T_PO_DTLDataToDictionary(Info) Then
              SendMessageToLog("UpdateDic WMS_T_PO_DTLData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("UpdateDic WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
        Else
          If SendToDB Then
            If UpdateWMS_T_PO_DTLDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("UpdateDB WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
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
  Public Shared Function DeleteWMS_T_PO_DTLData(ByVal Info As clsPO_DTL_Bak, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If DeletelstWMS_T_PO_DTLData(New List(Of clsPO_DTL_Bak)({Info}), SendToDB) = True Then
          Return True
        End If '-載不載入記憶體都是呼叫同一個function
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function DeletelstWMS_T_PO_DTLData(ByVal Info As List(Of clsPO_DTL_Bak), Optional ByVal SendToDB As Boolean = True) As Boolean
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
            If DeleteWMS_T_PO_DTLDataToDB(Info) Then
              If DeleteWMS_T_PO_DTLDataToDictionary(Info) Then
                SendMessageToLog("DeleteDic WMS_T_PO_DTLData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Else
                SendMessageToLog("DeleteDic WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
              End If
            Else
              SendMessageToLog("DeleteDB WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          Else
            If DeleteWMS_T_PO_DTLDataToDB(Info) Then
              SendMessageToLog("DeleteDic WMS_T_PO_DTLData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Else
              SendMessageToLog("DeleteDB WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
          Return True
        Else
          If SendToDB Then
            If DeleteWMS_T_PO_DTLDataToDB(Info) Then
              Return True
            Else
              SendMessageToLog("DeleteDB WMS_T_PO_DTLData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
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
  Public Shared Function GetWMS_T_PO_DTLDataListByALL() As List(Of clsPO_DTL_Bak)
    SyncLock objLock
      Try
        Dim _lstReturn As New List(Of clsPO_DTL_Bak)
        If DictionaryNeeded = 1 Then '-載入記憶體
          Dim LinqFind As IEnumerable(Of clsPO_DTL_Bak) = From TC In dicData Select TC.Value
          '- From TC In dicData Where TC.Value.xxx = xxx AND TC.Value.xxx = xxx AND TC.Value.xxx = xxx Select TC.Value '-範例
          For Each objTC As clsPO_DTL_Bak In LinqFind
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
                  Dim Info As clsPO_DTL_Bak = Nothing
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
  Public Shared Function GetclsPO_DTLListByPO_ID_PO_SERIAL_NO(ByVal po_id As String, po_serial_no As String) As List(Of clsPO_DTL_Bak)
    SyncLock objLock
      Try
        Dim _lstReturn As New List(Of clsPO_DTL_Bak)
        If DictionaryNeeded = 1 Then '-載入記憶體
          Dim strSQL As String = String.Empty
          Dim rs As ADODB.Recordset = Nothing

          Dim LinqFind As IEnumerable(Of clsPO_DTL_Bak) = From TC In dicData Where TC.Value.get_PO_ID = po_id And TC.Value.get_PO_SERIAL_NO = po_serial_no Select TC.Value
          '- From TC In dicData Where TC.Value.xxx = xxx AND TC.Value.xxx = xxx AND TC.Value.xxx = xxx Select TC.Value '-範例
          For Each objTC As clsPO_DTL_Bak In LinqFind
            _lstReturn.Add(objTC)
          Next

          Return _lstReturn
        Else
          If DBTool IsNot Nothing Then
            If DBTool.isConnection(DBTool.m_CN) = True Then
              Dim strSQL As String = String.Empty
              Dim rs As ADODB.Recordset = Nothing

              strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4} = '{5}' ",
              strSQL,
              TableName,
              IdxColumnName.PO_ID.ToString, po_id,
              IdxColumnName.PO_SERIAL_NO.ToString, po_serial_no
              )
              SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
              DBTool.SQLExcute(strSQL, rs)
              Dim DatasetMessage As New DataSet
              Dim OLEDBAdapter As New OleDbDataAdapter
              OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

              If DatasetMessage.Tables(TableName).Rows.Count > 0 Then
                For RowIndex = 0 To DatasetMessage.Tables(TableName).Rows.Count - 1
                  Dim Info As clsPO_DTL_Bak = Nothing
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
  Private Shared Function InsertWMS_T_PO_DTLDataToDB(ByRef Info As List(Of clsPO_DTL_Bak)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88},{90},{92},{94},{96},{98},{100},{102},{104},{106},{108},{110},{112},{114},{116},{118},{120},{122},{124},{126},{128},{130},{132},{134},{136},{138}) values ('{3}','{5}','{7}','{9}','{11}',{13},{15},{17},'{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}',{43},'{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}','{79}','{81}','{83}','{85}','{87}','{89}','{91}','{93}','{95}','{97}','{99}','{101}','{103}','{105}','{107}','{109}','{111}','{113}','{115}','{117}','{119}','{121}','{123}','{125}','{127}','{129}','{131}','{133}','{135}','{137}','{139}')",
        strSQL,
        TableName,
        IdxColumnName.PO_ID.ToString, CI.get_PO_ID,
        IdxColumnName.PO_LINE_NO.ToString, CI.get_PO_LINE_NO,
        IdxColumnName.PO_SERIAL_NO.ToString, CI.get_PO_SERIAL_NO,
        IdxColumnName.SKU_NO.ToString, CI.get_SKU_NO,
        IdxColumnName.LOT_NO.ToString, CI.get_LOT_NO,
        IdxColumnName.QTY.ToString, CI.get_QTY,
        IdxColumnName.QTY_PROCESS.ToString, CI.get_QTY_PROCESS,
        IdxColumnName.QTY_FINISH.ToString, CI.get_QTY_FINISH,
        IdxColumnName.COMMENTS.ToString, CI.get_COMMENTS,
        IdxColumnName.PACKAGE_ID.ToString, CI.get_PACKAGE_ID,
        IdxColumnName.ITEM_COMMON1.ToString, CI.get_ITEM_COMMON1,
        IdxColumnName.ITEM_COMMON2.ToString, CI.get_ITEM_COMMON2,
        IdxColumnName.ITEM_COMMON3.ToString, CI.get_ITEM_COMMON3,
        IdxColumnName.ITEM_COMMON4.ToString, CI.get_ITEM_COMMON4,
        IdxColumnName.ITEM_COMMON5.ToString, CI.get_ITEM_COMMON5,
        IdxColumnName.ITEM_COMMON6.ToString, CI.get_ITEM_COMMON6,
        IdxColumnName.ITEM_COMMON7.ToString, CI.get_ITEM_COMMON7,
        IdxColumnName.ITEM_COMMON8.ToString, CI.get_ITEM_COMMON8,
        IdxColumnName.ITEM_COMMON9.ToString, CI.get_ITEM_COMMON9,
        IdxColumnName.ITEM_COMMON10.ToString, CI.get_ITEM_COMMON10,
        IdxColumnName.HOST_STEP_NO.ToString, CI.get_HOST_STEP_NO,
        IdxColumnName.HOST_MOVE_TYPE.ToString, CI.get_HOST_MOVE_TYPE,
        IdxColumnName.HOST_FINISH_TIME.ToString, CI.get_HOST_FINISH_TIME,
        IdxColumnName.HOST_BILLING_DATE.ToString, CI.get_HOST_BILLING_DATE,
        IdxColumnName.HOST_CREATE_TIME.ToString, CI.get_HOST_CREATE_TIME,
        IdxColumnName.HOST_FACTORY_NO.ToString, CI.get_HOST_FACTORY_NO,
        IdxColumnName.HOST_AREA_NO.ToString, CI.get_HOST_AREA_NO,
        IdxColumnName.HOST_OWNER_NO.ToString, CI.get_HOST_OWNER_NO,
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
        IdxColumnName.HOST_COMMON21.ToString, CI.get_HOST_COMMON21,
        IdxColumnName.HOST_COMMON22.ToString, CI.get_HOST_COMMON22,
        IdxColumnName.HOST_COMMON23.ToString, CI.get_HOST_COMMON23,
        IdxColumnName.HOST_COMMON24.ToString, CI.get_HOST_COMMON24,
        IdxColumnName.HOST_COMMON25.ToString, CI.get_HOST_COMMON25,
        IdxColumnName.HOST_COMMON26.ToString, CI.get_HOST_COMMON26,
        IdxColumnName.HOST_COMMON27.ToString, CI.get_HOST_COMMON27,
        IdxColumnName.HOST_COMMON28.ToString, CI.get_HOST_COMMON28,
        IdxColumnName.HOST_COMMON29.ToString, CI.get_HOST_COMMON29,
        IdxColumnName.HOST_COMMON30.ToString, CI.get_HOST_COMMON30,
        IdxColumnName.HOST_COMMON31.ToString, CI.get_HOST_COMMON31,
        IdxColumnName.HOST_COMMON32.ToString, CI.get_HOST_COMMON32,
        IdxColumnName.HOST_COMMON33.ToString, CI.get_HOST_COMMON33,
        IdxColumnName.HOST_COMMON34.ToString, CI.get_HOST_COMMON34,
        IdxColumnName.HOST_COMMON35.ToString, CI.get_HOST_COMMON35,
        IdxColumnName.HOST_COMMON36.ToString, CI.get_HOST_COMMON36,
        IdxColumnName.HOST_COMMON37.ToString, CI.get_HOST_COMMON37,
        IdxColumnName.HOST_COMMON38.ToString, CI.get_HOST_COMMON38,
        IdxColumnName.HOST_COMMON39.ToString, CI.get_HOST_COMMON39,
        IdxColumnName.HOST_COMMON40.ToString, CI.get_HOST_COMMON40,
        IdxColumnName.HOST_COMMENTS.ToString, CI.get_HOST_COMMENTS
        )
        lstSql.Add(strSQL)
      Next
      If SendSQLToDB(lstSql) = True Then
        Return True
      Else
        SendMessageToLog("Insert to WMS_T_PO_DTLData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateWMS_T_PO_DTLDataToDB(ByRef Info As List(Of clsPO_DTL_Bak)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Update {1} SET {4}='{5}',{8}='{9}',{10}='{11}',{12}={13},{14}={15},{16}={17},{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}={43},{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}',{90}='{91}',{92}='{93}',{94}='{95}',{96}='{97}',{98}='{99}',{100}='{101}',{102}='{103}',{104}='{105}',{106}='{107}',{108}='{109}',{110}='{111}',{112}='{113}',{114}='{115}',{116}='{117}',{118}='{119}',{120}='{121}',{122}='{123}',{124}='{125}',{126}='{127}',{128}='{129}',{130}='{131}',{132}='{133}',{134}='{135}',{136}='{137}',{138}='{139}' WHERE {2}='{3}' And {6}='{7}'",
        strSQL,
        TableName,
        IdxColumnName.PO_ID.ToString, CI.get_PO_ID,
        IdxColumnName.PO_LINE_NO.ToString, CI.get_PO_LINE_NO,
        IdxColumnName.PO_SERIAL_NO.ToString, CI.get_PO_SERIAL_NO,
        IdxColumnName.SKU_NO.ToString, CI.get_SKU_NO,
        IdxColumnName.LOT_NO.ToString, CI.get_LOT_NO,
        IdxColumnName.QTY.ToString, CI.get_QTY,
        IdxColumnName.QTY_PROCESS.ToString, CI.get_QTY_PROCESS,
        IdxColumnName.QTY_FINISH.ToString, CI.get_QTY_FINISH,
        IdxColumnName.COMMENTS.ToString, CI.get_COMMENTS,
        IdxColumnName.PACKAGE_ID.ToString, CI.get_PACKAGE_ID,
        IdxColumnName.ITEM_COMMON1.ToString, CI.get_ITEM_COMMON1,
        IdxColumnName.ITEM_COMMON2.ToString, CI.get_ITEM_COMMON2,
        IdxColumnName.ITEM_COMMON3.ToString, CI.get_ITEM_COMMON3,
        IdxColumnName.ITEM_COMMON4.ToString, CI.get_ITEM_COMMON4,
        IdxColumnName.ITEM_COMMON5.ToString, CI.get_ITEM_COMMON5,
        IdxColumnName.ITEM_COMMON6.ToString, CI.get_ITEM_COMMON6,
        IdxColumnName.ITEM_COMMON7.ToString, CI.get_ITEM_COMMON7,
        IdxColumnName.ITEM_COMMON8.ToString, CI.get_ITEM_COMMON8,
        IdxColumnName.ITEM_COMMON9.ToString, CI.get_ITEM_COMMON9,
        IdxColumnName.ITEM_COMMON10.ToString, CI.get_ITEM_COMMON10,
        IdxColumnName.HOST_STEP_NO.ToString, CI.get_HOST_STEP_NO,
        IdxColumnName.HOST_MOVE_TYPE.ToString, CI.get_HOST_MOVE_TYPE,
        IdxColumnName.HOST_FINISH_TIME.ToString, CI.get_HOST_FINISH_TIME,
        IdxColumnName.HOST_BILLING_DATE.ToString, CI.get_HOST_BILLING_DATE,
        IdxColumnName.HOST_CREATE_TIME.ToString, CI.get_HOST_CREATE_TIME,
        IdxColumnName.HOST_FACTORY_NO.ToString, CI.get_HOST_FACTORY_NO,
        IdxColumnName.HOST_AREA_NO.ToString, CI.get_HOST_AREA_NO,
        IdxColumnName.HOST_OWNER_NO.ToString, CI.get_HOST_OWNER_NO,
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
        IdxColumnName.HOST_COMMON21.ToString, CI.get_HOST_COMMON21,
        IdxColumnName.HOST_COMMON22.ToString, CI.get_HOST_COMMON22,
        IdxColumnName.HOST_COMMON23.ToString, CI.get_HOST_COMMON23,
        IdxColumnName.HOST_COMMON24.ToString, CI.get_HOST_COMMON24,
        IdxColumnName.HOST_COMMON25.ToString, CI.get_HOST_COMMON25,
        IdxColumnName.HOST_COMMON26.ToString, CI.get_HOST_COMMON26,
        IdxColumnName.HOST_COMMON27.ToString, CI.get_HOST_COMMON27,
        IdxColumnName.HOST_COMMON28.ToString, CI.get_HOST_COMMON28,
        IdxColumnName.HOST_COMMON29.ToString, CI.get_HOST_COMMON29,
        IdxColumnName.HOST_COMMON30.ToString, CI.get_HOST_COMMON30,
        IdxColumnName.HOST_COMMON31.ToString, CI.get_HOST_COMMON31,
        IdxColumnName.HOST_COMMON32.ToString, CI.get_HOST_COMMON32,
        IdxColumnName.HOST_COMMON33.ToString, CI.get_HOST_COMMON33,
        IdxColumnName.HOST_COMMON34.ToString, CI.get_HOST_COMMON34,
        IdxColumnName.HOST_COMMON35.ToString, CI.get_HOST_COMMON35,
        IdxColumnName.HOST_COMMON36.ToString, CI.get_HOST_COMMON36,
        IdxColumnName.HOST_COMMON37.ToString, CI.get_HOST_COMMON37,
        IdxColumnName.HOST_COMMON38.ToString, CI.get_HOST_COMMON38,
        IdxColumnName.HOST_COMMON39.ToString, CI.get_HOST_COMMON39,
        IdxColumnName.HOST_COMMON40.ToString, CI.get_HOST_COMMON40,
        IdxColumnName.HOST_COMMENTS.ToString, CI.get_HOST_COMMENTS
        )
        lstSql.Add(strSQL)
      Next

      If SendSQLToDB(lstSql) = True Then
        Return True
      Else
        SendMessageToLog("Update to WMS_T_PO_DTLData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function DeleteWMS_T_PO_DTLDataToDB(ByRef Info As List(Of clsPO_DTL_Bak)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As ADODB.Recordset = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {6}='{7}' ",
        strSQL,
        TableName,
        IdxColumnName.PO_ID.ToString, CI.get_PO_ID,
        IdxColumnName.PO_LINE_NO.ToString, CI.get_PO_LINE_NO,
        IdxColumnName.PO_SERIAL_NO.ToString, CI.get_PO_SERIAL_NO,
        IdxColumnName.SKU_NO.ToString, CI.get_SKU_NO,
        IdxColumnName.LOT_NO.ToString, CI.get_LOT_NO,
        IdxColumnName.QTY.ToString, CI.get_QTY,
        IdxColumnName.QTY_PROCESS.ToString, CI.get_QTY_PROCESS,
        IdxColumnName.QTY_FINISH.ToString, CI.get_QTY_FINISH,
        IdxColumnName.COMMENTS.ToString, CI.get_COMMENTS,
        IdxColumnName.PACKAGE_ID.ToString, CI.get_PACKAGE_ID,
        IdxColumnName.ITEM_COMMON1.ToString, CI.get_ITEM_COMMON1,
        IdxColumnName.ITEM_COMMON2.ToString, CI.get_ITEM_COMMON2,
        IdxColumnName.ITEM_COMMON3.ToString, CI.get_ITEM_COMMON3,
        IdxColumnName.ITEM_COMMON4.ToString, CI.get_ITEM_COMMON4,
        IdxColumnName.ITEM_COMMON5.ToString, CI.get_ITEM_COMMON5,
        IdxColumnName.ITEM_COMMON6.ToString, CI.get_ITEM_COMMON6,
        IdxColumnName.ITEM_COMMON7.ToString, CI.get_ITEM_COMMON7,
        IdxColumnName.ITEM_COMMON8.ToString, CI.get_ITEM_COMMON8,
        IdxColumnName.ITEM_COMMON9.ToString, CI.get_ITEM_COMMON9,
        IdxColumnName.ITEM_COMMON10.ToString, CI.get_ITEM_COMMON10,
        IdxColumnName.HOST_STEP_NO.ToString, CI.get_HOST_STEP_NO,
        IdxColumnName.HOST_MOVE_TYPE.ToString, CI.get_HOST_MOVE_TYPE,
        IdxColumnName.HOST_FINISH_TIME.ToString, CI.get_HOST_FINISH_TIME,
        IdxColumnName.HOST_BILLING_DATE.ToString, CI.get_HOST_BILLING_DATE,
        IdxColumnName.HOST_CREATE_TIME.ToString, CI.get_HOST_CREATE_TIME,
        IdxColumnName.HOST_FACTORY_NO.ToString, CI.get_HOST_FACTORY_NO,
        IdxColumnName.HOST_AREA_NO.ToString, CI.get_HOST_AREA_NO,
        IdxColumnName.HOST_OWNER_NO.ToString, CI.get_HOST_OWNER_NO,
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
        IdxColumnName.HOST_COMMON21.ToString, CI.get_HOST_COMMON21,
        IdxColumnName.HOST_COMMON22.ToString, CI.get_HOST_COMMON22,
        IdxColumnName.HOST_COMMON23.ToString, CI.get_HOST_COMMON23,
        IdxColumnName.HOST_COMMON24.ToString, CI.get_HOST_COMMON24,
        IdxColumnName.HOST_COMMON25.ToString, CI.get_HOST_COMMON25,
        IdxColumnName.HOST_COMMON26.ToString, CI.get_HOST_COMMON26,
        IdxColumnName.HOST_COMMON27.ToString, CI.get_HOST_COMMON27,
        IdxColumnName.HOST_COMMON28.ToString, CI.get_HOST_COMMON28,
        IdxColumnName.HOST_COMMON29.ToString, CI.get_HOST_COMMON29,
        IdxColumnName.HOST_COMMON30.ToString, CI.get_HOST_COMMON30,
        IdxColumnName.HOST_COMMON31.ToString, CI.get_HOST_COMMON31,
        IdxColumnName.HOST_COMMON32.ToString, CI.get_HOST_COMMON32,
        IdxColumnName.HOST_COMMON33.ToString, CI.get_HOST_COMMON33,
        IdxColumnName.HOST_COMMON34.ToString, CI.get_HOST_COMMON34,
        IdxColumnName.HOST_COMMON35.ToString, CI.get_HOST_COMMON35,
        IdxColumnName.HOST_COMMON36.ToString, CI.get_HOST_COMMON36,
        IdxColumnName.HOST_COMMON37.ToString, CI.get_HOST_COMMON37,
        IdxColumnName.HOST_COMMON38.ToString, CI.get_HOST_COMMON38,
        IdxColumnName.HOST_COMMON39.ToString, CI.get_HOST_COMMON39,
        IdxColumnName.HOST_COMMON40.ToString, CI.get_HOST_COMMON40,
        IdxColumnName.HOST_COMMENTS.ToString, CI.get_HOST_COMMENTS
        )
        lstSql.Add(strSQL)
      Next

      If SendSQLToDB(lstSql) = True Then
        Return True
      Else
        SendMessageToLog("Delete WMS_T_PO_DTLData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '-內部記憶體增刪修
  Private Shared Function AddOrUpdateWMS_T_PO_DTLDataToDictionary(ByRef Info As List(Of clsPO_DTL_Bak)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      For Each CI In Info
        Dim _Data As clsPO_DTL_Bak = CI
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
  Private Shared Function DeleteWMS_T_PO_DTLDataToDictionary(ByRef Info As List(Of clsPO_DTL_Bak)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      For i = 0 To Info.Count - 1
        Dim key As String = Info(i).get_gid()
        If dicData.TryRemove(key, Nothing) = False Then

          SendMessageToLog("dicData.TryRemove Failed -WMS_T_PO_DTLData", eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
      Next

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsPO_DTL_Bak, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim PO_LINE_NO = "" & RowData.Item(IdxColumnName.PO_LINE_NO.ToString)
        Dim PO_SERIAL_NO = "" & RowData.Item(IdxColumnName.PO_SERIAL_NO.ToString)
        Dim SKU_NO = "" & RowData.Item(IdxColumnName.SKU_NO.ToString)
        Dim LOT_NO = "" & RowData.Item(IdxColumnName.LOT_NO.ToString)
        Dim QTY = 0 & RowData.Item(IdxColumnName.QTY.ToString)
        Dim QTY_PROCESS = 0 & RowData.Item(IdxColumnName.QTY_PROCESS.ToString)
        Dim QTY_FINISH = 0 & RowData.Item(IdxColumnName.QTY_FINISH.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim PACKAGE_ID = "" & RowData.Item(IdxColumnName.PACKAGE_ID.ToString)
        Dim ITEM_COMMON1 = "" & RowData.Item(IdxColumnName.ITEM_COMMON1.ToString)
        Dim ITEM_COMMON2 = "" & RowData.Item(IdxColumnName.ITEM_COMMON2.ToString)
        Dim ITEM_COMMON3 = "" & RowData.Item(IdxColumnName.ITEM_COMMON3.ToString)
        Dim ITEM_COMMON4 = "" & RowData.Item(IdxColumnName.ITEM_COMMON4.ToString)
        Dim ITEM_COMMON5 = "" & RowData.Item(IdxColumnName.ITEM_COMMON5.ToString)
        Dim ITEM_COMMON6 = "" & RowData.Item(IdxColumnName.ITEM_COMMON6.ToString)
        Dim ITEM_COMMON7 = "" & RowData.Item(IdxColumnName.ITEM_COMMON7.ToString)
        Dim ITEM_COMMON8 = "" & RowData.Item(IdxColumnName.ITEM_COMMON8.ToString)
        Dim ITEM_COMMON9 = "" & RowData.Item(IdxColumnName.ITEM_COMMON9.ToString)
        Dim ITEM_COMMON10 = "" & RowData.Item(IdxColumnName.ITEM_COMMON10.ToString)
        Dim HOST_STEP_NO = 0 & RowData.Item(IdxColumnName.HOST_STEP_NO.ToString)
        Dim HOST_MOVE_TYPE = "" & RowData.Item(IdxColumnName.HOST_MOVE_TYPE.ToString)
        Dim HOST_FINISH_TIME = "" & RowData.Item(IdxColumnName.HOST_FINISH_TIME.ToString)
        Dim HOST_BILLING_DATE = "" & RowData.Item(IdxColumnName.HOST_BILLING_DATE.ToString)
        Dim HOST_CREATE_TIME = "" & RowData.Item(IdxColumnName.HOST_CREATE_TIME.ToString)
        Dim HOST_FACTORY_NO = "" & RowData.Item(IdxColumnName.HOST_FACTORY_NO.ToString)
        Dim HOST_AREA_NO = "" & RowData.Item(IdxColumnName.HOST_AREA_NO.ToString)
        Dim HOST_OWNER_NO = "" & RowData.Item(IdxColumnName.HOST_OWNER_NO.ToString)
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
        Dim HOST_COMMON21 = "" & RowData.Item(IdxColumnName.HOST_COMMON21.ToString)
        Dim HOST_COMMON22 = "" & RowData.Item(IdxColumnName.HOST_COMMON22.ToString)
        Dim HOST_COMMON23 = "" & RowData.Item(IdxColumnName.HOST_COMMON23.ToString)
        Dim HOST_COMMON24 = "" & RowData.Item(IdxColumnName.HOST_COMMON24.ToString)
        Dim HOST_COMMON25 = "" & RowData.Item(IdxColumnName.HOST_COMMON25.ToString)
        Dim HOST_COMMON26 = "" & RowData.Item(IdxColumnName.HOST_COMMON26.ToString)
        Dim HOST_COMMON27 = "" & RowData.Item(IdxColumnName.HOST_COMMON27.ToString)
        Dim HOST_COMMON28 = "" & RowData.Item(IdxColumnName.HOST_COMMON28.ToString)
        Dim HOST_COMMON29 = "" & RowData.Item(IdxColumnName.HOST_COMMON29.ToString)
        Dim HOST_COMMON30 = "" & RowData.Item(IdxColumnName.HOST_COMMON30.ToString)
        Dim HOST_COMMON31 = "" & RowData.Item(IdxColumnName.HOST_COMMON31.ToString)
        Dim HOST_COMMON32 = "" & RowData.Item(IdxColumnName.HOST_COMMON32.ToString)
        Dim HOST_COMMON33 = "" & RowData.Item(IdxColumnName.HOST_COMMON33.ToString)
        Dim HOST_COMMON34 = "" & RowData.Item(IdxColumnName.HOST_COMMON34.ToString)
        Dim HOST_COMMON35 = "" & RowData.Item(IdxColumnName.HOST_COMMON35.ToString)
        Dim HOST_COMMON36 = "" & RowData.Item(IdxColumnName.HOST_COMMON36.ToString)
        Dim HOST_COMMON37 = "" & RowData.Item(IdxColumnName.HOST_COMMON37.ToString)
        Dim HOST_COMMON38 = "" & RowData.Item(IdxColumnName.HOST_COMMON38.ToString)
        Dim HOST_COMMON39 = "" & RowData.Item(IdxColumnName.HOST_COMMON39.ToString)
        Dim HOST_COMMON40 = "" & RowData.Item(IdxColumnName.HOST_COMMON40.ToString)
        Dim HOST_COMMENTS = "" & RowData.Item(IdxColumnName.HOST_COMMENTS.ToString)
        Info = New clsPO_DTL_Bak(PO_ID, PO_LINE_NO, PO_SERIAL_NO, SKU_NO, LOT_NO, QTY, QTY_PROCESS, QTY_FINISH, COMMENTS, PACKAGE_ID, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, HOST_FACTORY_NO, HOST_AREA_NO, HOST_STEP_NO, HOST_MOVE_TYPE, HOST_FINISH_TIME, HOST_BILLING_DATE, HOST_CREATE_TIME, HOST_OWNER_NO, HOST_COMMON1, HOST_COMMON2, HOST_COMMON3, HOST_COMMON4, HOST_COMMON5, HOST_COMMON6, HOST_COMMON7, HOST_COMMON8, HOST_COMMON9, HOST_COMMON10, HOST_COMMON11, HOST_COMMON12, HOST_COMMON13, HOST_COMMON14, HOST_COMMON15, HOST_COMMON16, HOST_COMMON17, HOST_COMMON18, HOST_COMMON19, HOST_COMMON20, HOST_COMMON21, HOST_COMMON22, HOST_COMMON23, HOST_COMMON24, HOST_COMMON25, HOST_COMMON26, HOST_COMMON27, HOST_COMMON28, HOST_COMMON29, HOST_COMMON30, HOST_COMMON31, HOST_COMMON32, HOST_COMMON33, HOST_COMMON34, HOST_COMMON35, HOST_COMMON36, HOST_COMMON37, HOST_COMMON38, HOST_COMMON39, HOST_COMMON40)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function UpdateInfo(ByRef Key As String, ByRef Info As clsPO_DTL_Bak, ByRef objNewTC As clsPO_DTL_Bak) As clsPO_DTL_Bak
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
