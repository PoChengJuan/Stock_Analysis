'20190117
'V1.0.0
'Mark
'WMS回覆訂單轉工單的結果
'狀態:Open

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T5F3U23_POToWO
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~發送~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  ''發送訊息給Host
  ''組成要發送的Message的Class
  'Private Function CombinationClass(ByRef ret_strResultMsg As String,
  '                                  ByRef tmp_dicPOCloseData As Dictionary(Of String, clsPOCloseData),
  '                                  ByRef tmp_dicTextInfo As Dictionary(Of String, String),
  '                                  ByVal WO_ID As String,
  '                                  ByVal User_ID As String,
  '                                  ByVal Forced_Close As enuWOForceClose,
  '                                  ByVal To_Next As String,
  '                                  ByRef ret_objMsg_T11F1S1_POClose As MSG_T11F1S1_POClose) As Boolean
  '  Try
  '    '取得UUID
  '    Dim objUUID As clsUUID = Nothing
  '    If gMain.objWMS.O_Get_UUID(enuUUID_No.To_Host_UUID.ToString, objUUID) = False Then
  '      ret_strResultMsg = "GTE UUID Failed, UUID_No = " & enuUUID_No.To_Host_UUID.ToString
  '      Return False
  '    End If
  '    Dim To_Host_UUID As String = objUUID.Get_NewUUID
  '    Dim Receive_System As enuSystemType = enuSystemType.WMS
  '    Dim Function_ID As String = enuHostMessageFunctionID.T11F1S1_POClose.ToString
  '    Dim SEQ As Long = 1
  '    'Dim User_ID As String = enuSystemType.WMS.ToString
  '    Dim Create_Time As String = ModuleHelpFunc.GetNewTime_DBFormat
  '    Dim Result As String = ""
  '    Dim Result_Message As String = ""
  '    '組成給Host的Message
  '    '先進行PO_ID和PO_Line_No的排序
  '    Dim lstPOCloseData = (From obj In tmp_dicPOCloseData.Values Order By obj.PO_ID, obj.PO_Line_No Select obj).Distinct.ToList
  '    Dim tmp_dicPOCloseID As New Dictionary(Of String, String)
  '    '组Message给Host
  '    Dim Header As New clsHeader
  '    Header.UUID = To_Host_UUID
  '    Header.EventID = Function_ID
  '    Header.Direction = "Primary"
  '    Dim ClientInfo As New clsHeader.clsClientInfo
  '    ClientInfo.ClientID = "WMS"
  '    ClientInfo.IP = ""
  '    ClientInfo.MachineID = ""
  '    ClientInfo.UserID = User_ID
  '    Header.ClientInfo = ClientInfo
  '    ret_objMsg_T11F1S1_POClose.Header = Header
  '    Dim Body As New MSG_T11F1S1_POClose.clsBody
  '    Dim POList As New MSG_T11F1S1_POClose.clsBody.clsPOList
  '    Dim lstPOInfo As New List(Of MSG_T11F1S1_POClose.clsBody.clsPOList.clsPOInfo)
  '    Dim tmp_strPO_ID As String = ""

  '    For Each objPOCloseData As clsPOCloseData In lstPOCloseData
  '      Dim PODTLInfo As New MSG_T11F1S1_POClose.clsBody.clsPOList.clsPOInfo.clsPO_DTLList.clsPO_DTLInfo
  '      PODTLInfo.PO_LINE_NO = objPOCloseData.PO_Line_No
  '      PODTLInfo.PO_SERIAL_NO = objPOCloseData.PO_Serial_No
  '      PODTLInfo.SKU_NO = objPOCloseData.SKU_No
  '      PODTLInfo.SORT_ITEM_COMMON1 = objPOCloseData.Sort_Item_Common1
  '      PODTLInfo.SORT_ITEM_COMMON2 = objPOCloseData.Sort_Item_Common2
  '      PODTLInfo.SORT_ITEM_COMMON3 = objPOCloseData.Sort_Item_Common3
  '      PODTLInfo.SORT_ITEM_COMMON4 = objPOCloseData.Sort_Item_Common4
  '      PODTLInfo.SORT_ITEM_COMMON5 = objPOCloseData.Sort_Item_Common5
  '      PODTLInfo.QTY = objPOCloseData.Qty
  '      Dim TextList As New MSG_T11F1S1_POClose.clsBody.clsPOList.clsPOInfo.clsPO_DTLList.clsPO_DTLInfo.clsTextList
  '      Dim lstTextInfo As New List(Of MSG_T11F1S1_POClose.clsBody.clsPOList.clsPOInfo.clsPO_DTLList.clsPO_DTLInfo.clsTextList.clsTextInfo)
  '      For Each TestKey As String In tmp_dicTextInfo.Keys
  '        Dim TextInfo As New MSG_T11F1S1_POClose.clsBody.clsPOList.clsPOInfo.clsPO_DTLList.clsPO_DTLInfo.clsTextList.clsTextInfo
  '        TextInfo.Name = TestKey
  '        TextInfo.Value = tmp_dicTextInfo.Item(TestKey)
  '        lstTextInfo.Add(TextInfo)
  '      Next
  '      TextList.TextInfo = lstTextInfo
  '      PODTLInfo.TextList = TextList

  '      If tmp_strPO_ID <> objPOCloseData.PO_ID Then
  '        tmp_strPO_ID = objPOCloseData.PO_ID
  '        Dim objPOInfo As New MSG_T11F1S1_POClose.clsBody.clsPOList.clsPOInfo
  '        objPOInfo.PO_ID = objPOCloseData.PO_ID
  '        objPOInfo.H_PO_ORDER_TYPE = objPOCloseData.H_PO_Order_Type
  '        Dim PODTLList As New MSG_T11F1S1_POClose.clsBody.clsPOList.clsPOInfo.clsPO_DTLList
  '        objPOInfo.PO_DTLList = PODTLList
  '        Dim lstPODTLInfo As New List(Of MSG_T11F1S1_POClose.clsBody.clsPOList.clsPOInfo.clsPO_DTLList.clsPO_DTLInfo)
  '        PODTLList.PO_DTLInfo = lstPODTLInfo
  '        lstPOInfo.Add(objPOInfo)
  '      End If
  '      Dim objLastPOInfo = lstPOInfo.Last()
  '      objLastPOInfo.PO_DTLList.PO_DTLInfo.Add(PODTLInfo)
  '    Next
  '    POList.POInfo = lstPOInfo
  '    Body.WO_ID = WO_ID
  '    Body.Forced_Close = Forced_Close
  '    Body.To_Next = To_Next
  '    Body.POList = POList
  '    ret_objMsg_T11F1S1_POClose.Body = Body
  '    Return True
  '  Catch ex As Exception
  '    ret_strResultMsg = ex.ToString
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~接收回傳的結果~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  Public Function O_CheckMessageResult(ByVal Receive_Msg As MSG_T5F3U23_POToWO,
                                       ByRef ret_strResultMsg As String,
                                       ByVal strRejectReason As String,
                                       Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      Dim dic_AddProductionInfo As New Dictionary(Of String, clsProduce_Info)

      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
      Dim dicGluePO_DTL As New Dictionary(Of String, clsPO_DTL)

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料處理
      If Get_Data(Receive_Msg, ret_strResultMsg, blnResult, dic_AddProductionInfo, dicGluePO_DTL, strRejectReason) = False Then
        Return False
      End If
      '取得要更新到DB的SQL
      If Get_SQL(ret_strResultMsg, dic_AddProductionInfo, lstSql) = False Then
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(ret_strResultMsg, dic_AddProductionInfo, lstSql) = False Then
        Return False
      End If

#Region "將膠塊的部份另外自動執行"

      '將要送給WMS的COMMAND先清除
      Host_Command.Clear()
      lstSql.Clear()
      '膠塊需要另外處理
      If dicGluePO_DTL.Any AndAlso Send_T5F2U62_AutoInbound_to_WMS(ret_strResultMsg, Host_Command, dicGluePO_DTL) = False Then
        Return False
      End If
      '取得SQL
      If Get_SQL_Host_Command(ret_strResultMsg, Host_Command, lstSql) = False Then

        Return False
      End If
      '執行SQL與更新物件
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then

        Return False
      End If
#End Region

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T5F3U23_POToWO,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      Dim UUID As String = Receive_Msg.Header.UUID
      If UUID = "" Then
        ret_strResultMsg = "UUID is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '資料處理
  Private Function Get_Data(ByVal Receive_Msg As MSG_T5F3U23_POToWO,
                            ByRef ret_strResultMsg As String,
                            ByRef blnResult As Boolean,
                            ByRef ret_dic_AddProductionInfo As Dictionary(Of String, clsProduce_Info),
                            ByRef ret_dicGluePO_DTL As Dictionary(Of String, clsPO_DTL),
                            ByVal strRejectReason As String) As Boolean
    Try
      Dim Hist_UUID As String = GetNewTime_ByDataTimeFormat(DBFullTimeUUIDFormat)
      Dim UUID As String = Receive_Msg.Header.UUID
      Dim Now_Time As String = GetNewTime_DBFormat()
      '取出所有PO_ID
      Dim tmp_dicPO_ID As New Dictionary(Of String, String)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      For Each POInfo In Receive_Msg.Body.POList.POInfo
        Dim PO_ID As String = POInfo.PO_ID
        If tmp_dicPO_ID.ContainsKey(PO_ID) = False Then
          tmp_dicPO_ID.Add(PO_ID, PO_ID)
        End If
      Next
      '抓取資料庫PO的資料
      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(tmp_dicPO_ID, tmp_dicPO) = True Then
        For Each objPO As clsPO In tmp_dicPO.Values
          Dim PO_TYPE1 = objPO.PO_Type1
          Dim PO_TYPE2 = objPO.PO_Type2
          'If objPO.H_PO_ORDER_TYPE = enuOrderType.produce_in_after Or objPO.H_PO_ORDER_TYPE = enuOrderType.produce_in_before Then
          '如果已經生產列表中了，就不新增
          'If gMain.objHandling.O_Get_dicProductionInfoByPOID(objPO.PO_ID) = True Then
          '  Continue For
          'End If
          'Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
          'If gMain.objHandling.O_GetDB_dicPODTLByPOID_POSerialNo(objPO.PO_ID, "", tmp_dicPO_DTL) = False Then
          '  ret_strResultMsg = "Not Find PO DTL PO_ID=" & objPO.PO_ID
          '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '  Continue For  '找不到單據明細
          'End If
          ''把PO_ID加入Production
          'For Each objLineArea As clsLine_Area In gMain.objHandling.gdicLine_Area.Values
          '  '分前後製程 加入
          '  If objPO.H_PO_ORDER_TYPE = enuOrderType.produce_in_after Then
          '    If objLineArea.AREA_TYPE1 = enuAreaType1.BackProcess Then
          '      With objLineArea
          '        For Each objPO_DTL As clsPO_DTL In tmp_dicPO_DTL.Values
          '          Dim objNewProductionInfo As New clsProduce_Info(.Factory_No, .Area_No, objPO.PO_ID, objPO_DTL.SKU_No, enuProduceStatus.Queued, objPO_DTL.QTY, 0, 0, 0, 0, objLineArea.PREVIOUS_AREA_NO, Now_Time, "", "", "", objPO.H_PO9, objPO.H_PO10, objPO.H_PO8, objPO_DTL.H_POD1, "", "", "", "", "", "")
          '          If ret_dic_AddProductionInfo.ContainsKey(objNewProductionInfo.gid) = False Then
          '            ret_dic_AddProductionInfo.Add(objNewProductionInfo.gid, objNewProductionInfo)
          '          End If
          '        Next
          '      End With
          '    End If
          '  ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.produce_in_before Then
          '    If objLineArea.AREA_TYPE1 = enuAreaType1.FrontProcess Then
          '      With objLineArea
          '        For Each objPO_DTL As clsPO_DTL In tmp_dicPO_DTL.Values
          '          Dim objNewProductionInfo As New clsProduce_Info(.Factory_No, .Area_No, objPO.PO_ID, objPO_DTL.SKU_No, enuProduceStatus.Queued, objPO_DTL.QTY, 0, 0, 0, 0, objLineArea.PREVIOUS_AREA_NO, Now_Time, "", "", "", objPO.H_PO9, objPO.H_PO10, objPO.H_PO8, objPO_DTL.H_POD1, "", "", "", "", "", "")
          '          If ret_dic_AddProductionInfo.ContainsKey(objNewProductionInfo.gid) = False Then
          '            ret_dic_AddProductionInfo.Add(objNewProductionInfo.gid, objNewProductionInfo)
          '          End If
          '        Next
          '      End With
          '    End If
          '  End If
          'Next

          'End If

          '上報單據已放行 
          If objPO.H_PO_ORDER_TYPE = enuOrderType.Inbound_Data Then
            '採購入庫單
            SendInboundData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
            'ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.ProduceInData Then
            '  '生產入庫單
            '  SendProduceInData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
            'ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.OtherInData Then
            '  '雜收單
            '  SendOtherInData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
            'ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.SellReturn Then
            '  '銷退單
            '  SendSellReturnData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
            'ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.transaction_in Then
            '  '調撥入
            '  SendTransactionData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
            'ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.transaction_out Then
            '  '調撥出
            '  SendTransactionData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
            'ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.OtherOutData Then
            '  '雜發單
            '  SendOtherOutData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
            'ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.SellData Then
            '  '銷貨單
            '  SendSellData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
            'ElseIf objPO.H_PO_ORDER_TYPE = enuOrderType.InboundReturn_Data Then
            '  '退供應商單
            '  SendInbound_ReturnData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
          End If

          '使用dicPO取得資料庫裡的PO_DTL資料
          If gMain.objHandling.O_Get_dicPODTLBydicPO_ID(tmp_dicPO_ID, tmp_dicPO_DTL) = False Then
            ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          For Each objPO_DTL In tmp_dicPO_DTL.Values
            '檢查料品主檔是否存在
            Dim dicSKU As New Dictionary(Of String, clsSKU)
            '判斷是否是膠塊
            Dim blnGlue As Boolean = False
            If gMain.objHandling.O_GetDB_dicSKUBySKUNo(objPO_DTL.SKU_NO, dicSKU) = True Then
              If dicSKU.Any Then
                Dim objSKU = dicSKU.First.Value

                '將需要另外處理的紀錄起來
                If objSKU.SKU_TYPE2 = "1" Then  '膠塊
                  blnGlue = True
                End If
              Else
                ret_strResultMsg = "料品不存在無法取得庫別資訊, PO_ID=" & objPO_DTL.PO_ID & ", SKU_NO=" & objPO_DTL.SKU_NO & " 請先建立品號資料"
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If
            End If
            '紀錄是膠塊的項次
            'If blnGlue = True AndAlso PO_TYPE1 = enuPOType_1.Combination_in AndAlso PO_TYPE2 = enuPOType_2.transaction_in Then
            If blnGlue = True AndAlso PO_TYPE1 = enuPOType_1.Combination_in Then

              If ret_dicGluePO_DTL.ContainsKey(objPO_DTL.gid) = False Then
                ret_dicGluePO_DTL.Add(objPO_DTL.gid, objPO_DTL)
              End If
            End If
          Next
        Next
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef ret_dic_AddProductionInfo As Dictionary(Of String, clsProduce_Info),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得 SQL
      For Each obj As clsProduce_Info In ret_dic_AddProductionInfo.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert Production Info SQL Failed"
          Return False
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL_Host_Command(ByRef ret_strResultMsg As String,
                           ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得Host_Command的SQL
      For Each _Host_COMMAND In Host_Command.Values
        If _Host_COMMAND.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert HOST_T_WMS_Command SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                     ByRef ret_dic_AddProductionInfo As Dictionary(Of String, clsProduce_Info),
                                      ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSql.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
        '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
        'ret_strResultMsg = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
        'SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False
        Return True
      End If
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        ret_strResultMsg = "WMS Update DB Failed"
        Return False
      End If
      '修改記憶體資料
      For Each objNew As clsProduce_Info In ret_dic_AddProductionInfo.Values
        objNew.Add_Relationship(gMain.objHandling)
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行刪除和新增的SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                      ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        ret_strResultMsg = "WMS Update DB Failed"
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '領退料單放行回報
  Public Sub SendReturnData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "TWRDC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TWRDT = WorkType '單別	
      Rocord_Head_Info.TWRNO = WorkID '單號
      Rocord_Head_Info.TWRRS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP領退料單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      'If RtnCode = "-1" Then

      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      'End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Public Sub SendPickUpData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "TWIDC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TWIDT = WorkType '單別	
      Rocord_Head_Info.TWINO = WorkID '單號
      Rocord_Head_Info.TWIRS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP領料單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      'If RtnCode = "-1" Then

      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      'End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '退料單
  Public Sub SendPickingData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "MOCTC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TC001 = WorkType '單別	
      Rocord_Head_Info.TC002 = WorkID '單號
      Rocord_Head_Info.TC200 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP領退料單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      'If RtnCode = "-1" Then

      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      'End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Public Sub SendOtherOutData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "TXIDC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TXIDT = WorkType '單別	
      Rocord_Head_Info.TXINO = WorkID '單號
      Rocord_Head_Info.TXIRS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP雜發單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      'If RtnCode = "-1" Then

      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      'End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Public Sub SendOtherInData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "TXSDC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TXSDT = WorkType '單別	
      Rocord_Head_Info.TXSNO = WorkID '單號
      Rocord_Head_Info.TXSRS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP雜收單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      'If RtnCode = "-1" Then

      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      'End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Public Sub SendSellData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "TDNDC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TDNDT = WorkType '單別	
      Rocord_Head_Info.TDNNO = WorkID '單號
      Rocord_Head_Info.TDNRS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP銷售單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      'If RtnCode = "-1" Then

      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      'End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Public Sub SendSellReturnData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "SALRT"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.ST001 = WorkType '單別	
      Rocord_Head_Info.ST002 = WorkID '單號
      Rocord_Head_Info.ST003 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP銷退單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      'If RtnCode = "-1" Then

      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      'End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '生產入庫單放行回報
  Public Sub SendProduceInData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "TWSDC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TWSDT = WorkType '單別	
      Rocord_Head_Info.TWSNO = WorkID '單號
      Rocord_Head_Info.TWSRS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP領退料單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If RtnCode = "-1" Then

      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
        If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
          SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          'Return False '將obj轉成xml
        End If
        'End If
        STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '託外進貨單放行回報
  Public Sub SendOutsourcePurchaseData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "MOCTH"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TH001 = WorkType '單別	
      Rocord_Head_Info.TH002 = WorkID '單號
      Rocord_Head_Info.TH200 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP領退料單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      ''If RtnCode = "-1" Then

      ''Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      '' End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '庫存異動單放行回報
  Public Sub SendTransactionData_Normal(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "INVTA"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TA001 = WorkType '單別	
      Rocord_Head_Info.TA002 = WorkID '單號
      Rocord_Head_Info.TA200 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP領退料單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If RtnCode = "-1" Then
      STD_IN(STDIN)
      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
        If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
          SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          'Return False '將obj轉成xml
        End If
      'End If
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '轉播單放行回報
  Public Sub SendTransactionData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "TWTDC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TWTDT = WorkType '單別	
      Rocord_Head_Info.TWTNO = WorkID '單號
      Rocord_Head_Info.TWTRS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP調撥單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If RtnCode = "-1" Then
      STD_IN(STDIN)
      ' Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
        If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
          SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          'Return False '將obj轉成xml
        End If
      'End If
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '暫出/入單放行回報
  Public Sub SendTempInOutData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "INVTH"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TH001 = WorkType '單別	
      Rocord_Head_Info.TH002 = WorkID '單號
      Rocord_Head_Info.TH200 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP領退料單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If RtnCode = "-1" Then
      STD_IN(STDIN)
      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
        If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
          SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          'Return False '將obj轉成xml
        End If
      'End If
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '採購單
  Public Sub SendPurchaserData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "TPODC"  'ALAN_0708

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TPODT = WorkType '單別	
      Rocord_Head_Info.TPONO = WorkID '單號
      Rocord_Head_Info.TPORS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP進貨單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If RtnCode = "-1" Then
      STD_IN(STDIN)
      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      ' End If
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '採購入庫單
  Public Sub SendInboundData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "PURTG"  'ALAN_0708

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TG001 = WorkType '單別	
      Rocord_Head_Info.TG002 = WorkID '單號
      Rocord_Head_Info.TG200 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP採購入庫單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If RtnCode = "-1" Then
      STD_IN(STDIN)
      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      ' End If
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '採購入庫單
  Public Sub SendInbound_ReturnData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "PURRT"  'ALAN_0708

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TG001 = WorkType '單別	
      Rocord_Head_Info.TG002 = WorkID '單號
      Rocord_Head_Info.TG200 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP採購入庫單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If RtnCode = "-1" Then
      STD_IN(STDIN)
      'Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      ' End If
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '新品號
  Public Sub SendSKUData(ByVal RtnCode As String, ByVal SKU As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "INVMO"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.MO001 = SKU '單別	
      'Rocord_Head_Info.TG002 = WorkID '單號
      Rocord_Head_Info.MO200 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP品號：" & SKU & "接收成功", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '新品號
  Public Sub SendSKUChangeData(ByVal RtnCode As String, ByVal SKU As String, ByVal Edition As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "INVTL"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.TL001 = SKU '單別	
      Rocord_Head_Info.TL004 = Edition '單號
      Rocord_Head_Info.TL200 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP品號變更：" & SKU & "接收成功", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '新品號
  Public Sub SendWarehouseData(ByVal RtnCode As String, ByVal Warehouse As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "NWHDC"

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.NWHDT = Warehouse '單別	
      'Rocord_Head_Info.TG002 = WorkID '單號
      Rocord_Head_Info.NWHRS = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP倉別：" & Warehouse & "接收成功", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  '貨主轉播單放行回報
  Public Sub SendTransactionOwnerData(ByVal RtnCode As String, ByVal WorkType As String, ByVal WorkID As String, Optional ByRef strXML As String = "")
    Try
      Dim STDIN As New MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = Companyid ' "Leader"
      STDIN.Userid = "DS"
      STDIN.DoAction = "CMSAB01"
      STDIN.Docase = "1"

      'Data
      Dim Data As New STD_INData
      Dim FormHead As New STD_INDataFormHead
      FormHead.TableName = "INTRS" '等對方開

      '組成Header
      Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '<RecordList>
      '  <TG001>進貨單單別</TG001>    
      '  <TG002>進貨單單號</TG002>
      '  <TG200>WMS接收成功</TG200>                 
      '</RecordList>   
      Rocord_Head_Info.INT01 = WorkType '單別	
      Rocord_Head_Info.INT02 = WorkID '單號
      Rocord_Head_Info.INT03 = RtnCode ' -1:已放行 0:未放行 1:錯誤
      Rocord_Head(0) = Rocord_Head_Info
      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead
      STDIN.Result = "success"
      STDIN.Data = Data
      SendMessageToLog("通知ERP貨主調撥單放行，單號：" & WorkID & " ，單別：" & WorkType, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      'If RtnCode = "-1" Then
      STD_IN(STDIN)
      ' Else
      'Dim strXML = ""
      Dim ReturnMessage = ""
      If PrepareMessage_SendTransferDataToERP(strXML, STDIN, ReturnMessage) = False Then
        SendMessageToLog("Obj轉換Xml失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False '將obj轉成xml
      End If
      'End If
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub


End Module
