'20180629
'V1.0.0
'Jerry

'執行PO單 送PO to WO 給WMS

Imports eCA_HOSTObject
Imports eCA_TransactionMessage

Module Module_T5F1U11_POExecution

  Public Function O_T5F1U11_POExecution(ByVal Receive_Msg As MSG_T5F1U11_POExecution,
                                        ByRef ret_strResultMsg As String,
                                        ByRef ret_Wait_UUID As String) As Boolean
    Try
      ''儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '要變更的資料
      Dim dic_PO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If

      '進行資料處理
      If Get_Data(Receive_Msg, ret_strResultMsg, dic_PO_DTL, Host_Command, ret_Wait_UUID) = False Then
        Return False
      End If

      '取得SQL
      If _Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
        Return False
      End If

      '執行SQL與更新物件
      If _Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
        Return False
      End If


      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T5F1U11_POExecution,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        '檢查PO_ID是否為空
        If PO_ID = "" Then
          ret_strResultMsg = "PO_ID is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        ''檢查PO_Type1是否正確
        'If H_PO_ORDER_TYPE = "" Then
        '  ret_strResultMsg = "H_PO_ORDER_TYPE is empty"
        '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  Return False
        'ElseIf ModuleHelpFunc.CheckValueInEnum(Of enuOrderType)(H_PO_ORDER_TYPE) = False Then
        '  ret_strResultMsg = "H_PO_ORDER_TYPE 不存在于定义中"
        '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  Return False
        'End If

      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '資料處理
  Private Function Get_Data(ByVal Receive_Msg As MSG_T5F1U11_POExecution,
                            ByRef ret_strResultMsg As String,
                            ByRef dic_PO_DTL As Dictionary(Of String, clsPO_DTL),
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                            ByRef ret_Wait_UUID As String) As Boolean
    Try
      Dim User_ID = Receive_Msg.Header.ClientInfo.UserID

      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      '先進行資料邏輯檢查
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        'Dim IN_PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        Dim PO_ID = objPOInfo.PO_ID
        If ExcutePO.ContainsKey(PO_ID) = False Then ExcutePO.Add(PO_ID, PO_ID) '排除時間差問題 20190628
        'If H_PO_ORDER_TYPE = enuOrderType.Z_L_PO_READ_MDE_2005 Then
        '  If IN_PO_ID.Length = 15 Then
        '    For Each item In IN_PO_ID
        '      PO_ID += item
        '      If PO_ID.Length = 10 Then
        '        Exit For
        '      End If
        '    Next
        '  Else
        '    PO_ID = IN_PO_ID
        '  End If
        'Else
        '  PO_ID = IN_PO_ID
        'End If


        Dim COMMENTS As String = objPOInfo.COMMENTS
        '取得單據
        If gMain.objHandling.O_GetDB_dicPOByPOID(PO_ID, tmp_dicPO) = False Then
          ret_strResultMsg = "Select From WMS_T_PO False"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查數量
        If tmp_dicPO.Any = False Then
          ret_strResultMsg = "無法取得單據，單號:" & PO_ID & " 單據類型:" & H_PO_ORDER_TYPE
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查狀態
        If tmp_dicPO.Values(0).PO_Status = enuPOStatus.Process Then
          ret_strResultMsg = "單據正在執行中，無法再次執行。"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        Dim tmp_dicPOID As New Dictionary(Of String, String)
        tmp_dicPOID.Add(PO_ID, PO_ID)


        '使用dicPO取得資料庫裡的PO_Line資料
        If gMain.objHandling.O_Get_dicPOLineBydicPO_ID(tmp_dicPOID, tmp_dicPO_Line) = False Then
          ret_strResultMsg = "WMS get PO_Line data From DB Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '使用dicPO取得資料庫裡的PO_DTL資料
        If gMain.objHandling.O_Get_dicPODTLBydicPO_ID(tmp_dicPOID, tmp_dicPO_DTL) = False Then
          ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If


        'Dim QTY = 0
        For Each objPO_DTL In tmp_dicPO_DTL
          ' QTY += objPO_DTL.Value.QTY
          If dic_PO_DTL.ContainsKey(objPO_DTL.Key) = False Then
            dic_PO_DTL.Add(objPO_DTL.Key, objPO_DTL.Value)
          End If
        Next

        '根據單據類型(OrderType=1) 進行GVS假過帳
        'If H_PO_ORDER_TYPE = enuOrderType.Comso_Outbound_Info Then
        '  SendMessageToLog("原料(进口)出库单据执行前向GVS进行假过账", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        '  If TranTriggerSAPPostingFromTIMMStoGVS(ret_strResultMsg, PO_ID, False) = False Then
        '    Return False
        '  End If
        'End If

        'Dim LGORT = tmp_dicPO.Values(0).H_PO18 '由UI、WMS触发
        'If H_PO_ORDER_TYPE = enuOrderType.ZHLES_TIDAN_GXCK_LTK Then
        '  '执行到车登记
        '  '先进行到车登记 '10-登记时间、20-取消登记、30-扫描开始时间、40-扫描结束时间、50-快速出库时间修改
        '  If TransSMHCFromHDLTKToLES(ret_strResultMsg, PO_ID, "30", LGORT) = False Then '此次PO_ID是运单号
        '    Return False
        '  End If
        'End If


      Next


      '取得流水號
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        ret_strResultMsg = "Get UUID False"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        ret_strResultMsg = "Get UUID False"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)


      Dim SHIPPING_NO = Receive_Msg.Body.WOInfo.SHIPPING_NO

      '寫入Command給WMS
      If Send_Command_to_WMS(ret_strResultMsg, dic_PO_DTL, objUUID, Host_Command, ret_Wait_UUID, User_ID, Receive_Msg.Body.WOInfo) = False Then
        Return False
      End If




      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  'SQL
  Private Function _Get_SQL(ByRef Result_Message As String,
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                            ByRef lstSql As List(Of String)) As Boolean
    Try

      '取得要送給WMS的CMD
      For Each objHost_Command In Host_Command.Values
        If objHost_Command.O_Add_Insert_SQLString(lstSql) = False Then
          Result_Message = "Get Insert HOST_T_WMS_Command SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '執行刪除和新增的SQL語句，並進行記憶體資料更新
  Private Function _Execute_DataUpdate(ByRef Result_Message As String,
                                      ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        Result_Message = "eHOST 更新资料库失败"
        Return False
      End If

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function



  Private Function Send_Command_to_WMS(ByRef Result_Message As String, ByVal dicUpdate_PO_DTL As Dictionary(Of String, clsPO_DTL), ByRef objUUID As clsUUID,
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                                       ByRef ret_Wait_UUID As String, ByVal User_ID As String, ByVal WOInfo As MSG_T5F1U11_POExecution.clsWOInfo) As Boolean
    Try
      Dim UUID = objUUID.Get_NewUUID

      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim dicPOtoWO As New MSG_T5F1U12_POToWO
      dicPOtoWO.Header = New clsHeader
      ret_Wait_UUID = UUID
      dicPOtoWO.Header.UUID = UUID
      dicPOtoWO.Header.EventID = "T5F3U23_POToWO"
      dicPOtoWO.Header.Direction = "Primary"

      dicPOtoWO.Header.ClientInfo = New clsHeader.clsClientInfo
      dicPOtoWO.Header.ClientInfo.ClientID = "Handler"
      dicPOtoWO.Header.ClientInfo.UserID = User_ID
      dicPOtoWO.Header.ClientInfo.IP = ""
      dicPOtoWO.Header.ClientInfo.MachineID = ""

      dicPOtoWO.Body = New MSG_T5F1U12_POToWO.clsBody
      dicPOtoWO.Body.Action = "Create"
      'dicPOtoWO.Body.AutoFlag = "1"


      Dim PO_ID = ""
      Dim POList As New MSG_T5F1U12_POToWO.clsBody.clsPOList
      For Each PO_DTL In dicUpdate_PO_DTL.Values
        Dim lstPOInfo As New MSG_T5F1U12_POToWO.clsBody.clsPOList.clsPOInfo
        PO_ID = PO_DTL.PO_ID
        lstPOInfo.PO_ID = PO_DTL.PO_ID
        lstPOInfo.PO_SERIAL_NO = PO_DTL.PO_SERIAL_NO
        lstPOInfo.QTY = PO_DTL.QTY



        lstPOInfo.SKU_NO = PO_DTL.SKU_NO
        lstPOInfo.LOT_NO = PO_DTL.LOT_NO
        lstPOInfo.PACKAGE_ID = PO_DTL.PACKAGE_ID
        lstPOInfo.ITEM_COMMON1 = PO_DTL.ITEM_COMMON1
        lstPOInfo.ITEM_COMMON2 = PO_DTL.ITEM_COMMON2
        lstPOInfo.ITEM_COMMON3 = PO_DTL.ITEM_COMMON3
        lstPOInfo.ITEM_COMMON4 = PO_DTL.ITEM_COMMON4
        lstPOInfo.ITEM_COMMON5 = PO_DTL.ITEM_COMMON5
        lstPOInfo.ITEM_COMMON6 = PO_DTL.ITEM_COMMON6
        lstPOInfo.ITEM_COMMON7 = PO_DTL.ITEM_COMMON7
        lstPOInfo.ITEM_COMMON8 = PO_DTL.ITEM_COMMON8
        lstPOInfo.ITEM_COMMON9 = PO_DTL.ITEM_COMMON9
        lstPOInfo.ITEM_COMMON10 = PO_DTL.ITEM_COMMON10
        lstPOInfo.SORT_ITEM_COMMON1 = PO_DTL.SORT_ITEM_COMMON1
        lstPOInfo.SORT_ITEM_COMMON2 = PO_DTL.SORT_ITEM_COMMON2
        lstPOInfo.SORT_ITEM_COMMON3 = PO_DTL.SORT_ITEM_COMMON3
        lstPOInfo.SORT_ITEM_COMMON4 = PO_DTL.SORT_ITEM_COMMON4
        lstPOInfo.SORT_ITEM_COMMON5 = PO_DTL.SORT_ITEM_COMMON5
        lstPOInfo.FROM_OWNER_NO = PO_DTL.FROM_OWNER_ID
        lstPOInfo.FROM_SUB_OWNER_NO = PO_DTL.FROM_SUB_OWNER_ID
        lstPOInfo.TO_OWNER_NO = PO_DTL.TO_OWNER_ID
        lstPOInfo.TO_SUB_OWNER_NO = PO_DTL.TO_SUB_OWNER_ID
        lstPOInfo.FACTORY_NO = PO_DTL.FACTORY_ID
        lstPOInfo.DEST_AREA_NO = PO_DTL.DEST_AREA_ID
        lstPOInfo.DEST_LOCATION_NO = PO_DTL.DEST_LOCATION_ID
        POList.POInfo.Add(lstPOInfo)
      Next

      Dim WO_Info As New MSG_T5F1U12_POToWO.clsBody.clsWOInfo
      WO_Info.WO_ID = WOInfo.WO_ID
      'WO_Info.RECEIPT_ENABLE = WOInfo.
      WO_Info.COMMENTS = WOInfo.COMMENTS
      WO_Info.SOURCE_AREA_NO = WOInfo.SOURCE_AREA_NO
      WO_Info.SOURCE_LOCATION_NO = WOInfo.SOURCE_LOCATION_NO
      WO_Info.SHIPPING_NO = WOInfo.SHIPPING_NO
      WO_Info.SHIPPING_PRIORITY = WOInfo.SHIPPING_PRIORITY


      dicPOtoWO.Body.WOInfo = WO_Info
      dicPOtoWO.Body.POList = POList '資料填寫完成

      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_T5F1U12_POToWO(strXML, dicPOtoWO, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(T5F1U12_POToWO)"
        End If
        Return False
      End If

      '寫Command 
      'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, "T5F3U23_POToWO", 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")


      O_Send_ToWMSCommand(strXML, dicPOtoWO.Header, Host_Command)

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module
