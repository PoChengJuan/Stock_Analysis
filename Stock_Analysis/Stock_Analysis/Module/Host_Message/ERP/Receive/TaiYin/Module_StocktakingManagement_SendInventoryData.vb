'20200106
'V1.0.0
'Vito
'Vito_20106
'WMS向HostHandler進行收料資訊的回報

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_StocktakingManagement_SendInventoryData
  Public Function O_StocktakingManagement_SendInventoryData(ByVal Receive_Msg As MSG_SendInventoryData,
                                          ByRef ret_strResultMsg As String) As Boolean
    Try

      Dim lstSql As New List(Of String)
      Dim dicAddStockTaking As New Dictionary(Of String, clsTSTOCKTAKING)
      Dim dicAddStockTaking_dtl As New Dictionary(Of String, clsTSTOCKTAKINGDTL)
      Dim dicDeleteStockTaking As New Dictionary(Of String, clsTSTOCKTAKING)
      Dim dicDeleteStockTaking_dtl As New Dictionary(Of String, clsTSTOCKTAKINGDTL)
      Dim dicUpdateStockTaking As New Dictionary(Of String, clsTSTOCKTAKING)
      Dim dicUpdateStockTaking_dtl As New Dictionary(Of String, clsTSTOCKTAKINGDTL)
      Dim Host_command As New Dictionary(Of String, clsFromHostCommand)
      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料處理
      If Process_Data(Receive_Msg, ret_strResultMsg, dicAddStockTaking, dicAddStockTaking_dtl, dicDeleteStockTaking, dicDeleteStockTaking_dtl, dicUpdateStockTaking, dicUpdateStockTaking_dtl, Host_command) = False Then
        Return False
      End If

      If Get_SQL(ret_strResultMsg, lstSql, Host_command) = False Then
        Return False
      End If
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_SendInventoryData,
                              ByRef ret_strResultMsg As String) As Boolean
    Try

      '先進行資料邏輯檢查
      For Each objInventoryDataInfo In Receive_Msg.InventoryDataList.InventoryDataInfo
        Dim POId As String = objInventoryDataInfo.POId
        If POId = "" Then
          ret_strResultMsg = "POId is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        For Each objInventoryDeatilInfo In objInventoryDataInfo.InventoryDetailDataList.InventoryDetailDataInfo
          Dim SerialId As String = objInventoryDeatilInfo.SerialId
          Dim SKU As String = objInventoryDeatilInfo.SKU
          Dim LotInventory As String = objInventoryDeatilInfo.LotInventory
          Dim LotId As String = objInventoryDeatilInfo.LotId
          Dim InventoryQty As String = objInventoryDeatilInfo.InventoryQty
          If SerialId = "" Then
            ret_strResultMsg = "SreialId is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If SKU = "" Then
            ret_strResultMsg = "SKU is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If InventoryQty = "" Then
            ret_strResultMsg = "InventoryQty is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        Next
      Next

      'If IsShowQty = "" Then
      '  ret_strResultMsg = "IsShowQty is empty"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If


      'Dim StocktakingInfo = Receive_Msg.Body.StocktakingInfo
      'Dim STOCKTAKING_ID As String = StocktakingInfo.STOCKTAKING_ID
      'Dim LOCATION_GROUP_NO As String = StocktakingInfo.LOCATION_GROUP_NO
      'Dim PRIORITY As String = StocktakingInfo.PRIORITY
      'Dim STOCKTAKING_TYPE1 As String = StocktakingInfo.STOCKTAKING_TYPE1
      'Dim STOCKTAKING_TYPE2 As String = StocktakingInfo.STOCKTAKING_TYPE2
      'Dim STOCKTAKING_TYPE3 As String = StocktakingInfo.STOCKTAKING_TYPE3
      'Dim SEND_TO_HOST As String = StocktakingInfo.SEND_TO_HOST
      'Dim CHANGGE_INVENTORY As String = StocktakingInfo.CHANGE_INVENTORY

      'If STOCKTAKING_ID = "" Then
      '  ret_strResultMsg = "STOCKTAKING_ID is empty"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      'If STOCKTAKING_TYPE1 = "" Then
      '  ret_strResultMsg = "STOCKTAKING_TYPE1 is empty"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      'If STOCKTAKING_TYPE2 = "" Then
      '  ret_strResultMsg = "STOCKTAKING_TYPE2 is empty"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      'If STOCKTAKING_TYPE3 = "" Then
      '  ret_strResultMsg = "STOCKTAKING_TYPE3 is empty"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If

      'For Each objStockTakingDtlInfo In StocktakingInfo.StocktakingDTLList.StocktakingDTLInfo
      '  Dim STOCKTAKING_SERIAL_NO As String = objStockTakingDtlInfo.STOCKTAKING_SERIAL_NO
      '  Dim AREA_NO As String = objStockTakingDtlInfo.AREA_NO
      '  Dim BLICK_NO As String = objStockTakingDtlInfo.BLOCK_NO
      '  Dim SKU_NO As String = objStockTakingDtlInfo.SKU_NO
      '  Dim BND As String = objStockTakingDtlInfo.BND
      '  Dim SL_NO As String = objStockTakingDtlInfo.SL_NO
      '  Dim CARRIER_ID As String = objStockTakingDtlInfo.CARRIER_ID
      '  Dim PERCENETAGE As String = objStockTakingDtlInfo.PERCENTAGE
      '  Dim LOT_NO As String = objStockTakingDtlInfo.LOT_NO
      '  DIM ITEM_COMMON1 AS STRING =OBJSTOCKTAKINGDTLINFO.ITEM_COMMON1
      '  Dim ITEM_COMMON2 As String = objStockTakingDtlInfo.ITEM_COMMON2
      '  Dim ITEM_COMMON3 As String = objStockTakingDtlInfo.ITEM_COMMON3
      '  Dim ITEM_COMMON4 As String = objStockTakingDtlInfo.ITEM_COMMON4
      '  Dim ITEM_COMMON5 As String = objStockTakingDtlInfo.ITEM_COMMON5
      '  Dim ITEM_COMMON6 As String = objStockTakingDtlInfo.ITEM_COMMON6
      '  Dim ITEM_COMMON7 As String = objStockTakingDtlInfo.ITEM_COMMON7
      '  Dim ITEM_COMMON8 As String = objStockTakingDtlInfo.ITEM_COMMON8
      '  Dim ITEM_COMMON9 As String = objStockTakingDtlInfo.ITEM_COMMON9
      '  Dim ITEM_COMMON10 As String = objStockTakingDtlInfo.ITEM_COMMON10
      '  Dim SORT_ITEM_COMMON1 As String = objStockTakingDtlInfo.SORT_ITEM_COMMON1
      '  Dim SORT_ITEM_COMMON2 As String = objStockTakingDtlInfo.SORT_ITEM_COMMON2
      '  Dim SORT_ITEM_COMMON3 As String = objStockTakingDtlInfo.SORT_ITEM_COMMON3
      '  Dim SORT_ITEM_COMMON4 As String = objStockTakingDtlInfo.SORT_ITEM_COMMON4
      '  Dim SORT_ITEM_COMMON5 As String = objStockTakingDtlInfo.SORT_ITEM_COMMON5
      '  Dim OWNER_NO As String = objStockTakingDtlInfo.OWNER_NO
      '  Dim SUB_OWNER_NO As String = objStockTakingDtlInfo.SUB_OWNER_NO
      '  Dim SUPPLIER_NO As String = objStockTakingDtlInfo.SUPPLIER_NO
      '  Dim CUSTOMER_NO As String = objStockTakingDtlInfo.CUSTOMER_NO
      '  Dim RECEIPT_DATE As String = objStockTakingDtlInfo.RECEIPT_DATE
      '  Dim MANUFAETURE_DATE As String = objStockTakingDtlInfo.MANUFACETURE_DATE
      '  Dim EXPIRED_DATE As String = objStockTakingDtlInfo.EXPIRED_DATE
      '  Dim ERP_QTY As String = objStockTakingDtlInfo.ERP_QTY

      'Next

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '資料處理
  Private Function Process_Data(ByVal Receive_Msg As MSG_SendInventoryData,
                                ByRef ret_strResultMsg As String, ByRef ret_dicAddStockTaking As Dictionary(Of String, clsTSTOCKTAKING), ByRef ret_dicAddStockTaking_dtl As Dictionary(Of String, clsTSTOCKTAKINGDTL),
                                ByRef ret_dicDeleteStockTaking As Dictionary(Of String, clsTSTOCKTAKING), ret_dicDeleteStockTaking_dtl As Dictionary(Of String, clsTSTOCKTAKINGDTL),
                                ByRef ret_dicUpdateStockTaking As Dictionary(Of String, clsTSTOCKTAKING), ByRef ret_dicUpdateSotckTaking_dtl As Dictionary(Of String, clsTSTOCKTAKINGDTL),
                                ByRef host_command As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try
      Dim NOW_TIME = GetNewTime_DBFormat()

      For Each objInventoryInfo In Receive_Msg.InventoryDataList.InventoryDataInfo
        Dim Warehouse = objInventoryInfo.Warehouse


        Dim STOCKTAKING_ID = objInventoryInfo.POId
        Dim LOCATION_GROUP_NO As String = ""
        Dim PRIORITY As Long = 50
        Dim STOCKTAKING_TYPE1 As Long = 2
        Dim STOCKTAKING_TYPE2 As Long = 1
        Dim STOCKTAKING_TYPE3 As Long = 0  '0=明盤 1=暗盤
        Dim CREATE_TIME As String = NOW_TIME
        Dim START_TIME As String = ""
        Dim FINISH_TIME As String = ""
        Dim CREATE_USER As String = ""
        Dim STATUS As Long = 0
        Dim CARRIER_QTY As Long = 0
        Dim CARRIER_QTY_CHECKED As Long = 0
        Dim MATCH_TYPE As Long = 0
        Dim UPLOAD_STATUS As Long = 0
        Dim UPLOAD_COMMENTS As String = ""
        Dim SEND_TO_HOST As Long = 1 '上傳=1 不上傳=0
        Dim CHANGE_INVENTORY As Long = 0    '不修改=0 修改=1

        Dim objNewStockTaking = New clsTSTOCKTAKING(STOCKTAKING_ID, STOCKTAKING_TYPE1, STOCKTAKING_TYPE2, STOCKTAKING_TYPE3, CREATE_TIME, START_TIME, FINISH_TIME, CREATE_USER, STATUS, LOCATION_GROUP_NO, PRIORITY, CARRIER_QTY, CARRIER_QTY_CHECKED,
                                    MATCH_TYPE, SEND_TO_HOST, CHANGE_INVENTORY, UPLOAD_STATUS, UPLOAD_COMMENTS)
        Dim dicStocktaking As New Dictionary(Of String, clsTSTOCKTAKING)
        If gMain.objHandling.O_GetDB_dicStocktakingByStocktakingID(STOCKTAKING_ID, dicStocktaking) = False Then
          ret_strResultMsg = String.Format("Get WMS_M_STOCKTAKING Fail")
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If dicStocktaking.Any = True Then
          For Each objStocktkaing In dicStocktaking.Values
            If objStocktkaing.STATUS <> enuSTOCKTAKING_STATUS.Queued Then
              ret_strResultMsg = String.Format("盤點單已執行，不允許進行修改，盤點單號={0}", STOCKTAKING_ID)
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False  '要直接拒絕還是跳過
            Else
              If ret_dicUpdateStockTaking.ContainsKey(objNewStockTaking.gid) = False Then
                ret_dicUpdateStockTaking.Add(objNewStockTaking.gid, objNewStockTaking)
              End If
            End If
          Next
        Else
          If ret_dicAddStockTaking.ContainsKey(objNewStockTaking.gid) = False Then
            ret_dicAddStockTaking.Add(objNewStockTaking.gid, objNewStockTaking)
          End If
        End If

        '組成表身
        For Each objdata In objInventoryInfo.InventoryDetailDataList.InventoryDetailDataInfo

          Dim SeaildId As String = objdata.SerialId
          Dim StockTaking_SERIAL_No = SeaildId

          Dim SKU As String = objdata.SKU & "_" & "TYM" & "_" & "ERP"
          Dim LotInventory As String = objdata.LotInventory
          Dim LotId As String = objdata.LotId
          Dim InventoryQty As String = objdata.InventoryQty



          Dim AREA_NO As String = ""
          Dim BLOCK_NO As String = ""
          Dim SKU_NO As String = SKU
          Dim CARRIER_ID As String = ""
          Dim PERCENTAGE As Long = 100
          Dim LOT_NO As String = LotId
          Dim SL_NO As String = ""
          Dim STORAGE_TYPE As String = ""
          Dim BND As String = ""


          Dim ITEM_COMMON1 As String = ""
          Dim ITEM_COMMON2 As String = ""
          Dim ITEM_COMMON3 As String = Warehouse
          Dim ITEM_COMMON4 As String = ""
          Dim ITEM_COMMON5 As String = ""
          Dim ITEM_COMMON6 As String = ""
          Dim ITEM_COMMON7 As String = ""
          Dim ITEM_COMMON8 As String = ""
          Dim ITEM_COMMON9 As String = ""
          Dim ITEM_COMMON10 As String = ""
          Dim SORT_ITEM_COMMON1 As String = ""
          Dim SORT_ITEM_COMMON2 As String = ""
          Dim SORT_ITEM_COMMON3 As String = ""
          Dim SORT_ITEM_COMMON4 As String = ""
          Dim SORT_ITEM_COMMON5 As String = ""
          Dim OWNER_NO As String = ""
          Dim SUB_OWNER_NO As String = ""
          Dim SUPPLIER_NO As String = ""
          Dim CUSTOMER_NO As String = ""
          Dim RECEIPT_DATE As String = ""
          Dim MANUFACETURE_DATE As String = ""
          Dim EXPIRED_DATE As String = ""
          Dim ERP_QTY As Double = InventoryQty

          '如果沒結束料號則只輸出開始料號盤點明細

          Dim objNewStockTaking_dtl = New clsTSTOCKTAKINGDTL(STOCKTAKING_ID, StockTaking_SERIAL_NO, AREA_NO, BLOCK_NO, SKU_NO, OWNER_NO, SUB_OWNER_NO, SL_NO, STORAGE_TYPE, BND, CARRIER_ID, PERCENTAGE, CARRIER_QTY, CARRIER_QTY_CHECKED, LOT_NO,
                                        ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5,
                                        SUPPLIER_NO, CUSTOMER_NO, RECEIPT_DATE, MANUFACETURE_DATE, EXPIRED_DATE, ERP_QTY)
          If ret_dicAddStockTaking.ContainsKey(objNewStockTaking.gid) = False AndAlso ret_dicUpdateStockTaking.ContainsKey(objNewStockTaking.gid) = True Then
            If ret_dicUpdateSotckTaking_dtl.ContainsKey(objNewStockTaking_dtl.gid) = False Then
              ret_dicUpdateSotckTaking_dtl.Add(objNewStockTaking_dtl.gid, objNewStockTaking_dtl)
            End If
          Else
            If ret_dicAddStockTaking_dtl.ContainsKey(objNewStockTaking_dtl.gid) = False Then
              ret_dicAddStockTaking_dtl.Add(objNewStockTaking_dtl.gid, objNewStockTaking_dtl)
            End If
          End If

        Next
      Next
      If ret_dicAddStockTaking.Any Then
        If Module_Send_WMSMessage.Send_T10F2U1_StocktakingManagement_to_WMS(ret_strResultMsg, ret_dicAddStockTaking, ret_dicAddStockTaking_dtl, host_command, "Create") = False Then
          Return False
        End If
      End If
      If ret_dicUpdateStockTaking.Any Then
        If Module_Send_WMSMessage.Send_T10F2U1_StocktakingManagement_to_WMS(ret_strResultMsg, ret_dicUpdateStockTaking, ret_dicUpdateSotckTaking_dtl, host_command, "Modify") = False Then
          Return False
        End If
      End If
      '先進行資料邏輯檢查
      'Dim Now_Time As String = GetNewTime_DBFormat()
      'Dim Now_Date As String = GetNewDate_DBFormat_YYYYMMDD()
      'Dim USER_ID = Receive_Msg.Header.ClientInfo.UserID
      'Dim UUID = Receive_Msg.Header.UUID
      'Dim ReceiptUUID As String = ""
      'gMain.objHandling.O_Get_UUID(ReceiptUUID)
      'DeliveryReport.BatchID = CombinationBatchID(Now_Date, ReceiptUUID)
      'DeliveryReport.TotalRowCount = Receive_Msg.Body.DeliveryInfo.DeliveryDTLList.DeliveryDTLInfo.Count
      ''Dim msg_IWMS_2001 As New MSG_IWMS2001
      'Dim DeliveryInfo = Receive_Msg.Body.DeliveryInfo

      'Dim WO_ID As String = DeliveryInfo.WO_ID
      'Dim OWNER_NO As String = DeliveryInfo.OWNER_NO
      'Dim SUB_OWNER_NO As String = DeliveryInfo.SUB_OWNER_NO
      'Dim SHIPPING_COMPANY_NO As String = DeliveryInfo.SHIPPING_COMPANY_NO
      'Dim SHIPPING_UID As String = DeliveryInfo.SHIPPING_UID
      'Dim COMMON1 As String = DeliveryInfo.COMMON1
      'Dim COMMON2 As String = DeliveryInfo.COMMON2
      'Dim COMMON3 As String = DeliveryInfo.COMMON3
      'Dim COMMON4 As String = DeliveryInfo.COMMON4
      'Dim COMMON5 As String = DeliveryInfo.COMMON5
      'Dim COMMON6 As String = DeliveryInfo.COMMON6
      'Dim COMMON7 As String = DeliveryInfo.COMMON7
      'Dim COMMON8 As String = DeliveryInfo.COMMON8
      'Dim COMMON9 As String = DeliveryInfo.COMMON9
      'Dim COMMON10 As String = DeliveryInfo.COMMON10
      'Dim msg_DeliveryReport As New MSG_DeliveryReport_Info
      'DeliveryReport.data = New List(Of MSG_DeliveryReport_Info)
      'msg_DeliveryReport.LineNumber = COMMON6   '暫定
      'msg_DeliveryReport.LotString01 = COMMON1
      'msg_DeliveryReport.LotString01Descr = COMMON2
      'msg_DeliveryReport.LotNo = COMMON7
      'msg_DeliveryReport.LotID = COMMON3
      'msg_DeliveryReport.StockOutOP = COMMON4
      'For Each objDeliveryDTLInfo In Receive_Msg.Body.DeliveryInfo.DeliveryDTLList.DeliveryDTLInfo
      '  '資料檢查
      '  Dim CARRIER_ID As String = objDeliveryDTLInfo.CARRIER_ID
      '  Dim CARRIER_LABEL_ID As String = objDeliveryDTLInfo.CARRIER_LABEL_ID
      '  Dim CARRIER_KIND As String = objDeliveryDTLInfo.CARRIER_KIND
      '  Dim PO_ID As String = objDeliveryDTLInfo.PO_ID
      '  Dim DELIVERY_TIME As String = objDeliveryDTLInfo.DELIVERY_TIME
      '  msg_DeliveryReport.PalletID = CARRIER_ID
      '  msg_DeliveryReport.StockOutDate = DELIVERY_TIME
      '  For Each objItemInfo In objDeliveryDTLInfo.ItemList.ItemInfo
      '    Dim SKU_KIND As String = objItemInfo.SKU_KIND
      '    Dim SKU_NO As String = objItemInfo.SKU_NO
      '    Dim SKU_ID1 As String = objItemInfo.SKU_ID1
      '    Dim QTY As String = objItemInfo.QTY
      '    msg_DeliveryReport.SkuID = SKU_NO
      '    msg_DeliveryReport.Qty = QTY
      '    DeliveryReport.data.Add(msg_DeliveryReport)
      '  Next

      'Next

      'msg_IWMS_2001.MANDT = ""
      'msg_IWMS_2001.ZLGNUM = ""
      'msg_IWMS_2001.TIMESTAMPL = ""
      'msg_IWMS_2001.ZORDER_TYPE = ""
      'msg_IWMS_2001.ZORDER_NO = ""
      'msg_IWMS_2001.BWART = ""
      'msg_IWMS_2001.ZPST_WMS = ""
      'msg_IWMS_2001.ERDAT_WMS = ""
      'msg_IWMS_2001.ERZET_WMS = ""
      'msg_IWMS_2001.BUDAT = ""
      'msg_IWMS_2001.BLDAT = ""
      'msg_IWMS_2001.BKTXT = ""

      'Dim DATA_LIST As New MSG_IWMS2001.clsDATA_LIST
      'Dim DATA_INFO As New MSG_IWMS2001.clsDATA_LIST.clsDATA_INFO
      'DATA_LIST.DATA_INFO = New List(Of MSG_IWMS2001.clsDATA_LIST.clsDATA_INFO)

      'DATA_INFO.ZORDER_ITEM = ""
      'DATA_INFO.MATNR = ""
      'DATA_INFO.WERKS = ""
      'DATA_INFO.LGORT = ""
      'DATA_INFO.CHARG = ""
      'DATA_INFO.BWART = ""
      'DATA_INFO.ERFMG = ""
      'DATA_INFO.MEINS = ""
      'DATA_INFO.GRUND = ""
      'DATA_INFO.SOBKZ = ""
      'DATA_INFO.ZSPEC_STOCK_RELATION = ""
      'DATA_INFO.ELIKZ = ""
      'DATA_INFO.INSMK = ""
      'DATA_INFO.KZEAR = ""
      'DATA_INFO.CHARG_TO = ""
      'DATA_INFO.WERKS_TO = ""
      'DATA_INFO.LGORT_TO = ""
      'DATA_INFO.ZSPEC_STOCK_RELATION_TO = ""
      'DATA_INFO.SGTXT = ""
      'DATA_INFO.EBELN = ""
      'DATA_INFO.EBELP = ""
      'DATA_INFO.EBELN_REF = ""
      'DATA_INFO.EBELP_REF = ""

      'DATA_LIST.DATA_INFO.Add(DATA_INFO)
      '      If Receive_Msg.Body.ReceiptList Is Nothing Or Receive_Msg.Body.ReceiptList.ReceiptInfo.Count = 0 Then
      '        ret_strResultMsg = "WMS 给的收料资讯有缺(ReceiptList)，无法回報。"
      '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '        Return False
      '      End If
      '      Dim ID_Trans As String = GetNewID_Trans()    'task id
      '      Dim STS_Trans As String = "1"                     '傳送狀態
      '      Dim dicPO As New Dictionary(Of String, clsPO)
      '      Dim dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      '      Dim dicWO As New Dictionary(Of String, clsWO)
      '      Dim TO_OWNER_NO As String = ""
      '      Dim PO_ID As String = ""
      '      Dim PO_SERIAL_NO As String = ""
      '      Dim WO_ID As String = ""
      '      Dim TO_SUB_OWNER_NO As String = ""
      '      Dim dat_result As String = ""               'WMS工單日期
      '      Dim state_result As String = "1"             '計錄收貨狀態，只要有其中一筆還沒收或即設0 '1:完成點收, 0:未完成點收
      '      Dim trans_control_ID As String = ""
      '      Dim ID_Owner As String = ""               '貨主編號
      '      Dim id_sub As String = ""                 '事業單位
      '      Dim num_buy As String = ""                '採購單號
      '      Dim lin_buy As String = ""                '採購單項次
      '      Dim num_result As String = ""             'WMS工單單號
      '      Dim lin_result As String = ""             'WMS工單項次
      '      Dim cod_item As String = ""               '商品編號
      '      Dim loc_result As String = ""             '驗收儲位/區
      '      Dim qty_result = ""                       '驗收量(實收量)
      '      Dim ser_pcs = ""                          '批號
      '      Dim dat_expiry = ""                       '效期
      '      Dim DAT_INSERT As String = Now_Time                 '寫入日期時間
      '      Dim DAT_POST As String = ""                         '取走日期時間

      '      '退倉
      '      Dim num_rtn As String = ""                  '訂單單號
      '      Dim state_end As String = ""                '結案狀態
      '      Dim cod_cust As String = ""                 '客戶編號
      '      Dim rtn_reason As String = ""               '託運號碼
      '      Dim lin_rtn As String = ""                  '訂單項次
      '      Dim typ_working As String = ""              '處理方式
      '      Dim memo As String = ""                     '後續處理註記
      '      Dim PO_TYPE1 As enuPOType_1 = enuPOType_1.Combination_in
      '      Dim PO_TYPE2 As enuPOType_2 = enuPOType_2.malfunction_in
      '      Dim CUSTOMER_NO As String = ""
      '      Dim H_PO15 As String = ""
      '      Dim H_PO13 As String = ""

      '      '貨故
      '      Dim num_brow = ""
      '      Dim lin_brow = ""
      '      Dim cls_brow = ""
      '      Dim unt_stk = ""

      '      Dim Wo_Serial_No_Index As Integer = 0                                 'Vito_20109

      '      For Each objReceiptInfo In Receive_Msg.Body.ReceiptList.ReceiptInfo
      '        PO_ID = objReceiptInfo.PO_ID
      '        PO_SERIAL_NO = objReceiptInfo.PO_SERIAL_NO
      '        WO_ID = objReceiptInfo.WO_ID
      '        Dim WO_SERIAL_NO As String = objReceiptInfo.WO_SERIAL_NO
      '        TO_OWNER_NO = objReceiptInfo.TO_OWNER_NO
      '        TO_SUB_OWNER_NO = objReceiptInfo.TO_SUB_OWNER_NO
      '        Dim RECEIPT_DATE As String = objReceiptInfo.RECEIPT_DATE
      '        Dim SKU_NO As String = objReceiptInfo.SKU_NO
      '        Dim PACKAGE_ID As String = objReceiptInfo.PACKAGE_ID
      '        Dim QTY As String = objReceiptInfo.QTY
      '        Dim LOT_NO As String = objReceiptInfo.LOT_NO
      '        Dim ITEM_COMMON1 As String = objReceiptInfo.ITEM_COMMON1
      '        Dim ITEM_COMMON2 As String = objReceiptInfo.ITEM_COMMON2
      '        Dim ITEM_COMMON3 As String = objReceiptInfo.ITEM_COMMON3
      '        Dim ITEM_COMMON4 As String = objReceiptInfo.ITEM_COMMON4
      '        Dim ITEM_COMMON5 As String = objReceiptInfo.ITEM_COMMON5
      '        Dim ITEM_COMMON6 As String = objReceiptInfo.ITEM_COMMON6
      '        Dim ITEM_COMMON7 As String = objReceiptInfo.ITEM_COMMON7
      '        Dim ITEM_COMMON8 As String = objReceiptInfo.ITEM_COMMON8
      '        Dim ITEM_COMMON9 As String = objReceiptInfo.ITEM_COMMON9
      '        Dim ITEM_COMMON10 As String = objReceiptInfo.ITEM_COMMON10
      '        Dim SORT_ITEM_COMMON1 As String = objReceiptInfo.SORT_ITEM_COMMON1
      '        Dim SORT_ITEM_COMMON2 As String = objReceiptInfo.SORT_ITEM_COMMON2
      '        Dim SORT_ITEM_COMMON3 As String = objReceiptInfo.SORT_ITEM_COMMON3
      '        Dim SORT_ITEM_COMMON4 As String = objReceiptInfo.SORT_ITEM_COMMON4
      '        Dim SORT_ITEM_COMMON5 As String = objReceiptInfo.SORT_ITEM_COMMON5
      '        Dim CONTRACT_NO As String = objReceiptInfo.CONTRACT_NO
      '        Dim CONTRACT_SERIAL_NO As String = objReceiptInfo.CONTRACT_SERIAL_NO
      '        Dim STORAGE_TYPE As String = objReceiptInfo.STORAGE_TYPE
      '        Dim BND As String = objReceiptInfo.BND
      '        Dim QC_STATUS As String = objReceiptInfo.QC_STATUS
      '        Dim MANUFACETURE_DATE As String = objReceiptInfo.MANUFACETURE_DATE
      '        Dim EXPIRED_DATE As String = objReceiptInfo.EXPIRED_DATE
      '        Dim EFFECTIVE_DATE As String = objReceiptInfo.EFFECTIVE_DATE
      '        Dim COMMENTS As String = objReceiptInfo.COMMENTS

      '        state_result = "1"  '1:完成點收, 0:未完成點收

      '        '取得PO
      '        Dim PO_QTY As Integer = 0
      '        Dim tmp_loc_result As String = ""
      '        gMain.objHandling.O_Get_dicPOByPO_ID(PO_ID, dicPO)
      '        For Each objPO In dicPO.Values
      '          PO_TYPE1 = objPO.PO_Type1
      '          PO_TYPE2 = objPO.PO_Type2
      '          CUSTOMER_NO = objPO.Customer_No
      '          H_PO13 = objPO.H_PO13
      '          H_PO15 = objPO.H_PO15
      '          If objPO.PO_Type2 = enuPOType_2.m_material_in Then
      '            '手工單不回報
      '            Return True
      '          End If
      '          If objPO.PO_Type1 = enuPOType_1.Combination_in And (PO_TYPE2 = enuPOType_2.Retn Or PO_TYPE2 = enuPOType_2.malfunction_in Or PO_TYPE2 = enuPOType_2.material_in) Then  'Vito_20109_2
      '            '採購單與退倉單需回報
      '            If PO_TYPE2 = enuPOType_2.material_in Then
      '              tmp_loc_result = "M000"
      '            ElseIf PO_TYPE2 = enuPOType_2.Retn Or PO_TYPE2 = enuPOType_2.malfunction_in Then

      '            End If
      '            'Vito_20109_2
      '          Else
      '            '其他從採購單外，入庫上架作業的單據不回報
      '            Return True
      '          End If                                                                                            'Vito_20109_2
      '        Next
      '        '取得PO_DTL
      '        Dim tmp_dicPOID As New Dictionary(Of String, String)
      '        tmp_dicPOID.Add(PO_ID, PO_ID)
      '        gMain.objHandling.O_Get_dicPODTLBydicPO_ID(tmp_dicPOID, dicPO_DTL)


      '        '取得WO
      '        gMain.objHandling.O_Get_dicWOByWO_ID(WO_ID, dicWO)

      '        For Each objWO In dicWO.Values
      '          dat_result = objWO.CREATE_TIME
      '        Next


      '        Dim PO_ID_str As String() = Split(PO_ID, "_")
      '        If PO_TYPE1 = enuPOType_1.Combination_in And PO_TYPE2 = enuPOType_2.material_in Then
      '#Region "採購單"
      '          'Rcvd
      '          ID_Owner = TO_OWNER_NO                '貨主編號
      '          id_sub = TO_SUB_OWNER_NO              '事業單位
      '          'Dim PO_ID_str As String() = Split(PO_ID, "_")
      '          num_buy = PO_ID_str(0)                '採購單號
      '          lin_buy = PO_SERIAL_NO                '採購單項次
      '          num_result = WO_ID                    'WMS工單單號
      '          'Vito_20109 lin_result = WO_SERIAL_NO             'WMS工單項次
      '          Dim cod_item_str As String() = Split(SKU_NO, "_")
      '          cod_item = cod_item_str(0).ToString   '商品編號
      '          loc_result = tmp_loc_result           '驗收儲位/區                                                       'Vito_20109_2
      '          ser_pcs = LOT_NO                                '批號
      '          If IsDate(EXPIRED_DATE) Then
      '            dat_expiry = EXPIRED_DATE                       '效期
      '          Else
      '            dat_expiry = "NULL"                       '效期
      '          End If

      '          DAT_INSERT = Now_Time                 '寫入日期時間
      '          DAT_POST = ""                         '取走日期時間

      '          'PO_ID PO_SERIAL WO_ID WO_SERIAL_NO SKU_NO相同，但是LOT_NO和EXPIRED_DATE不同時加 1
      '          '找出已存的資料中有沒有重覆的
      '          Dim tmpkey As String = ""                                                                               'Vito_20109
      '          Dim objdic As clsRcvd = Nothing                                                                         'Vito_20109
      '          Dim Find_f As Boolean = False                                                                           'Vito_20109
      '          For index As Integer = 0 To Wo_Serial_No_Index Step 1                                                   'Vito_20109
      '            tmpkey = clsRcvd.Get_Combination_Key(ID_Trans, TO_OWNER_NO, num_buy, PO_SERIAL_NO, WO_ID, index)      'Vito_20109
      '            If ret_dicAddW2E_Rcvd.TryGetValue(tmpkey, objdic) Then                                                'Vito_20109
      '              If objdic.ser_pcs = LOT_NO And objdic.dat_expiry = EXPIRED_DATE And objdic.cod_item = SKU_NO Then   'Vito_20109
      '                '有找到重覆的話，數量要加總                                                                       'Vito_20109
      '                Dim tmp As Integer = Val(objdic.qty_result) + Val(QTY)                                            'Vito_20109
      '                qty_result = tmp.ToString                                                                         'Vito_20109
      '                Find_f = True                                                                                     'Vito_20109
      '              End If                                                                                              'Vito_20109
      '            End If                                                                                                'Vito_20109
      '          Next                                                                                                    'Vito_20109
      '          If Find_f = False Then                                                                                  'Vito_20109
      '            Wo_Serial_No_Index += 1                                                                               'Vito_20109
      '            qty_result = CInt(QTY).ToString                                                                                      'Vito_20109
      '          End If                                                                                                  'Vito_20109
      '          '檢查是否點收完成，抓PO_DTL的總數比對                                                                   'Vito_20109
      '          For Each objPO_DTL In dicPO_DTL.Values                                                                  'Vito_20109
      '            'Vito_20113_2 If objPO_DTL.PO_SERIAL_NO = PO_SERIAL_NO Then                                                         'Vito_20109
      '            If objPO_DTL.QTY <> objPO_DTL.QTY_PROCESS Then                                                      'Vito_20109
      '              state_result = "0"                                                                                'Vito_20109
      '            End If                                                                                              'Vito_20109
      '            'Vito_20113_2 End If                                                                                                'Vito_20109
      '          Next                                                                                                    'Vito_20109
      '          lin_result = Wo_Serial_No_Index             'WMS工單項次                                                'Vito_20109

      '          Dim objNewW2E_Rcvd As New clsRcvd(ID_Trans, STS_Trans, ID_Owner, id_sub, num_buy, lin_buy, num_result, lin_result, cod_item, loc_result, qty_result, ser_pcs, dat_expiry, dat_result, DAT_INSERT, DAT_POST)
      '          If ret_dicAddW2E_Rcvd.ContainsKey(objNewW2E_Rcvd.gid) = False Then '新增
      '            ret_dicAddW2E_Rcvd.Add(objNewW2E_Rcvd.gid, objNewW2E_Rcvd)
      '          Else '更新
      '            Dim temp_Dic As clsRcvd = Nothing
      '            If ret_dicAddW2E_Rcvd.TryGetValue(objNewW2E_Rcvd.gid, temp_Dic) Then
      '              If temp_Dic.ser_pcs = objNewW2E_Rcvd.ser_pcs And temp_Dic.dat_expiry = objNewW2E_Rcvd.dat_expiry Then
      '                ret_dicAddW2E_Rcvd.Remove(objNewW2E_Rcvd.gid)
      '                ret_dicAddW2E_Rcvd.Add(objNewW2E_Rcvd.gid, objNewW2E_Rcvd)
      '              End If
      '            Else '不可能有Else
      '            End If
      '          End If
      '#End Region
      '        ElseIf PO_TYPE1 = enuPOType_1.Combination_in And PO_TYPE2 = enuPOType_2.Retn Then
      '#Region "退倉單 Vito_20211"
      '          ID_Owner = TO_OWNER_NO 'PO_ID_str(3)        '貨主編號
      '          id_sub = TO_SUB_OWNER_NO 'PO_ID_str(4)      '事業單位
      '          cod_cust = CUSTOMER_NO                      '客戶編號
      '          num_rtn = PO_ID_str(0)                      '退倉單號
      '          lin_rtn = PO_SERIAL_NO                      '退倉項次
      '          Dim cod_item_str As String() = Split(SKU_NO, "_")
      '          cod_item = cod_item_str(0)                  '商品編號
      '          num_result = PO_ID_str(0)                   '改為退倉單號 
      '          WO_SERIAL_NO = gMain.objHandling.O_Get_WOSerialNObyPOID_PO_SERIAL_NO(PO_ID, PO_SERIAL_NO)
      '          lin_result = PO_SERIAL_NO                   'WMS工單項次
      '          qty_result = CInt(QTY).ToString             '實退量
      '          ser_pcs = LOT_NO                            '批號
      '          dat_expiry = ""                             '效期
      '          typ_working = ""                            '處理方式

      '          Dim dicWO_DTL As New Dictionary(Of String, clsWO_DTL)
      '          gMain.objHandling.O_GetDB_dicWODTLByWOID_WOSerialNo(WO_ID, WO_SERIAL_NO, dicWO_DTL)
      '          If dicWO_DTL.Any Then
      '            For Each objWO_DTL In dicWO_DTL.Values
      '              QC_STATUS = objWO_DTL.QC_STATUS
      '            Next
      '          Else
      '            ret_strResultMsg = "WMS 给的WO_ID查不到相對應WO_DTL，无法回報。"
      '            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '            Return False
      '          End If
      '          If QC_STATUS = enuQCStatus.D Then
      '            typ_working = "A"
      '          ElseIf QC_STATUS = enuQCStatus.E Then
      '            typ_working = "D"
      '          ElseIf QC_STATUS = enuQCStatus.B Then
      '            typ_working = "E"
      '          ElseIf QC_STATUS = enuQCStatus.C Then
      '            typ_working = "F"
      '          End If


      '          If IsDate(EXPIRED_DATE) Then
      '            dat_expiry = EXPIRED_DATE
      '          Else
      '            dat_expiry = "NULL"
      '          End If
      '          If typ_working = "F" Then
      '            loc_result = "M999"                             '驗收儲位
      '          Else
      '            loc_result = "M998"                             '驗收儲位
      '          End If
      '          dat_result = Now_Time                         '出貨日期時間
      '          memo = H_PO15                               '後續處理註記
      '          DAT_INSERT = Now_Time                       '寫入日期時間
      '          DAT_POST = "NULL"                           '取走日期時間

      '          Dim objNewW2E_RetnResult_Item As New clsRetnResultItem(ID_Trans, STS_Trans, ID_Owner, id_sub, cod_cust, num_rtn, lin_rtn, cod_item, num_result, lin_result, qty_result, ser_pcs, dat_expiry,
      '                                                                 typ_working, loc_result, dat_result, memo, DAT_INSERT, DAT_POST)
      '          Dim objW2E_RetnResult_Item As clsRetnResultItem = Nothing

      '          If ret_dicAddW2E_RetnResult_item.TryGetValue(objNewW2E_RetnResult_Item.gid, objW2E_RetnResult_Item) = True Then
      '            Dim old_qty = CInt(objW2E_RetnResult_Item.qty_Result)
      '            Dim new_qty = CInt(objNewW2E_RetnResult_Item.qty_Result)
      '            Dim total_qty = old_qty + new_qty
      '            objW2E_RetnResult_Item.qty_Result = total_qty.ToString
      '          Else
      '            ret_dicAddW2E_RetnResult_item.Add(objNewW2E_RetnResult_Item.gid, objNewW2E_RetnResult_Item)
      '          End If

      '          'If ret_dicAddW2E_RetnResult_item.ContainsKey(objNewW2E_RetnResult_Item.gid) = False Then
      '          '  ret_dicAddW2E_RetnResult_item.Add(objNewW2E_RetnResult_Item.gid, objNewW2E_RetnResult_Item)
      '          'End If

      '          'Wo_Serial_No_Index += 1
      '          'Next
      '#End Region
      '#Region "貨故單 Vito_20B11"
      '        ElseIf PO_TYPE1 = enuPOType_1.Combination_in And (PO_TYPE2 = enuPOType_2.malfunction_in) Then
      '          'W2E_BackResult_Item
      '          'Dim PO_ID_str As String() = Split(PO_ID, "_")
      '          ID_Owner = PO_ID_str(3)                     '貨主編號
      '          id_sub = PO_ID_str(4)                       '事業單位
      '          num_brow = PO_ID_str(0)                     'ERP貨故單號
      '          lin_brow = PO_SERIAL_NO                     'ERP貨故項次
      '          cls_brow = Getcls_browByItem_Common3(ITEM_COMMON3)  '貨故類型
      '          num_result = WO_ID            'WMS工單號
      '          WO_SERIAL_NO = gMain.objHandling.O_Get_WOSerialNObyPOID_PO_SERIAL_NO(PO_ID, PO_SERIAL_NO)
      '          lin_result = WO_SERIAL_NO             'WMS項次
      '          Dim cod_item_str As String() = Split(SKU_NO, "_")
      '          cod_item = cod_item_str(0)             '商品編號
      '          qty_result = CInt(QTY).ToString              '數量
      '          unt_stk = GetSKU_UNITbySKU_No(SKU_NO)        '單位
      '          ser_pcs = LOT_NO                            '批號
      '          dat_expiry = ""                             '效期

      '          Dim tmpPO_ID As New Dictionary(Of String, String)
      '          Dim tmpPO_DTL As New Dictionary(Of String, clsPO_DTL)
      '          tmpPO_ID.Add(PO_ID, PO_ID)


      '          If IsDate(EXPIRED_DATE) Then
      '            dat_expiry = EXPIRED_DATE
      '          Else
      '            dat_expiry = "NULL"
      '          End If
      '          DAT_INSERT = Now_Time                       '寫入日期時間
      '          DAT_POST = "NULL"                           '取走日期時間

      '          Dim objNewW2E_Damage_Item As New clsW2EDamageItem(ID_Trans, STS_Trans, ID_Owner, id_sub, num_brow, lin_brow, cls_brow, num_result, lin_result, cod_item, qty_result, unt_stk,
      '                                                          ser_pcs, dat_expiry, DAT_INSERT, DAT_POST)
      '          Dim objW2E_Damage_Item As clsW2EDamageItem = Nothing

      '          If ret_dicAddW2E_Damage_Item.TryGetValue(objNewW2E_Damage_Item.gid, objW2E_Damage_Item) = True Then
      '            '已存在
      '            Dim old_qty = CInt(objW2E_Damage_Item.qty_Result)
      '            Dim new_qty = CInt(objNewW2E_Damage_Item.qty_Result)
      '            Dim total_qty = old_qty + new_qty
      '            objW2E_Damage_Item.qty_Result = total_qty.ToString
      '          Else
      '            '不存在
      '            ret_dicAddW2E_Damage_Item.Add(objNewW2E_Damage_Item.gid, objNewW2E_Damage_Item)
      '          End If
      '#End Region
      '        End If

      '      Next

      '      Dim line_count
      '      'Rcvm 'Vito_20109 01/09改成一對多個Rcvd回報
      '      For Each objPO In dicPO.Values
      '#Region "採購單_表頭"
      '        If PO_TYPE1 = enuPOType_1.Combination_in And PO_TYPE2 = enuPOType_2.material_in Then
      '          'Dim STS_Trans As String = "0"          '傳送狀態
      '          ID_Owner = TO_OWNER_NO                  '貨主編號
      '          id_sub = TO_SUB_OWNER_NO                '事業單位
      '          num_buy = objPO.PO_KEY1                 '採購單號
      '          'Dim state_end As String = ""            '結案狀態
      '          If state_result = "1" Then
      '            state_end = "Y"
      '          Else
      '            state_end = "N"
      '          End If
      '          Dim typ_buy As String = ""              '採購單類型
      '          num_result = WO_ID                      'WMS工單單號
      '          dat_result = Now_Date          'WMS工單日期
      '          'Dim memo As String = ""                 '註記事項
      '          line_count = ret_dicAddW2E_Rcvd.Count '明細筆數
      '          DAT_INSERT = Now_Time                   '寫入日期時間
      '          DAT_POST = ""                           '取走日期時間

      '          Dim objNewW2E_Rcvm As New clsRcvm(ID_Trans, STS_Trans, ID_Owner, id_sub, num_buy, state_end, typ_buy, num_result, dat_result, memo, line_count, DAT_INSERT, DAT_POST)
      '          If ret_dicAddW2E_Rcvm.ContainsKey(objNewW2E_Rcvm.gid) = False Then
      '            ret_dicAddW2E_Rcvm.Add(objNewW2E_Rcvm.gid, objNewW2E_Rcvm)
      '          End If
      '#End Region
      '#Region "退倉單_表頭"
      '        ElseIf PO_TYPE1 = enuPOType_1.Combination_in And PO_TYPE2 = enuPOType_2.Retn Then
      '          'W2E_RetnResult_Head
      '          Dim PO_ID_str As String() = Split(PO_ID, "_")
      '          ID_Owner = PO_ID_str(3)                       '貨主編號
      '          id_sub = PO_ID_str(4)                         '事業單位
      '          num_rtn = PO_ID_str(0)                        '退倉單號 
      '          state_end = "N"                               '結案狀態
      '          cod_cust = CUSTOMER_NO                        '客戶編號
      '          rtn_reason = H_PO13                           '退倉原因
      '          num_result = WO_ID                            'WMS驗收單號
      '          dat_result = Now_Date                           '驗收日期
      '          line_count = ret_dicAddW2E_RetnResult_item.Count '明細筆數
      '          DAT_INSERT = Now_Time                         '寫入日期時間
      '          DAT_POST = "NULL"                             '取走日期時間

      '          Dim objNewW2E_RetnResult_Head As New clsRetnResultHead(ID_Trans, STS_Trans, ID_Owner, id_sub, num_rtn, state_end, cod_cust, rtn_reason, num_result, dat_result, line_count, DAT_INSERT, DAT_POST)
      '          If ret_dicAddW2E_RetnResult_Head.ContainsKey(objNewW2E_RetnResult_Head.gid) = False Then
      '            ret_dicAddW2E_RetnResult_Head.Add(objNewW2E_RetnResult_Head.gid, objNewW2E_RetnResult_Head)
      '          End If
      '#End Region
      '#Region "貨故單 表頭"
      '        ElseIf PO_TYPE1 = enuPOType_1.Combination_in And (PO_TYPE2 = enuPOType_2.malfunction_in) Then
      '          'W2E_Damage_Head
      '          Dim PO_ID_str As String() = Split(PO_ID, "_")
      '          ID_Owner = PO_ID_str(3)                       '貨主編號
      '          id_sub = PO_ID_str(4)                         '事業單位
      '          state_end = "Y"                               '結案狀態
      '          num_brow = PO_ID_str(0)                       'ERP貨故單號
      '          num_result = WO_ID                            'WMS工單號
      '          dat_result = Now_Date                         'WMS工單日期
      '          line_count = ret_dicAddW2E_Damage_Item.Count  '明細筆數
      '          DAT_INSERT = Now_Time                         '寫入日期時間
      '          DAT_POST = "NULL"                             '取走日期時間

      '          Dim objNewW2E_Damage_Head As New clsW2EDamageHead(ID_Trans, STS_Trans, ID_Owner, id_sub, state_end, num_brow, num_result, dat_result, line_count, DAT_INSERT, DAT_POST)
      '          If ret_dicAddW2E_Damage_Head.ContainsKey(objNewW2E_Damage_Head.gid) = False Then
      '            ret_dicAddW2E_Damage_Head.Add(objNewW2E_Damage_Head.gid, objNewW2E_Damage_Head)
      '          End If
      '#End Region
      '        End If



      '      Next


      '      If PO_TYPE1 = enuPOType_1.Combination_in And PO_TYPE2 = enuPOType_2.material_in Then
      '#Region "採購單"
      '        If ret_dicAddW2E_Rcvm.Count <> 0 Then
      '          'trans_control Rcvm
      '          trans_control_ID = "1"
      '          Dim Auto_Seq As String = trans_control_ID
      '          'Dim ID_Trans As String = "W" & Now_Time & UUID       'task id
      '          Dim Table_Trans As String = "W2E_Rcvm"                'table name
      '          Dim Type_Trans As String = "W"                        '傳送類型
      '          Dim Dat_Trans As String = Now_Time                    '傳送時間
      '          Dim Count_Trans As String = ret_dicAddW2E_Rcvm.Count  '傳送完成筆數
      '          Dim Count_ERR As String = "0"                         '接收錯誤筆數
      '          'Dim STS_Trans As String = "1"                        '傳送狀態
      '          Dim Message_Trans As String = ""                      '訊息說明
      '          Dim objNewTrans_Control As New clsTrans_Control(Auto_Seq, ID_Trans, Table_Trans, Type_Trans, Dat_Trans, Count_Trans, Count_ERR, STS_Trans, Message_Trans)
      '          If ret_dicAddTrans_Control.ContainsKey(objNewTrans_Control.gid) = False Then
      '            ret_dicAddTrans_Control.Add(objNewTrans_Control.gid, objNewTrans_Control)
      '          End If
      '        End If
      '        If ret_dicAddW2E_Rcvd.Count <> 0 Then
      '          'trans_control Rcvd
      '          trans_control_ID = "2"
      '          Dim Auto_Seq_d As String = trans_control_ID
      '          'Dim ID_Trans As String = "W" & Now_Time & UUID         'task id
      '          Dim Table_Trans_d As String = "W2E_Rcvd"                'table name
      '          Dim Type_Trans_d As String = "W"                        '傳送類型
      '          Dim Dat_Trans_d As String = Now_Time                    '傳送時間
      '          Dim Count_Trans_d As String = ret_dicAddW2E_Rcvd.Count  '傳送完成筆數
      '          Dim Count_ERR_d As String = "0"                         '接收錯誤筆數
      '          'Dim STS_Trans As String = "1"                          '傳送狀態
      '          Dim Message_Trans_d As String = ""                      '訊息說明
      '          Dim objNewTrans_Control_d As New clsTrans_Control(Auto_Seq_d, ID_Trans, Table_Trans_d, Type_Trans_d, Dat_Trans_d, Count_Trans_d, Count_ERR_d, STS_Trans, Message_Trans_d)
      '          If ret_dicAddTrans_Control.ContainsKey(objNewTrans_Control_d.gid) = False Then
      '            ret_dicAddTrans_Control.Add(objNewTrans_Control_d.gid, objNewTrans_Control_d)
      '          End If
      '        End If

      '      End If
      '#End Region
      '      If PO_TYPE1 = enuPOType_1.Combination_in And PO_TYPE2 = enuPOType_2.Retn Then
      '#Region "退倉單"
      '        If ret_dicAddW2E_RetnResult_Head.Count <> 0 Then
      '          'trans_control Rcvm
      '          trans_control_ID = "1"
      '          Dim Auto_Seq As String = trans_control_ID
      '          'Dim ID_Trans As String = "W" & Now_Time & UUID       'task id
      '          Dim Table_Trans As String = "W2E_RetnResult_Head"                'table name
      '          Dim Type_Trans As String = "W"                        '傳送類型
      '          Dim Dat_Trans As String = Now_Time                    '傳送時間
      '          Dim Count_Trans As String = ret_dicAddW2E_RetnResult_Head.Count  '傳送完成筆數
      '          Dim Count_ERR As String = "0"                         '接收錯誤筆數
      '          'Dim STS_Trans As String = "1"                        '傳送狀態
      '          Dim Message_Trans As String = ""                      '訊息說明
      '          Dim objNewTrans_Control As New clsTrans_Control(Auto_Seq, ID_Trans, Table_Trans, Type_Trans, Dat_Trans, Count_Trans, Count_ERR, STS_Trans, Message_Trans)
      '          If ret_dicAddTrans_Control.ContainsKey(objNewTrans_Control.gid) = False Then
      '            ret_dicAddTrans_Control.Add(objNewTrans_Control.gid, objNewTrans_Control)
      '          End If
      '        End If

      '        If ret_dicAddW2E_RetnResult_item.Count <> 0 Then
      '          'trans_control Rcvd
      '          trans_control_ID = "2"
      '          Dim Auto_Seq_d As String = trans_control_ID
      '          'Dim ID_Trans As String = "W" & Now_Time & UUID         'task id
      '          Dim Table_Trans_d As String = "W2E_RetnResult_Item"                'table name
      '          Dim Type_Trans_d As String = "W"                        '傳送類型
      '          Dim Dat_Trans_d As String = Now_Time                    '傳送時間
      '          Dim Count_Trans_d As String = ret_dicAddW2E_RetnResult_item.Count  '傳送完成筆數
      '          Dim Count_ERR_d As String = "0"                         '接收錯誤筆數
      '          'Dim STS_Trans As String = "1"                          '傳送狀態
      '          Dim Message_Trans_d As String = ""                      '訊息說明
      '          Dim objNewTrans_Control_d As New clsTrans_Control(Auto_Seq_d, ID_Trans, Table_Trans_d, Type_Trans_d, Dat_Trans_d, Count_Trans_d, Count_ERR_d, STS_Trans, Message_Trans_d)
      '          If ret_dicAddTrans_Control.ContainsKey(objNewTrans_Control_d.gid) = False Then
      '            ret_dicAddTrans_Control.Add(objNewTrans_Control_d.gid, objNewTrans_Control_d)
      '          End If
      '        End If

      '#End Region

      '      End If

      '      If PO_TYPE1 = enuPOType_1.Combination_in And (PO_TYPE2 = enuPOType_2.malfunction_in) Then
      '#Region "貨故單"
      '        If ret_dicAddW2E_Damage_Head.Count <> 0 Then
      '          'trans_control Rcvm
      '          trans_control_ID = "1"
      '          Dim Auto_Seq As String = trans_control_ID
      '          'Dim ID_Trans As String = "W" & Now_Time & UUID       'task id
      '          Dim Table_Trans As String = "W2E_Damage_Head"         'table name
      '          Dim Type_Trans As String = "W"                        '傳送類型
      '          Dim Dat_Trans As String = Now_Time                    '傳送時間
      '          Dim Count_Trans As String = ret_dicAddW2E_Damage_Head.Count  '傳送完成筆數
      '          Dim Count_ERR As String = "0"                         '接收錯誤筆數
      '          'Dim STS_Trans As String = "1"                        '傳送狀態
      '          Dim Message_Trans As String = ""                      '訊息說明
      '          Dim objNewTrans_Control As New clsTrans_Control(Auto_Seq, ID_Trans, Table_Trans, Type_Trans, Dat_Trans, Count_Trans, Count_ERR, STS_Trans, Message_Trans)
      '          If ret_dicAddTrans_Control.ContainsKey(objNewTrans_Control.gid) = False Then
      '            ret_dicAddTrans_Control.Add(objNewTrans_Control.gid, objNewTrans_Control)
      '          End If
      '        End If

      '        If ret_dicAddW2E_Damage_Item.Count <> 0 Then
      '          'trans_control Rcvd
      '          trans_control_ID = "2"
      '          Dim Auto_Seq_d As String = trans_control_ID
      '          'Dim ID_Trans As String = "W" & Now_Time & UUID         'task id
      '          Dim Table_Trans_d As String = "W2E_Damage_Item"         'table name
      '          Dim Type_Trans_d As String = "W"                        '傳送類型
      '          Dim Dat_Trans_d As String = Now_Time                    '傳送時間
      '          Dim Count_Trans_d As String = ret_dicAddW2E_Damage_Item.Count  '傳送完成筆數
      '          Dim Count_ERR_d As String = "0"                         '接收錯誤筆數
      '          'Dim STS_Trans As String = "1"                          '傳送狀態
      '          Dim Message_Trans_d As String = ""                      '訊息說明
      '          Dim objNewTrans_Control_d As New clsTrans_Control(Auto_Seq_d, ID_Trans, Table_Trans_d, Type_Trans_d, Dat_Trans_d, Count_Trans_d, Count_ERR_d, STS_Trans, Message_Trans_d)
      '          If ret_dicAddTrans_Control.ContainsKey(objNewTrans_Control_d.gid) = False Then
      '            ret_dicAddTrans_Control.Add(objNewTrans_Control_d.gid, objNewTrans_Control_d)
      '          End If
      '        End If
      '#End Region
      '      End If


      Return True

    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef lstSql As List(Of String), ByRef host_command As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try
      For Each objHost In host_command.Values
        If objHost.O_Add_Insert_SQLString(lstSql) = False Then
          Result_Message = "get insert HOST_T_COMMAND SQL failed"
          Return False
        End If
      Next
      'For Each objtrans_control In dicAddTrans_Control.Values
      '  If objtrans_control.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "get insert trans_control sql failed"
      '    Return False
      '  End If
      'Next
      'For Each objRcvm In dicAddW2E_Rcvm.Values
      '  If objRcvm.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_Rcvm SQL Failed"
      '    Return False
      '  End If
      'Next
      'For Each objRcvd In dicAddW2E_Rcvd.Values
      '  If objRcvd.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_Rcvd SQL Failed"
      '    Return False
      '  End If
      'Next

      'For Each objW2E_RetnResult_Head In ret_dicAddW2E_RetnResult_Head.Values
      '  If objW2E_RetnResult_Head.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_RetnResult_Head SQL Failed"
      '    Return False
      '  End If
      'Next
      'For Each objW2E_RetnResult_Item In ret_dicAddW2E_RetnResult_item.Values
      '  If objW2E_RetnResult_Item.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_RetnResult_Item SQL Failed"
      '    Return False
      '  End If
      'Next

      'For Each obj In ret_dicAddW2E_Damage_Head.Values
      '  If obj.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_Damage_Head SQL Failed"
      '    Return False
      '  End If
      'Next
      'For Each obj In ret_dicAddW2E_Damage_Item.Values
      '  If obj.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_Damage_Item SQL Failed"
      '    Return False
      '  End If
      'Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行新增的Carrier和Carrier_Status的SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef Result_Message As String,
                                         ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL   更新杏一的資料庫
      If lstSql.Any = True Then
        If Common_DBManagement.BatchUpdate(lstSql) = False Then
          '更新DB失敗則回傳False
          Result_Message = "eHOST 更新资料库失败"
          Return False
        End If
      End If
      'Common_DBManagement.AddQueued(lstQueueSql)

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Getcls_browByItem_Common3(ByVal ITEM_COMMON3 As String) As String

    Select Case ITEM_COMMON3
      Case "21"
        Return "1"  '遺失
      Case "22"
        Return "2"  '破損
      Case "23"
        Return "3"  '浸水
    End Select
    Return "0"
  End Function
  Private Function CombinationBatchID(ByVal nowtime As String, ByVal BatchNum As String)
    Try
      Dim ret As String = ""
      ret = "R" + nowtime + BatchNum
      Return ret
    Catch ex As Exception
      Return ""
    End Try
  End Function
  Private Function CombinationSKU(ByVal SKUkey1 As String, ByVal SKUkey2 As String)
    Try
      Dim ret_sku As String = ""
      ret_sku = SKUkey1 & "_" & SKUkey2
      Return ret_sku
    Catch ex As Exception
      Return Nothing
    End Try
  End Function
End Module



