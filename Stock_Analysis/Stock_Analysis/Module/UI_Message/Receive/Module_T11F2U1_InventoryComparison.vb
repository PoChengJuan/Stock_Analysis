'20190827
'V1.0.0
'Jerry

'庫存對比

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T11F2U1_InventoryComparison
  Public Function O_Process(ByVal Receive_Msg As MSG_T11F2U1_InventoryComparison,
                                          ByRef ret_strResultMsg As String,
                                       ByRef ret_Wait_UUID As String) As Boolean
    Try
      Dim lstSql As New List(Of String)
      Dim lstQueueSql As New List(Of String)
      Dim dicAddInventoryComparison As New Dictionary(Of String, clsWMS_CT_INVENTORY_COMPARISON)

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料處理
      If Process_Data(Receive_Msg, dicAddInventoryComparison, ret_strResultMsg) = False Then
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, dicAddInventoryComparison, lstSql, lstQueueSql) = False Then
        Return False
      End If
      '執行SQL與更新物件
      If Execute_DataUpdate(ret_strResultMsg, lstSql, lstQueueSql) = False Then
        Return False
      End If

      '產生Excel 發送Mail 'on著等程式觸發
      bln_SendAccountMail = True

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T11F2U1_InventoryComparison,
                              ByRef ret_strResultMsg As String) As Boolean

    Try
      ''先進行資料邏輯檢查
      'For Each objPOInfo In Receive_Msg.Body.POList.POInfo
      '  '資料檢查
      '  Dim PO_ID As String = objPOInfo.PO_ID
      '  Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
      '  Dim FORCED_UPDATE As String = objPOInfo.FORCED_UPDATE
      '  '檢查PO_ID是否為空
      '  If PO_ID = "" Then
      '    ret_strResultMsg = "PO_ID is empty"
      '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If

      'Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

#Region "笨方法"
  '  '資料處理
  '  Private Function Process_Data(ByVal Receive_Msg As MSG_T11F2U1_InventoryComparison,
  '                                ByVal ret_lstAddInventoryComparison As List(Of clsWMS_CT_INVENTORY_COMPARISON),
  '                                ByRef ret_strResultMsg As String) As Boolean
  '    Try
  '      Dim S_DATE As String = GetNewTime_ByDataTimeFormat("yyyyMMdd")
  '      Dim E_DATE As String = GetNewTime_ByDataTimeFormat("yyyyMMdd")
  '      Dim objERP_Report As New MSG_ASRS_getCompStock

  '      '取出全部庫存
  '      Dim tmp_dicTransformPOID As New Dictionary(Of String, String)
  '      Dim tmp_dicTransformPO As New Dictionary(Of String, clsPO)
  '      Dim tmp_dicTransformPO_Line As New Dictionary(Of String, clsPO_LINE)
  '      Dim tmp_dicTransformPO_DTL As New Dictionary(Of String, clsPO_DTL)
  '      Dim tmp_dicTransformPO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)

  '      Dim tmp_dicDischargePOID As New Dictionary(Of String, String)
  '      Dim tmp_dicDischargePO As New Dictionary(Of String, clsPO)
  '      Dim tmp_dicDischargePO_Line As New Dictionary(Of String, clsPO_LINE)
  '      Dim tmp_dicDischargePO_DTL As New Dictionary(Of String, clsPO_DTL)

  '      Dim tmp_dicReceiptPOID As New Dictionary(Of String, String)
  '      Dim tmp_dicReceiptPO As New Dictionary(Of String, clsPO)
  '      Dim tmp_dicReceiptPO_Line As New Dictionary(Of String, clsPO_LINE)
  '      Dim tmp_dicReceiptPO_DTL As New Dictionary(Of String, clsPO_DTL)
  '#Region "取得未結案的入庫單"
  '      '取出未結案的單據，若為入庫單則排除庫存
  '      Dim dicPO As New Dictionary(Of String, clsPO)
  '      gMain.objHandling.O_GetDB_dicPOByALL(dicPO)
  '      tmp_dicReceiptPO = dicPO.Where(Function(q)
  '                                       If q.Value.WO_Type = enuWOType.Receipt Then Return True
  '                                       Return False
  '                                     End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '      For Each objPO In tmp_dicReceiptPO.Values
  '        If tmp_dicReceiptPOID.ContainsKey(objPO.PO_ID) = False Then
  '          tmp_dicReceiptPOID.Add(objPO.PO_ID, objPO.PO_ID)
  '        End If
  '      Next
  '      '使用dicPO取得資料庫裡的PO_Line資料
  '      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicReceiptPOID, tmp_dicReceiptPO_Line) = False Then
  '        ret_strResultMsg = "WMS get PO_Line data From DB Failed"
  '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '      '使用dicPO取得資料庫裡的PO_DTL資料
  '      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicReceiptPOID, tmp_dicReceiptPO_DTL) = False Then
  '        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
  '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '#End Region
  '#Region "取得未結案的出庫單"
  '      '取出未結案的單據，若為出庫單則排除庫存
  '      'dicPO = New Dictionary(Of String, clsPO)
  '      'gMain.objHandling.O_GetDB_dicPOByALL(dicPO)
  '      tmp_dicDischargePO = dicPO.Where(Function(q)
  '                                         If q.Value.WO_Type = enuWOType.Discharge AndAlso q.Value.PO_Type2 = enuPOType_2.material_out Then Return True
  '                                         Return False
  '                                       End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '      For Each objPO In tmp_dicDischargePO.Values
  '        If tmp_dicDischargePOID.ContainsKey(objPO.PO_ID) = False Then
  '          tmp_dicDischargePOID.Add(objPO.PO_ID, objPO.PO_ID)
  '        End If
  '      Next
  '      '使用dicPO取得資料庫裡的PO_Line資料
  '      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicDischargePOID, tmp_dicDischargePO_Line) = False Then
  '        ret_strResultMsg = "WMS get PO_Line data From DB Failed"
  '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '      '使用dicPO取得資料庫裡的PO_DTL資料
  '      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicDischargePOID, tmp_dicDischargePO_DTL) = False Then
  '        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
  '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '#End Region
  '#Region "取得未結案的轉播單"
  '      '取出未結案的單據，若為出庫單則排除庫存
  '      tmp_dicTransformPO = dicPO.Where(Function(q)
  '                                         If q.Value.WO_Type = enuWOType.Transform AndAlso q.Value.PO_Type2 = enuPOType_2.transaction_account Then Return True
  '                                         Return False
  '                                       End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '      For Each objPO In tmp_dicTransformPO.Values
  '        If tmp_dicTransformPOID.ContainsKey(objPO.PO_ID) = False Then
  '          tmp_dicTransformPOID.Add(objPO.PO_ID, objPO.PO_ID)
  '        End If
  '      Next
  '      '使用dicPO取得資料庫裡的PO_Line資料
  '      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicTransformPOID, tmp_dicTransformPO_Line) = False Then
  '        ret_strResultMsg = "WMS get PO_Line data From DB Failed"
  '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '      '使用dicPO取得資料庫裡的PO_DTL資料
  '      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicTransformPOID, tmp_dicTransformPO_DTL) = False Then
  '        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
  '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '      '使用dicPO取得資料庫裡的PO_DTL資料
  '      If gMain.objHandling.O_GetDB_dicPODTLTRANSACTIONBydicPO_ID(tmp_dicTransformPOID, tmp_dicTransformPO_DTL_TRANSACTION) = False Then
  '        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
  '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '#End Region
  '#Region "取得全部庫存"
  '      Dim dicCarrierItem = gMain.objHandling.GetCarrierItemByACCEPTING_STATUS(enuAcceptingStatus.Inventory)
  '#End Region
  '      Dim dicCarrier = gMain.objHandling.GetCarrierStatusByAll

  '      If Mod_WCFHost.ASRS_getCompStock(S_DATE, E_DATE, objERP_Report, ret_strResultMsg) = False Then
  '        Return False
  '      End If

  '      If objERP_Report Is Nothing Then
  '        ret_strResultMsg = "ERP回傳結果序列化後為空"
  '        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '      '開始比對

  '      SendMessageToLog("開始比對", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
  '      For Each objOutput In objERP_Report.output
  '        Dim SKU_NO = objOutput.MATNR
  '        Dim LOT_NO = objOutput.CHARG
  '        Dim ITEM_COMMON1 = ""
  '        Dim ITEM_COMMON2 = ""
  '        Dim ITEM_COMMON3 = ""
  '        Dim ITEM_COMMON4 = ""
  '        Dim ITEM_COMMON5 = ""
  '        Dim ITEM_COMMON6 = ""
  '        Dim ITEM_COMMON7 = ""
  '        Dim ITEM_COMMON8 = ""
  '        Dim ITEM_COMMON9 = ""
  '        Dim ITEM_COMMON10 = ""
  '        Dim SORT_ITEM_COMMON1 = ""
  '        Dim SORT_ITEM_COMMON2 = ""
  '        Dim SORT_ITEM_COMMON3 = ""
  '        Dim SORT_ITEM_COMMON4 = ""
  '        Dim SORT_ITEM_COMMON5 = ""
  '        Dim OWNER_NO = objOutput.WERKS
  '        Dim SUB_OWNER_NO = objOutput.LGORT
  '        Dim STORAGE_TYPE = 0
  '        Dim BND = 0
  '        Dim QC_STATUS = 0
  '        Dim WMS_STOCK_QTY = 0
  '        Dim WMS_UNFINISH_QTY = 0
  '        Dim WMS_COMPARSON_QTY = 0
  '        Dim ERP_STOCK_QTY = IIf(objOutput.LABST = "", 0, objOutput.LABST)
  '        Dim ERP_UNFINISH_QTY = IIf(objOutput.MENGE = "", 0, objOutput.MENGE)
  '        Dim ERP_COMPARSON_QTY = IIf(objOutput.CLABS = "", 0, objOutput.CLABS)
  '        Dim QUANTITY_VARIANCE = 0
  '        Dim ERP_SYSTEM = "ERP"
  '        Dim CREATE_TIME = ""
  '        Dim ACC_COMMON1 = "" '最後收料的棧板
  '        Dim ACC_COMMON2 = "" '該棧板的位置
  '        Dim ACC_COMMON3 = 0 '總托盤數
  '        Dim ACC_COMMON4 = ""
  '        Dim ACC_COMMON5 = ""
  '        Dim ACC_COMMON6 = ""
  '        Dim ACC_COMMON7 = ""
  '        Dim ACC_COMMON8 = ""
  '        Dim ACC_COMMON9 = ""
  '        Dim ACC_COMMON10 = ""


  '#Region "1.WMS全部庫存 (入庫單+現有庫存+轉撥入-轉播出)"
  '        '庫存
  '        Dim dicTmp_CarrierItem = dicCarrierItem.Where(Function(obj)
  '                                                        If obj.Value.SKU_No = SKU_NO AndAlso obj.Value.Lot_No = LOT_NO AndAlso obj.Value.Owner_No = OWNER_NO AndAlso obj.Value.Sub_Owner_No = SUB_OWNER_NO Then
  '                                                          Return True
  '                                                        End If
  '                                                        Return False
  '                                                      End Function).OrderByDescending(Function(obj) obj.Value.Receipt_Date).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
  '        '入庫單
  '        Dim dicTmp_ReceiptPO_DTL = tmp_dicReceiptPO_DTL.Where(Function(q)
  '                                                                If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                  Return True
  '                                                                End If
  '                                                                Return False
  '                                                              End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)

  '        '轉播入
  '        Dim dicTmp_TransformPO_DTL_TRANSACTION = tmp_dicTransformPO_DTL_TRANSACTION.Where(Function(q)
  '                                                                                            If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                                              Return True
  '                                                                                            End If
  '                                                                                            Return False
  '                                                                                          End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)

  '        '轉播出
  '        Dim dicTmp_TransformPO_DTL = tmp_dicTransformPO_DTL.Where(Function(q)
  '                                                                    If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                      Return True
  '                                                                    End If
  '                                                                    Return False
  '                                                                  End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '        '以上三組 才算WMS的庫存 開始計算WMS實際庫存
  '        '庫存
  '        For Each objCarrierItem In dicTmp_CarrierItem.Values
  '          ACC_COMMON3 += 1
  '          If ACC_COMMON1 = "" Then
  '            ACC_COMMON1 = objCarrierItem.Carrier_ID
  '          End If

  '          WMS_STOCK_QTY += objCarrierItem.QTY '計算庫存
  '          '移除已出現的
  '          If dicCarrierItem.ContainsKey(objCarrierItem.gid) Then
  '            dicCarrierItem.Remove(objCarrierItem.gid)
  '          End If
  '        Next
  '        '入庫單
  '        For Each objReceiptPO_DTL In dicTmp_ReceiptPO_DTL.Values
  '          WMS_STOCK_QTY += objReceiptPO_DTL.QTY - objReceiptPO_DTL.QTY_FINISH '計算庫存 這個會部分轉
  '          '移除已出現的
  '          If tmp_dicReceiptPO_DTL.ContainsKey(objReceiptPO_DTL.gid) Then
  '            tmp_dicReceiptPO_DTL.Remove(objReceiptPO_DTL.gid)
  '          End If
  '        Next
  '        '轉播入
  '        For Each objTransformPO_DTL_TRANSACTION In dicTmp_TransformPO_DTL_TRANSACTION.Values
  '          WMS_STOCK_QTY += objTransformPO_DTL_TRANSACTION.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicTransformPO_DTL_TRANSACTION.ContainsKey(objTransformPO_DTL_TRANSACTION.gid) Then
  '            tmp_dicTransformPO_DTL_TRANSACTION.Remove(objTransformPO_DTL_TRANSACTION.gid)
  '          End If
  '        Next
  '        '轉播出
  '        For Each objTransformPO_DTL In dicTmp_TransformPO_DTL.Values
  '          WMS_STOCK_QTY -= objTransformPO_DTL.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicTransformPO_DTL.ContainsKey(objTransformPO_DTL.gid) Then
  '            tmp_dicTransformPO_DTL.Remove(objTransformPO_DTL.gid)
  '          End If
  '        Next
  '#End Region
  '#Region "2.WMS未結案的出庫單 (出庫單+轉播出)"
  '        '出庫單
  '        Dim dicTmp_DischargePO_DTL = tmp_dicDischargePO_DTL.Where(Function(q)
  '                                                                    If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                      Return True
  '                                                                    End If
  '                                                                    Return False
  '                                                                  End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '        '加總未處理數量
  '        '出庫單
  '        For Each objDischargePO_DTL In dicTmp_DischargePO_DTL.Values
  '          WMS_UNFINISH_QTY += objDischargePO_DTL.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicDischargePO_DTL.ContainsKey(objDischargePO_DTL.gid) Then
  '            tmp_dicDischargePO_DTL.Remove(objDischargePO_DTL.gid)
  '          End If
  '        Next
  '#End Region
  '        '實際數量=庫存(上帳數量)-未處理數量(未結數量)
  '        WMS_COMPARSON_QTY = WMS_STOCK_QTY - WMS_UNFINISH_QTY
  '        '差異數量=ERP比對數量-WMS比對數量
  '        QUANTITY_VARIANCE = ERP_COMPARSON_QTY - WMS_COMPARSON_QTY

  '        '最後入的托盤號、托盤位置
  '        If ACC_COMMON1 <> "" Then
  '          If dicCarrier.ContainsKey(clsCarrier.Get_Combination_Key(ACC_COMMON1)) Then
  '            ACC_COMMON2 = dicCarrier.Item(clsCarrier.Get_Combination_Key(ACC_COMMON1)).Location_No
  '          End If
  '        End If

  '        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
  '        ret_lstAddInventoryComparison.Add(objInfo)
  '      Next
  '      Dim str_log = ""
  '      '剩餘的入出+轉播單 (WMS有帳ERP沒有的)
  '      '1. 先看庫存
  '      For Each objCarrierItem In dicCarrierItem.Values
  '        Dim SKU_NO = objCarrierItem.SKU_No
  '        Dim LOT_NO = objCarrierItem.Lot_No
  '        Dim ITEM_COMMON1 = ""
  '        Dim ITEM_COMMON2 = ""
  '        Dim ITEM_COMMON3 = ""
  '        Dim ITEM_COMMON4 = ""
  '        Dim ITEM_COMMON5 = ""
  '        Dim ITEM_COMMON6 = ""
  '        Dim ITEM_COMMON7 = ""
  '        Dim ITEM_COMMON8 = ""
  '        Dim ITEM_COMMON9 = ""
  '        Dim ITEM_COMMON10 = ""
  '        Dim SORT_ITEM_COMMON1 = ""
  '        Dim SORT_ITEM_COMMON2 = ""
  '        Dim SORT_ITEM_COMMON3 = ""
  '        Dim SORT_ITEM_COMMON4 = ""
  '        Dim SORT_ITEM_COMMON5 = ""
  '        Dim OWNER_NO = objCarrierItem.Owner_No
  '        Dim SUB_OWNER_NO = objCarrierItem.Sub_Owner_No
  '        Dim STORAGE_TYPE = 0
  '        Dim BND = 0
  '        Dim QC_STATUS = 0
  '        Dim WMS_STOCK_QTY = 0
  '        Dim WMS_UNFINISH_QTY = 0
  '        Dim WMS_COMPARSON_QTY = 0
  '        Dim ERP_STOCK_QTY = 0 ' objOutput.LABST
  '        Dim ERP_UNFINISH_QTY = 0 ' objOutput.MENGE
  '        Dim ERP_COMPARSON_QTY = 0 ' objOutput.CLABS
  '        Dim QUANTITY_VARIANCE = 0
  '        Dim ERP_SYSTEM = "ERP"
  '        Dim CREATE_TIME = ""
  '        Dim ACC_COMMON1 = "" '最後收料的棧板
  '        Dim ACC_COMMON2 = "" '該棧板的位置
  '        Dim ACC_COMMON3 = 0 '總托盤數
  '        Dim ACC_COMMON4 = ""
  '        Dim ACC_COMMON5 = ""
  '        Dim ACC_COMMON6 = ""
  '        Dim ACC_COMMON7 = ""
  '        Dim ACC_COMMON8 = ""
  '        Dim ACC_COMMON9 = ""
  '        Dim ACC_COMMON10 = ""
  '        'SendMessageToLog("1. 剩餘的WMS庫存. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        'str_log += "1. 剩餘的WMS庫存. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & vbCrLf

  '#Region "1.WMS全部庫存 (入庫單+現有庫存+轉撥入-轉播出)"
  '        '入庫單
  '        Dim dicTmp_ReceiptPO_DTL = tmp_dicReceiptPO_DTL.Where(Function(q)
  '                                                                If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                  Return True
  '                                                                End If
  '                                                                Return False
  '                                                              End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)

  '        '轉播入
  '        Dim dicTmp_TransformPO_DTL_TRANSACTION = tmp_dicTransformPO_DTL_TRANSACTION.Where(Function(q)
  '                                                                                            If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                                              Return True
  '                                                                                            End If
  '                                                                                            Return False
  '                                                                                          End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)

  '        '轉播出
  '        Dim dicTmp_TransformPO_DTL = tmp_dicTransformPO_DTL.Where(Function(q)
  '                                                                    If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                      Return True
  '                                                                    End If
  '                                                                    Return False
  '                                                                  End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '        '以上三組 才算WMS的庫存 開始計算WMS實際庫存
  '        '庫存
  '        WMS_STOCK_QTY += objCarrierItem.QTY
  '        ACC_COMMON3 += 1
  '        If ACC_COMMON1 <> "" Then
  '          ACC_COMMON1 = objCarrierItem.Carrier_ID
  '        End If
  '        '入庫單
  '        For Each objReceiptPO_DTL In dicTmp_ReceiptPO_DTL.Values
  '          WMS_STOCK_QTY += objReceiptPO_DTL.QTY - objReceiptPO_DTL.QTY_FINISH '計算庫存 這個會部分轉
  '          '移除已出現的
  '          If tmp_dicReceiptPO_DTL.ContainsKey(objReceiptPO_DTL.gid) Then
  '            tmp_dicReceiptPO_DTL.Remove(objReceiptPO_DTL.gid)
  '          End If
  '        Next
  '        '轉播入
  '        For Each objTransformPO_DTL_TRANSACTION In dicTmp_TransformPO_DTL_TRANSACTION.Values
  '          WMS_STOCK_QTY += objTransformPO_DTL_TRANSACTION.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicTransformPO_DTL_TRANSACTION.ContainsKey(objTransformPO_DTL_TRANSACTION.gid) Then
  '            tmp_dicTransformPO_DTL_TRANSACTION.Remove(objTransformPO_DTL_TRANSACTION.gid)
  '          End If
  '        Next
  '        '轉播出
  '        For Each objTransformPO_DTL In dicTmp_TransformPO_DTL.Values
  '          WMS_STOCK_QTY -= objTransformPO_DTL.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicTransformPO_DTL.ContainsKey(objTransformPO_DTL.gid) Then
  '            tmp_dicTransformPO_DTL.Remove(objTransformPO_DTL.gid)
  '          End If
  '        Next
  '#End Region
  '#Region "2.WMS未結案的出庫單 (出庫單+轉播出)"
  '        '出庫單
  '        Dim dicTmp_DischargePO_DTL = tmp_dicDischargePO_DTL.Where(Function(q)
  '                                                                    If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                      Return True
  '                                                                    End If
  '                                                                    Return False
  '                                                                  End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '        '加總未處理數量
  '        '出庫單
  '        For Each objDischargePO_DTL In dicTmp_DischargePO_DTL.Values
  '          WMS_UNFINISH_QTY += objDischargePO_DTL.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicDischargePO_DTL.ContainsKey(objDischargePO_DTL.gid) Then
  '            tmp_dicDischargePO_DTL.Remove(objDischargePO_DTL.gid)
  '          End If
  '        Next
  '#End Region
  '        '實際數量=庫存(上帳數量)-未處理數量(未結數量)
  '        WMS_COMPARSON_QTY = WMS_STOCK_QTY - WMS_UNFINISH_QTY
  '        '差異數量=ERP比對數量-WMS比對數量
  '        QUANTITY_VARIANCE = ERP_COMPARSON_QTY - WMS_COMPARSON_QTY

  '        '最後入的托盤號、托盤位置
  '        If ACC_COMMON1 <> "" Then
  '          If dicCarrier.ContainsKey(clsCarrier.Get_Combination_Key(ACC_COMMON1)) Then
  '            ACC_COMMON2 = dicCarrier.Item(clsCarrier.Get_Combination_Key(ACC_COMMON1)).Location_No
  '          End If
  '        End If

  '        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
  '        ret_lstAddInventoryComparison.Add(objInfo)

  '      Next

  '      '2. 入庫單
  '      For Each objReceiptPO_DTL In tmp_dicReceiptPO_DTL.Values
  '        Dim SKU_NO = objReceiptPO_DTL.SKU_NO
  '        Dim LOT_NO = objReceiptPO_DTL.LOT_NO
  '        Dim ITEM_COMMON1 = ""
  '        Dim ITEM_COMMON2 = ""
  '        Dim ITEM_COMMON3 = ""
  '        Dim ITEM_COMMON4 = ""
  '        Dim ITEM_COMMON5 = ""
  '        Dim ITEM_COMMON6 = ""
  '        Dim ITEM_COMMON7 = ""
  '        Dim ITEM_COMMON8 = ""
  '        Dim ITEM_COMMON9 = ""
  '        Dim ITEM_COMMON10 = ""
  '        Dim SORT_ITEM_COMMON1 = ""
  '        Dim SORT_ITEM_COMMON2 = ""
  '        Dim SORT_ITEM_COMMON3 = ""
  '        Dim SORT_ITEM_COMMON4 = ""
  '        Dim SORT_ITEM_COMMON5 = ""
  '        Dim OWNER_NO = objReceiptPO_DTL.TO_OWNER_ID
  '        Dim SUB_OWNER_NO = objReceiptPO_DTL.TO_SUB_OWNER_ID
  '        Dim STORAGE_TYPE = 0
  '        Dim BND = 0
  '        Dim QC_STATUS = 0
  '        Dim WMS_STOCK_QTY = 0
  '        Dim WMS_UNFINISH_QTY = 0
  '        Dim WMS_COMPARSON_QTY = 0
  '        Dim ERP_STOCK_QTY = 0 'objOutput.LABST
  '        Dim ERP_UNFINISH_QTY = 0 'objOutput.MENGE
  '        Dim ERP_COMPARSON_QTY = 0
  '        Dim QUANTITY_VARIANCE = 0
  '        Dim ERP_SYSTEM = "ERP"
  '        Dim CREATE_TIME = ""
  '        Dim ACC_COMMON1 = "" '最後收料的棧板
  '        Dim ACC_COMMON2 = "" '該棧板的位置
  '        Dim ACC_COMMON3 = 0 '總托盤數
  '        Dim ACC_COMMON4 = ""
  '        Dim ACC_COMMON5 = ""
  '        Dim ACC_COMMON6 = ""
  '        Dim ACC_COMMON7 = ""
  '        Dim ACC_COMMON8 = ""
  '        Dim ACC_COMMON9 = ""
  '        Dim ACC_COMMON10 = ""

  '        'SendMessageToLog("2. 剩餘的入庫單. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & " ,POID=" & objReceiptPO_DTL.PO_ID & " ,SerialNo=" & objReceiptPO_DTL.PO_SERIAL_NO, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        'str_log += "2. 剩餘的入庫單. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & " ,POID=" & objReceiptPO_DTL.PO_ID & " ,SerialNo=" & objReceiptPO_DTL.PO_SERIAL_NO & vbCrLf


  '#Region "1.WMS全部庫存 (入庫單+現有庫存+轉撥入-轉播出)"
  '        '轉播入
  '        Dim dicTmp_TransformPO_DTL_TRANSACTION = tmp_dicTransformPO_DTL_TRANSACTION.Where(Function(q)
  '                                                                                            If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                                              Return True
  '                                                                                            End If
  '                                                                                            Return False
  '                                                                                          End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)

  '        '轉播出
  '        Dim dicTmp_TransformPO_DTL = tmp_dicTransformPO_DTL.Where(Function(q)
  '                                                                    If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                      Return True
  '                                                                    End If
  '                                                                    Return False
  '                                                                  End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '        '以上三組 才算WMS的庫存 開始計算WMS實際庫存
  '        WMS_STOCK_QTY += objReceiptPO_DTL.QTY - objReceiptPO_DTL.QTY_FINISH
  '        '轉播入
  '        For Each objTransformPO_DTL_TRANSACTION In dicTmp_TransformPO_DTL_TRANSACTION.Values
  '          WMS_STOCK_QTY += objTransformPO_DTL_TRANSACTION.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicTransformPO_DTL_TRANSACTION.ContainsKey(objTransformPO_DTL_TRANSACTION.gid) Then
  '            tmp_dicTransformPO_DTL_TRANSACTION.Remove(objTransformPO_DTL_TRANSACTION.gid)
  '          End If
  '        Next
  '        '轉播出
  '        For Each objTransformPO_DTL In dicTmp_TransformPO_DTL.Values
  '          WMS_STOCK_QTY -= objTransformPO_DTL.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicTransformPO_DTL.ContainsKey(objTransformPO_DTL.gid) Then
  '            tmp_dicTransformPO_DTL.Remove(objTransformPO_DTL.gid)
  '          End If
  '        Next
  '#End Region
  '#Region "2.WMS未結案的出庫單 (出庫單+轉播出)"
  '        '出庫單
  '        Dim dicTmp_DischargePO_DTL = tmp_dicDischargePO_DTL.Where(Function(q)
  '                                                                    If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                      Return True
  '                                                                    End If
  '                                                                    Return False
  '                                                                  End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '        '加總未處理數量
  '        '出庫單
  '        For Each objDischargePO_DTL In dicTmp_DischargePO_DTL.Values
  '          WMS_UNFINISH_QTY += objDischargePO_DTL.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicDischargePO_DTL.ContainsKey(objDischargePO_DTL.gid) Then
  '            tmp_dicDischargePO_DTL.Remove(objDischargePO_DTL.gid)
  '          End If
  '        Next
  '#End Region
  '        '實際數量=庫存(上帳數量)-未處理數量(未結數量)
  '        WMS_COMPARSON_QTY = WMS_STOCK_QTY - WMS_UNFINISH_QTY
  '        '差異數量=ERP比對數量-WMS比對數量
  '        QUANTITY_VARIANCE = ERP_COMPARSON_QTY - WMS_COMPARSON_QTY

  '        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
  '        ret_lstAddInventoryComparison.Add(objInfo)
  '      Next

  '      '3. 轉播入
  '      For Each objTransformPO_DTL_TRANSACTION In tmp_dicTransformPO_DTL_TRANSACTION.Values
  '        Dim SKU_NO = objTransformPO_DTL_TRANSACTION.SKU_NO
  '        Dim LOT_NO = objTransformPO_DTL_TRANSACTION.LOT_NO
  '        Dim ITEM_COMMON1 = ""
  '        Dim ITEM_COMMON2 = ""
  '        Dim ITEM_COMMON3 = ""
  '        Dim ITEM_COMMON4 = ""
  '        Dim ITEM_COMMON5 = ""
  '        Dim ITEM_COMMON6 = ""
  '        Dim ITEM_COMMON7 = ""
  '        Dim ITEM_COMMON8 = ""
  '        Dim ITEM_COMMON9 = ""
  '        Dim ITEM_COMMON10 = ""
  '        Dim SORT_ITEM_COMMON1 = ""
  '        Dim SORT_ITEM_COMMON2 = ""
  '        Dim SORT_ITEM_COMMON3 = ""
  '        Dim SORT_ITEM_COMMON4 = ""
  '        Dim SORT_ITEM_COMMON5 = ""
  '        Dim OWNER_NO = objTransformPO_DTL_TRANSACTION.TO_OWNER_ID
  '        Dim SUB_OWNER_NO = objTransformPO_DTL_TRANSACTION.TO_SUB_OWNER_ID
  '        Dim STORAGE_TYPE = 0
  '        Dim BND = 0
  '        Dim QC_STATUS = 0
  '        Dim WMS_STOCK_QTY = 0
  '        Dim WMS_UNFINISH_QTY = 0
  '        Dim WMS_COMPARSON_QTY = 0
  '        Dim ERP_STOCK_QTY = 0 'objOutput.LABST
  '        Dim ERP_UNFINISH_QTY = 0 'objOutput.MENGE
  '        Dim ERP_COMPARSON_QTY = 0 'objOutput.CLABS
  '        Dim QUANTITY_VARIANCE = 0
  '        Dim ERP_SYSTEM = "ERP"
  '        Dim CREATE_TIME = ""
  '        Dim ACC_COMMON1 = "" '最後收料的棧板
  '        Dim ACC_COMMON2 = "" '該棧板的位置
  '        Dim ACC_COMMON3 = 0 '總托盤數
  '        Dim ACC_COMMON4 = ""
  '        Dim ACC_COMMON5 = ""
  '        Dim ACC_COMMON6 = ""
  '        Dim ACC_COMMON7 = ""
  '        Dim ACC_COMMON8 = ""
  '        Dim ACC_COMMON9 = ""
  '        Dim ACC_COMMON10 = ""

  '        'SendMessageToLog("5. 剩餘的轉播入. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & " ,POID=" & objTransformPO_DTL_TRANSACTION.PO_ID & " ,SerialNo=" & objTransformPO_DTL_TRANSACTION.PO_SERIAL_NO, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        'str_log += "5. 剩餘的轉播入. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & " ,POID=" & objTransformPO_DTL_TRANSACTION.PO_ID & " ,SerialNo=" & objTransformPO_DTL_TRANSACTION.PO_SERIAL_NO & vbCrLf



  '#Region "1.WMS全部庫存 (入庫單+現有庫存+轉撥入-轉播出)"
  '        '轉播出
  '        Dim dicTmp_TransformPO_DTL = tmp_dicTransformPO_DTL.Where(Function(q)
  '                                                                    If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                      Return True
  '                                                                    End If
  '                                                                    Return False
  '                                                                  End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '        '以上三組 才算WMS的庫存 開始計算WMS實際庫存
  '        WMS_STOCK_QTY += objTransformPO_DTL_TRANSACTION.QTY '計算庫存
  '        '轉播出
  '        For Each objTransformPO_DTL In dicTmp_TransformPO_DTL.Values
  '          WMS_STOCK_QTY -= objTransformPO_DTL.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicTransformPO_DTL.ContainsKey(objTransformPO_DTL.gid) Then
  '            tmp_dicTransformPO_DTL.Remove(objTransformPO_DTL.gid)
  '          End If
  '        Next
  '#End Region
  '#Region "2.WMS未結案的出庫單 (出庫單+轉播出)"
  '        '出庫單
  '        Dim dicTmp_DischargePO_DTL = tmp_dicDischargePO_DTL.Where(Function(q)
  '                                                                    If q.Value.SKU_NO = SKU_NO AndAlso q.Value.LOT_NO = LOT_NO AndAlso q.Value.TO_OWNER_ID = OWNER_NO AndAlso q.Value.TO_SUB_OWNER_ID = SUB_OWNER_NO Then
  '                                                                      Return True
  '                                                                    End If
  '                                                                    Return False
  '                                                                  End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
  '        '加總未處理數量
  '        '出庫單
  '        For Each objDischargePO_DTL In dicTmp_DischargePO_DTL.Values
  '          WMS_UNFINISH_QTY += objDischargePO_DTL.QTY '計算庫存
  '          '移除已出現的
  '          If tmp_dicDischargePO_DTL.ContainsKey(objDischargePO_DTL.gid) Then
  '            tmp_dicDischargePO_DTL.Remove(objDischargePO_DTL.gid)
  '          End If
  '        Next
  '#End Region
  '        '實際數量=庫存(上帳數量)-未處理數量(未結數量)
  '        WMS_COMPARSON_QTY = WMS_STOCK_QTY - WMS_UNFINISH_QTY
  '        '差異數量=ERP比對數量-WMS比對數量
  '        QUANTITY_VARIANCE = ERP_COMPARSON_QTY - WMS_COMPARSON_QTY

  '        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
  '        ret_lstAddInventoryComparison.Add(objInfo)
  '      Next

  '      '4. 轉播出
  '      For Each objTransformPO_DTL In tmp_dicTransformPO_DTL.Values
  '        Dim SKU_NO = objTransformPO_DTL.SKU_NO
  '        Dim LOT_NO = objTransformPO_DTL.LOT_NO
  '        Dim ITEM_COMMON1 = ""
  '        Dim ITEM_COMMON2 = ""
  '        Dim ITEM_COMMON3 = ""
  '        Dim ITEM_COMMON4 = ""
  '        Dim ITEM_COMMON5 = ""
  '        Dim ITEM_COMMON6 = ""
  '        Dim ITEM_COMMON7 = ""
  '        Dim ITEM_COMMON8 = ""
  '        Dim ITEM_COMMON9 = ""
  '        Dim ITEM_COMMON10 = ""
  '        Dim SORT_ITEM_COMMON1 = ""
  '        Dim SORT_ITEM_COMMON2 = ""
  '        Dim SORT_ITEM_COMMON3 = ""
  '        Dim SORT_ITEM_COMMON4 = ""
  '        Dim SORT_ITEM_COMMON5 = ""
  '        Dim OWNER_NO = objTransformPO_DTL.FROM_OWNER_ID
  '        Dim SUB_OWNER_NO = objTransformPO_DTL.FROM_SUB_OWNER_ID
  '        Dim STORAGE_TYPE = 0
  '        Dim BND = 0
  '        Dim QC_STATUS = 0
  '        Dim WMS_STOCK_QTY = 0
  '        Dim WMS_UNFINISH_QTY = objTransformPO_DTL.QTY
  '        Dim WMS_COMPARSON_QTY = 0
  '        Dim ERP_STOCK_QTY = 0 'objOutput.LABST
  '        Dim ERP_UNFINISH_QTY = 0 'objOutput.MENGE
  '        Dim ERP_COMPARSON_QTY = 0
  '        Dim QUANTITY_VARIANCE = 0
  '        Dim ERP_SYSTEM = "ERP"
  '        Dim CREATE_TIME = ""
  '        Dim ACC_COMMON1 = "" '最後收料的棧板
  '        Dim ACC_COMMON2 = "" '該棧板的位置
  '        Dim ACC_COMMON3 = 0 '總托盤數
  '        Dim ACC_COMMON4 = ""
  '        Dim ACC_COMMON5 = ""
  '        Dim ACC_COMMON6 = ""
  '        Dim ACC_COMMON7 = ""
  '        Dim ACC_COMMON8 = ""
  '        Dim ACC_COMMON9 = ""
  '        Dim ACC_COMMON10 = ""
  '        'SendMessageToLog("4. 剩餘的轉播出. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & " ,POID=" & objTransformPO_DTL.PO_ID & " ,SerialNo=" & objTransformPO_DTL.PO_SERIAL_NO, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        'str_log += "4. 剩餘的轉播出. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & " ,POID=" & objTransformPO_DTL.PO_ID & " ,SerialNo=" & objTransformPO_DTL.PO_SERIAL_NO & vbCrLf

  '        '實際數量=庫存(上帳數量)-未處理數量(未結數量)
  '        WMS_COMPARSON_QTY = WMS_STOCK_QTY - WMS_UNFINISH_QTY
  '        '差異數量=ERP比對數量-WMS比對數量
  '        QUANTITY_VARIANCE = ERP_COMPARSON_QTY - WMS_COMPARSON_QTY

  '        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
  '        ret_lstAddInventoryComparison.Add(objInfo)
  '      Next

  '      '5 出庫單
  '      For Each objDischargePO_DTL In tmp_dicDischargePO_DTL.Values
  '        Dim SKU_NO = objDischargePO_DTL.SKU_NO
  '        Dim LOT_NO = objDischargePO_DTL.LOT_NO
  '        Dim ITEM_COMMON1 = ""
  '        Dim ITEM_COMMON2 = ""
  '        Dim ITEM_COMMON3 = ""
  '        Dim ITEM_COMMON4 = ""
  '        Dim ITEM_COMMON5 = ""
  '        Dim ITEM_COMMON6 = ""
  '        Dim ITEM_COMMON7 = ""
  '        Dim ITEM_COMMON8 = ""
  '        Dim ITEM_COMMON9 = ""
  '        Dim ITEM_COMMON10 = ""
  '        Dim SORT_ITEM_COMMON1 = ""
  '        Dim SORT_ITEM_COMMON2 = ""
  '        Dim SORT_ITEM_COMMON3 = ""
  '        Dim SORT_ITEM_COMMON4 = ""
  '        Dim SORT_ITEM_COMMON5 = ""
  '        Dim OWNER_NO = objDischargePO_DTL.FROM_OWNER_ID
  '        Dim SUB_OWNER_NO = objDischargePO_DTL.FROM_SUB_OWNER_ID
  '        Dim STORAGE_TYPE = 0
  '        Dim BND = 0
  '        Dim QC_STATUS = 0
  '        Dim WMS_STOCK_QTY = 0
  '        Dim WMS_UNFINISH_QTY = objDischargePO_DTL.QTY
  '        Dim WMS_COMPARSON_QTY = 0
  '        Dim ERP_STOCK_QTY = 0 'objOutput.LABST
  '        Dim ERP_UNFINISH_QTY = 0 'objOutput.MENGE
  '        Dim ERP_COMPARSON_QTY = 0
  '        Dim QUANTITY_VARIANCE = 0
  '        Dim ERP_SYSTEM = "ERP"
  '        Dim CREATE_TIME = ""
  '        Dim ACC_COMMON1 = "" '最後收料的棧板
  '        Dim ACC_COMMON2 = "" '該棧板的位置
  '        Dim ACC_COMMON3 = 0 '總托盤數
  '        Dim ACC_COMMON4 = ""
  '        Dim ACC_COMMON5 = ""
  '        Dim ACC_COMMON6 = ""
  '        Dim ACC_COMMON7 = ""
  '        Dim ACC_COMMON8 = ""
  '        Dim ACC_COMMON9 = ""
  '        Dim ACC_COMMON10 = ""
  '        'SendMessageToLog("5. 剩餘的出庫單. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & " ,POID=" & objDischargePO_DTL.PO_ID & " ,SerialNo=" & objDischargePO_DTL.PO_SERIAL_NO, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        str_log += "5. 剩餘的出庫單. SKU=" & SKU_NO & " ,Lot=" & LOT_NO & " ,Owner=" & OWNER_NO & " ,SubOwner=" & SUB_OWNER_NO & " ,POID=" & objDischargePO_DTL.PO_ID & " ,SerialNo=" & objDischargePO_DTL.PO_SERIAL_NO & vbCrLf


  '        '實際數量=庫存(上帳數量)-未處理數量(未結數量)
  '        WMS_COMPARSON_QTY = WMS_STOCK_QTY - WMS_UNFINISH_QTY
  '        '差異數量=ERP比對數量-WMS比對數量
  '        QUANTITY_VARIANCE = ERP_COMPARSON_QTY - WMS_COMPARSON_QTY

  '        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
  '        ret_lstAddInventoryComparison.Add(objInfo)
  '      Next
  '      'If str_log <> "" Then
  '      '  SendMessageToLog(str_log, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
  '      'End If
  '      SendMessageToLog("比對完成", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

  '      Return True
  '    Catch ex As Exception
  '      ret_strResultMsg = ex.InnerException.Message
  '      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return False
  '    End Try
  '  End Function
#End Region
#Region "新方法"

  '資料處理
  Private Function Process_Data(ByVal Receive_Msg As MSG_T11F2U1_InventoryComparison,
                                ByVal ret_dicAddInventoryComparison As Dictionary(Of String, clsWMS_CT_INVENTORY_COMPARISON),
                                ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim S_DATE As String = GetNewTime_ByDataTimeFormat("yyyyMMdd")
      Dim E_DATE As String = GetNewTime_ByDataTimeFormat("yyyyMMdd")
      Dim objERP_Report As New MSG_ASRS_getCompStock

      '取出全部庫存
      Dim tmp_dicTransformPOID As New Dictionary(Of String, String)
      Dim tmp_dicTransformPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicTransformPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicTransformPO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim tmp_dicTransformPO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)

      Dim tmp_dicDischargePOID As New Dictionary(Of String, String)
      Dim tmp_dicDischargePO As New Dictionary(Of String, clsPO)
      Dim tmp_dicDischargePO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicDischargePO_DTL As New Dictionary(Of String, clsPO_DTL)

      Dim tmp_dicReceiptPOID As New Dictionary(Of String, String)
      Dim tmp_dicReceiptPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicReceiptPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicReceiptPO_DTL As New Dictionary(Of String, clsPO_DTL)
#Region "取得未結案的入庫單"
      '取出未結案的單據，若為入庫單則排除庫存
      Dim dicPO As New Dictionary(Of String, clsPO)
      gMain.objHandling.O_GetDB_dicPOByALL(dicPO)
      tmp_dicReceiptPO = dicPO.Where(Function(q)
                                       If q.Value.WO_Type = enuWOType.Receipt Then Return True
                                       Return False
                                     End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
      For Each objPO In tmp_dicReceiptPO.Values
        If tmp_dicReceiptPOID.ContainsKey(objPO.PO_ID) = False Then
          tmp_dicReceiptPOID.Add(objPO.PO_ID, objPO.PO_ID)
        End If
      Next
      '使用dicPO取得資料庫裡的PO_Line資料
      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicReceiptPOID, tmp_dicReceiptPO_Line) = False Then
        ret_strResultMsg = "WMS get PO_Line data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '使用dicPO取得資料庫裡的PO_DTL資料
      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicReceiptPOID, tmp_dicReceiptPO_DTL) = False Then
        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
#End Region
#Region "取得未結案的出庫單"
      '取出未結案的單據，若為出庫單則排除庫存
      'dicPO = New Dictionary(Of String, clsPO)
      'gMain.objHandling.O_GetDB_dicPOByALL(dicPO)
      tmp_dicDischargePO = dicPO.Where(Function(q)
                                         If q.Value.WO_Type = enuWOType.Discharge AndAlso q.Value.PO_Type2 = enuPOType_2.material_out Then Return True
                                         Return False
                                       End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
      For Each objPO In tmp_dicDischargePO.Values
        If tmp_dicDischargePOID.ContainsKey(objPO.PO_ID) = False Then
          tmp_dicDischargePOID.Add(objPO.PO_ID, objPO.PO_ID)
        End If
      Next
      '使用dicPO取得資料庫裡的PO_Line資料
      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicDischargePOID, tmp_dicDischargePO_Line) = False Then
        ret_strResultMsg = "WMS get PO_Line data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '使用dicPO取得資料庫裡的PO_DTL資料
      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicDischargePOID, tmp_dicDischargePO_DTL) = False Then
        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
#End Region
#Region "取得未結案的轉播單"
      '取出未結案的單據，若為出庫單則排除庫存
      tmp_dicTransformPO = dicPO.Where(Function(q)
                                         If q.Value.WO_Type = enuWOType.Transform AndAlso q.Value.PO_Type2 = enuPOType_2.transaction_account Then Return True
                                         Return False
                                       End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
      For Each objPO In tmp_dicTransformPO.Values
        If tmp_dicTransformPOID.ContainsKey(objPO.PO_ID) = False Then
          tmp_dicTransformPOID.Add(objPO.PO_ID, objPO.PO_ID)
        End If
      Next
      '使用dicPO取得資料庫裡的PO_Line資料
      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicTransformPOID, tmp_dicTransformPO_Line) = False Then
        ret_strResultMsg = "WMS get PO_Line data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '使用dicPO取得資料庫裡的PO_DTL資料
      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicTransformPOID, tmp_dicTransformPO_DTL) = False Then
        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '使用dicPO取得資料庫裡的PO_DTL資料
      If gMain.objHandling.O_GetDB_dicPODTLTRANSACTIONBydicPO_ID(tmp_dicTransformPOID, tmp_dicTransformPO_DTL_TRANSACTION) = False Then
        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
#End Region
#Region "取得全部庫存"
      Dim dicCarrierItem = gMain.objHandling.GetCarrierItemByACCEPTING_STATUS(enuAcceptingStatus.Inventory)
      Dim dicCarrier = gMain.objHandling.GetCarrierStatusByAll
#End Region



      If objERP_Report Is Nothing Then
        ret_strResultMsg = "ERP回傳結果序列化後為空"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '開始比對

      SendMessageToLog("開始比對", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      '1. ERP庫存
      SendMessageToLog("ERP_Report Start", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Each objOutput In objERP_Report.output
        Dim SKU_NO = objOutput.MATNR
        Dim LOT_NO = objOutput.CHARG
        Dim ITEM_COMMON1 = ""
        Dim ITEM_COMMON2 = ""
        Dim ITEM_COMMON3 = ""
        Dim ITEM_COMMON4 = ""
        Dim ITEM_COMMON5 = ""
        Dim ITEM_COMMON6 = ""
        Dim ITEM_COMMON7 = ""
        Dim ITEM_COMMON8 = ""
        Dim ITEM_COMMON9 = ""
        Dim ITEM_COMMON10 = ""
        Dim SORT_ITEM_COMMON1 = ""
        Dim SORT_ITEM_COMMON2 = ""
        Dim SORT_ITEM_COMMON3 = ""
        Dim SORT_ITEM_COMMON4 = ""
        Dim SORT_ITEM_COMMON5 = ""
        Dim OWNER_NO = objOutput.WERKS
        Dim SUB_OWNER_NO = objOutput.LGORT
        Dim STORAGE_TYPE = 0
        Dim BND = 0
        Dim QC_STATUS = 0
        Dim WMS_STOCK_QTY = 0
        Dim WMS_UNFINISH_QTY = 0
        Dim WMS_COMPARSON_QTY = 0
        Dim ERP_STOCK_QTY = IIf(objOutput.LABST = "", 0, objOutput.LABST)
        Dim ERP_UNFINISH_QTY = IIf(objOutput.MENGE = "", 0, objOutput.MENGE)
        Dim ERP_COMPARSON_QTY = IIf(objOutput.CLABS = "", 0, objOutput.CLABS)
        Dim QUANTITY_VARIANCE = 0
        Dim ERP_SYSTEM = "ERP"
        Dim CREATE_TIME = ""
        Dim ACC_COMMON1 = "" '最後收料的棧板
        Dim ACC_COMMON2 = "" '該棧板的位置
        Dim ACC_COMMON3 = 0 '總托盤數
        Dim ACC_COMMON4 = ""
        Dim ACC_COMMON5 = ""
        Dim ACC_COMMON6 = ""
        Dim ACC_COMMON7 = ""
        Dim ACC_COMMON8 = ""
        Dim ACC_COMMON9 = ""
        Dim ACC_COMMON10 = ""

        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
        Dim objTmp As clsWMS_CT_INVENTORY_COMPARISON = Nothing
        If ret_dicAddInventoryComparison.TryGetValue(objInfo.gid, objTmp) = False Then
          ret_dicAddInventoryComparison.Add(objInfo.gid, objInfo)
        Else
          objTmp.ERP_STOCK_QTY += IIf(objOutput.LABST = "", 0, objOutput.LABST)
          objTmp.ERP_UNFINISH_QTY += IIf(objOutput.MENGE = "", 0, objOutput.MENGE)
          objTmp.ERP_COMPARSON_QTY += IIf(objOutput.CLABS = "", 0, objOutput.CLABS)
        End If
      Next
      SendMessageToLog("ERP_Report Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      '2. WMS庫存
      SendMessageToLog("CarrierItem Start", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Each objCarrierItem In dicCarrierItem.Values
        Dim SKU_NO = objCarrierItem.SKU_No
        Dim LOT_NO = objCarrierItem.Lot_No
        Dim ITEM_COMMON1 = ""
        Dim ITEM_COMMON2 = ""
        Dim ITEM_COMMON3 = ""
        Dim ITEM_COMMON4 = ""
        Dim ITEM_COMMON5 = ""
        Dim ITEM_COMMON6 = ""
        Dim ITEM_COMMON7 = ""
        Dim ITEM_COMMON8 = ""
        Dim ITEM_COMMON9 = ""
        Dim ITEM_COMMON10 = ""
        Dim SORT_ITEM_COMMON1 = ""
        Dim SORT_ITEM_COMMON2 = ""
        Dim SORT_ITEM_COMMON3 = ""
        Dim SORT_ITEM_COMMON4 = ""
        Dim SORT_ITEM_COMMON5 = ""
        Dim OWNER_NO = objCarrierItem.Owner_No
        Dim SUB_OWNER_NO = objCarrierItem.Sub_Owner_No
        Dim STORAGE_TYPE = 0
        Dim BND = 0
        Dim QC_STATUS = 0
        Dim WMS_STOCK_QTY = objCarrierItem.QTY
        Dim WMS_UNFINISH_QTY = 0
        Dim WMS_COMPARSON_QTY = 0
        Dim ERP_STOCK_QTY = 0 ' objOutput.LABST
        Dim ERP_UNFINISH_QTY = 0 ' objOutput.MENGE
        Dim ERP_COMPARSON_QTY = 0 ' objOutput.CLABS
        Dim QUANTITY_VARIANCE = 0
        Dim ERP_SYSTEM = "ERP"
        Dim CREATE_TIME = ""
        Dim ACC_COMMON1 = "" '最後收料的棧板
        Dim ACC_COMMON2 = "" '該棧板的位置
        Dim ACC_COMMON3 = 1 '總托盤數
        Dim ACC_COMMON4 = ""
        Dim ACC_COMMON5 = ""
        Dim ACC_COMMON6 = ""
        Dim ACC_COMMON7 = ""
        Dim ACC_COMMON8 = ""
        Dim ACC_COMMON9 = ""
        Dim ACC_COMMON10 = ""

        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
        Dim objTmp As clsWMS_CT_INVENTORY_COMPARISON = Nothing
        If ret_dicAddInventoryComparison.TryGetValue(objInfo.gid, objTmp) = False Then
          ret_dicAddInventoryComparison.Add(objInfo.gid, objInfo)
        Else
          objTmp.WMS_STOCK_QTY += objCarrierItem.QTY
          objTmp.ACC_COMMON3 += ACC_COMMON3
          'objTmp.WMS_UNFINISH_QTY += IIf(objOutput.MENGE = "", 0, objOutput.MENGE)
          'objTmp.WMS_COMPARSON_QTY += IIf(objOutput.CLABS = "", 0, objOutput.CLABS)
        End If
      Next
      SendMessageToLog("CarrierItem Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      '3. 入庫單
      SendMessageToLog("ReceiptPO_DTL Start", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Each objReceiptPO_DTL In tmp_dicReceiptPO_DTL.Values
        Dim SKU_NO = objReceiptPO_DTL.SKU_NO
        Dim LOT_NO = objReceiptPO_DTL.LOT_NO
        Dim ITEM_COMMON1 = ""
        Dim ITEM_COMMON2 = ""
        Dim ITEM_COMMON3 = ""
        Dim ITEM_COMMON4 = ""
        Dim ITEM_COMMON5 = ""
        Dim ITEM_COMMON6 = ""
        Dim ITEM_COMMON7 = ""
        Dim ITEM_COMMON8 = ""
        Dim ITEM_COMMON9 = ""
        Dim ITEM_COMMON10 = ""
        Dim SORT_ITEM_COMMON1 = ""
        Dim SORT_ITEM_COMMON2 = ""
        Dim SORT_ITEM_COMMON3 = ""
        Dim SORT_ITEM_COMMON4 = ""
        Dim SORT_ITEM_COMMON5 = ""
        Dim OWNER_NO = objReceiptPO_DTL.TO_OWNER_ID
        Dim SUB_OWNER_NO = objReceiptPO_DTL.TO_SUB_OWNER_ID
        Dim STORAGE_TYPE = 0
        Dim BND = 0
        Dim QC_STATUS = 0
        Dim WMS_STOCK_QTY = objReceiptPO_DTL.QTY - objReceiptPO_DTL.QTY_FINISH
        Dim WMS_UNFINISH_QTY = 0
        Dim WMS_COMPARSON_QTY = 0
        Dim ERP_STOCK_QTY = 0 'objOutput.LABST
        Dim ERP_UNFINISH_QTY = 0 'objOutput.MENGE
        Dim ERP_COMPARSON_QTY = 0
        Dim QUANTITY_VARIANCE = 0
        Dim ERP_SYSTEM = "ERP"
        Dim CREATE_TIME = ""
        Dim ACC_COMMON1 = "" '最後收料的棧板
        Dim ACC_COMMON2 = "" '該棧板的位置
        Dim ACC_COMMON3 = 0 '總托盤數
        Dim ACC_COMMON4 = ""
        Dim ACC_COMMON5 = ""
        Dim ACC_COMMON6 = ""
        Dim ACC_COMMON7 = ""
        Dim ACC_COMMON8 = ""
        Dim ACC_COMMON9 = ""
        Dim ACC_COMMON10 = ""
        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
        Dim objTmp As clsWMS_CT_INVENTORY_COMPARISON = Nothing
        If ret_dicAddInventoryComparison.TryGetValue(objInfo.gid, objTmp) = False Then
          ret_dicAddInventoryComparison.Add(objInfo.gid, objInfo)
        Else
          objTmp.WMS_STOCK_QTY += WMS_STOCK_QTY
        End If
      Next
      SendMessageToLog("ReceiptPO_DTL Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      '4. 轉播入
      SendMessageToLog("TransformPO_DTL_TRANSACTION Start", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Each objTransformPO_DTL_TRANSACTION In tmp_dicTransformPO_DTL_TRANSACTION.Values
        Dim SKU_NO = objTransformPO_DTL_TRANSACTION.SKU_NO
        Dim LOT_NO = objTransformPO_DTL_TRANSACTION.LOT_NO
        Dim ITEM_COMMON1 = ""
        Dim ITEM_COMMON2 = ""
        Dim ITEM_COMMON3 = ""
        Dim ITEM_COMMON4 = ""
        Dim ITEM_COMMON5 = ""
        Dim ITEM_COMMON6 = ""
        Dim ITEM_COMMON7 = ""
        Dim ITEM_COMMON8 = ""
        Dim ITEM_COMMON9 = ""
        Dim ITEM_COMMON10 = ""
        Dim SORT_ITEM_COMMON1 = ""
        Dim SORT_ITEM_COMMON2 = ""
        Dim SORT_ITEM_COMMON3 = ""
        Dim SORT_ITEM_COMMON4 = ""
        Dim SORT_ITEM_COMMON5 = ""
        Dim OWNER_NO = objTransformPO_DTL_TRANSACTION.TO_OWNER_ID
        Dim SUB_OWNER_NO = objTransformPO_DTL_TRANSACTION.TO_SUB_OWNER_ID
        Dim STORAGE_TYPE = 0
        Dim BND = 0
        Dim QC_STATUS = 0
        Dim WMS_STOCK_QTY = objTransformPO_DTL_TRANSACTION.QTY
        Dim WMS_UNFINISH_QTY = 0
        Dim WMS_COMPARSON_QTY = 0
        Dim ERP_STOCK_QTY = 0 'objOutput.LABST
        Dim ERP_UNFINISH_QTY = 0 'objOutput.MENGE
        Dim ERP_COMPARSON_QTY = 0 'objOutput.CLABS
        Dim QUANTITY_VARIANCE = 0
        Dim ERP_SYSTEM = "ERP"
        Dim CREATE_TIME = ""
        Dim ACC_COMMON1 = "" '最後收料的棧板
        Dim ACC_COMMON2 = "" '該棧板的位置
        Dim ACC_COMMON3 = 0 '總托盤數
        Dim ACC_COMMON4 = ""
        Dim ACC_COMMON5 = ""
        Dim ACC_COMMON6 = ""
        Dim ACC_COMMON7 = ""
        Dim ACC_COMMON8 = ""
        Dim ACC_COMMON9 = ""
        Dim ACC_COMMON10 = ""

        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
        Dim objTmp As clsWMS_CT_INVENTORY_COMPARISON = Nothing
        If ret_dicAddInventoryComparison.TryGetValue(objInfo.gid, objTmp) = False Then
          ret_dicAddInventoryComparison.Add(objInfo.gid, objInfo)
        Else
          objTmp.WMS_STOCK_QTY += WMS_STOCK_QTY
        End If
      Next
      SendMessageToLog("TransformPO_DTL_TRANSACTION Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      '5. 轉播出
      SendMessageToLog("TransformPO_DTL Start", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Each objTransformPO_DTL In tmp_dicTransformPO_DTL.Values
        Dim SKU_NO = objTransformPO_DTL.SKU_NO
        Dim LOT_NO = objTransformPO_DTL.LOT_NO
        Dim ITEM_COMMON1 = ""
        Dim ITEM_COMMON2 = ""
        Dim ITEM_COMMON3 = ""
        Dim ITEM_COMMON4 = ""
        Dim ITEM_COMMON5 = ""
        Dim ITEM_COMMON6 = ""
        Dim ITEM_COMMON7 = ""
        Dim ITEM_COMMON8 = ""
        Dim ITEM_COMMON9 = ""
        Dim ITEM_COMMON10 = ""
        Dim SORT_ITEM_COMMON1 = ""
        Dim SORT_ITEM_COMMON2 = ""
        Dim SORT_ITEM_COMMON3 = ""
        Dim SORT_ITEM_COMMON4 = ""
        Dim SORT_ITEM_COMMON5 = ""
        Dim OWNER_NO = objTransformPO_DTL.FROM_OWNER_ID
        Dim SUB_OWNER_NO = objTransformPO_DTL.FROM_SUB_OWNER_ID
        Dim STORAGE_TYPE = 0
        Dim BND = 0
        Dim QC_STATUS = 0
        Dim WMS_STOCK_QTY = 0 - objTransformPO_DTL.QTY
        Dim WMS_UNFINISH_QTY = objTransformPO_DTL.QTY
        Dim WMS_COMPARSON_QTY = 0
        Dim ERP_STOCK_QTY = 0 'objOutput.LABST
        Dim ERP_UNFINISH_QTY = 0 'objOutput.MENGE
        Dim ERP_COMPARSON_QTY = 0
        Dim QUANTITY_VARIANCE = 0
        Dim ERP_SYSTEM = "ERP"
        Dim CREATE_TIME = ""
        Dim ACC_COMMON1 = "" '最後收料的棧板
        Dim ACC_COMMON2 = "" '該棧板的位置
        Dim ACC_COMMON3 = 0 '總托盤數
        Dim ACC_COMMON4 = ""
        Dim ACC_COMMON5 = ""
        Dim ACC_COMMON6 = ""
        Dim ACC_COMMON7 = ""
        Dim ACC_COMMON8 = ""
        Dim ACC_COMMON9 = ""
        Dim ACC_COMMON10 = ""

        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
        Dim objTmp As clsWMS_CT_INVENTORY_COMPARISON = Nothing
        If ret_dicAddInventoryComparison.TryGetValue(objInfo.gid, objTmp) = False Then
          ret_dicAddInventoryComparison.Add(objInfo.gid, objInfo)
        Else
          objTmp.WMS_STOCK_QTY += WMS_STOCK_QTY
          objTmp.WMS_UNFINISH_QTY += WMS_UNFINISH_QTY
        End If
      Next
      SendMessageToLog("TransformPO_DTL Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      '6 出庫單
      SendMessageToLog("DischargePO_DTL Start", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Each objDischargePO_DTL In tmp_dicDischargePO_DTL.Values
        Dim SKU_NO = objDischargePO_DTL.SKU_NO
        Dim LOT_NO = objDischargePO_DTL.LOT_NO
        Dim ITEM_COMMON1 = ""
        Dim ITEM_COMMON2 = ""
        Dim ITEM_COMMON3 = ""
        Dim ITEM_COMMON4 = ""
        Dim ITEM_COMMON5 = ""
        Dim ITEM_COMMON6 = ""
        Dim ITEM_COMMON7 = ""
        Dim ITEM_COMMON8 = ""
        Dim ITEM_COMMON9 = ""
        Dim ITEM_COMMON10 = ""
        Dim SORT_ITEM_COMMON1 = ""
        Dim SORT_ITEM_COMMON2 = ""
        Dim SORT_ITEM_COMMON3 = ""
        Dim SORT_ITEM_COMMON4 = ""
        Dim SORT_ITEM_COMMON5 = ""
        Dim OWNER_NO = objDischargePO_DTL.FROM_OWNER_ID
        Dim SUB_OWNER_NO = objDischargePO_DTL.FROM_SUB_OWNER_ID
        Dim STORAGE_TYPE = 0
        Dim BND = 0
        Dim QC_STATUS = 0
        Dim WMS_STOCK_QTY = 0
        Dim WMS_UNFINISH_QTY = objDischargePO_DTL.QTY
        Dim WMS_COMPARSON_QTY = 0
        Dim ERP_STOCK_QTY = 0 'objOutput.LABST
        Dim ERP_UNFINISH_QTY = 0 'objOutput.MENGE
        Dim ERP_COMPARSON_QTY = 0
        Dim QUANTITY_VARIANCE = 0
        Dim ERP_SYSTEM = "ERP"
        Dim CREATE_TIME = ""
        Dim ACC_COMMON1 = "" '最後收料的棧板
        Dim ACC_COMMON2 = "" '該棧板的位置
        Dim ACC_COMMON3 = 0 '總托盤數
        Dim ACC_COMMON4 = ""
        Dim ACC_COMMON5 = ""
        Dim ACC_COMMON6 = ""
        Dim ACC_COMMON7 = ""
        Dim ACC_COMMON8 = ""
        Dim ACC_COMMON9 = ""
        Dim ACC_COMMON10 = ""

        Dim objInfo = New clsWMS_CT_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10)
        Dim objTmp As clsWMS_CT_INVENTORY_COMPARISON = Nothing
        If ret_dicAddInventoryComparison.TryGetValue(objInfo.gid, objTmp) = False Then
          ret_dicAddInventoryComparison.Add(objInfo.gid, objInfo)
        Else
          objTmp.WMS_STOCK_QTY += WMS_STOCK_QTY
          objTmp.WMS_UNFINISH_QTY += WMS_UNFINISH_QTY
        End If
      Next
      SendMessageToLog("DischargePO_DTL Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      SendMessageToLog("比對完成。 進行第二步驟檢查。", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Each objInventoryComparison In ret_dicAddInventoryComparison.Values
        objInventoryComparison.WMS_COMPARSON_QTY = objInventoryComparison.WMS_STOCK_QTY - objInventoryComparison.WMS_UNFINISH_QTY 'WMS的比對數量
        objInventoryComparison.QUANTITY_VARIANCE = objInventoryComparison.ERP_COMPARSON_QTY - objInventoryComparison.WMS_COMPARSON_QTY '差異數量

        '如果數量有落差 則找出對應的最後一筆托盤及位置
        If objInventoryComparison.QUANTITY_VARIANCE <> 0 Then
          For Each objCarrierItem In dicCarrierItem.Values
            If objCarrierItem.SKU_No = objInventoryComparison.SKU_NO AndAlso objCarrierItem.Lot_No = objInventoryComparison.LOT_NO AndAlso
               objCarrierItem.Owner_No = objInventoryComparison.OWNER_NO AndAlso objCarrierItem.Sub_Owner_No = objInventoryComparison.SUB_OWNER_NO Then

              objInventoryComparison.ACC_COMMON1 = objCarrierItem.Carrier_ID
              Dim objCarrier As clsCarrier = Nothing
              If dicCarrier.TryGetValue(clsCarrier.Get_Combination_Key(objCarrierItem.Carrier_ID), objCarrier) Then
                objInventoryComparison.ACC_COMMON2 = objCarrier.Location_No
              End If
            End If
          Next
        End If
      Next

      '在最後的地方 加入標誌
      ret_dicAddInventoryComparison.Values(ret_dicAddInventoryComparison.Count - 1).ACC_COMMON10 = "1"

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
#End Region


  'SQL
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByVal ret_dicAddInventoryComparison As Dictionary(Of String, clsWMS_CT_INVENTORY_COMPARISON),
                           ByRef lstSql As List(Of String),
                           ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '取得要 寫入的資料 '有要寫入的就全刪
      If ret_dicAddInventoryComparison.Any = True Then
        ret_dicAddInventoryComparison.First.Value.O_Add_Delete_SQLString(lstSql)
        ret_dicAddInventoryComparison.First.Value.O_Add_Delete_SQLString(lstQueueSql)
      End If

      For Each obj In ret_dicAddInventoryComparison.Values
        If obj.O_Add_Insert_SQLString(lstSql, lstQueueSql) = False Then
          ret_strResultMsg = "Get Insert InventoryComparison SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                     ByRef lstSql As List(Of String),
                                     ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        ret_strResultMsg = "HostHandler Update DB Failed"
        Return False
      End If
      Common_DBManagement.AddQueued_BatchUpdate(lstQueueSql, False)
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
