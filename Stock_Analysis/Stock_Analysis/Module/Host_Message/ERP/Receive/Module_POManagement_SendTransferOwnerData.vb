'20200608
'V1.0.0
'Vito
'接收到ERP的領/退料單據

Imports eCA_TransactionMessage
Imports eCA_HostObject
Imports System.Math

Module Module_POManagement_SendTransferOwnerData
  Public Function O_POManagement_SendTransferOwnerData(ByRef objPO As MSG_SendTransferOwnerData,
                                     ByRef ret_strResultMsg As String) As Boolean

    Try
      '要變更的資料
      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
      Dim dicAdd_PO As New Dictionary(Of String, clsPO)
      Dim dicUpdate_PO As New Dictionary(Of String, clsPO)
      Dim dicAdd_PO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim dicDelete_PO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim dicUpdate_PO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim dicAdd_PO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim dicDelete_PO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim dicUpdate_PO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim dicAdd_PO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
      Dim dicDelete_PO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
      Dim dicUpdate_PO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
      Dim dicGluePO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim PO_ID = ""
      Dim PO_TYPE = ""
      Dim DocTypeCode = ""
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '檢查資料
      If Check_Data(objPO, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(PO_ID, PO_TYPE, DocTypeCode, objPO, ret_strResultMsg, Host_Command, dicAdd_PO, dicUpdate_PO, dicAdd_PO_Line, dicDelete_PO_Line, dicUpdate_PO_Line, dicAdd_PO_DTL, dicDelete_PO_DTL, dicUpdate_PO_DTL, dicAdd_PO_DTL_TRANSACTION, dicDelete_PO_DTL_TRANSACTION, dicUpdate_PO_DTL_TRANSACTION, dicGluePO_DTL) = False Then
        'If DocTypeCode = "11" Then
        '  SendTransactionData_Normal(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        'Else
        '  SendTransactionData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        'End If
        SendTransactionOwnerData(enuRtnCode.Fail, PO_TYPE, PO_ID, "")
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
        'If DocTypeCode = "11" Then
        '  SendTransactionData_Normal(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        'Else
        '  SendTransactionData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        'End If
        SendTransactionOwnerData(enuRtnCode.Fail, PO_TYPE, PO_ID, "")
        Return False
      End If
      '執行SQL與更新物件
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
        'If DocTypeCode = "11" Then
        '  SendTransactionData_Normal(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        'Else
        '  SendTransactionData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        'End If
        SendTransactionOwnerData(enuRtnCode.Fail, PO_TYPE, PO_ID, "")
        Return False
      End If










      'If DocTypeCode = "11" Then
      '  SendTransactionData_Normal(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)
      'Else
      '  SendTransactionData(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)
      'End If
      SendTransactionOwnerData(enuRtnCode.Sucess, PO_TYPE, PO_ID, "")
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Check_Data(ByVal objPO As MSG_SendTransferOwnerData,
                                                          ByRef ret_strResultMsg As String) As Boolean
    Try
      If objPO.EventID = "" Then
        ret_strResultMsg = "EventID is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      For Each objTransferOwnerDataInfo In objPO.TransferOwnerDataList.TransferOwnerDataInfo
        '檢查POId是否為空
        If objTransferOwnerDataInfo.POId = "" Then
          ret_strResultMsg = "POId is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查POType是否為空
        If objTransferOwnerDataInfo.POType = "" Then
          ret_strResultMsg = "POType is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查DocTypeCode是否為空
        'If objTransactionDataInfo.DocTypeCode = "" Then
        '  ret_strResultMsg = "DocTypeCode is empty"
        '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  Return False
        'End If
        For Each objTransferOwnerDetailDataInfo In objTransferOwnerDataInfo.TransferOwnerDetailDataList.TransferOwnerDetailDataInfo
          '檢查NoticeSerialId是否為空
          If objTransferOwnerDetailDataInfo.SerialId = "" Then
            ret_strResultMsg = "SerialId is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          '檢查SKU是否為空
          If objTransferOwnerDetailDataInfo.SKU = "" Then
            ret_strResultMsg = "SKU is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          '檢查SKU是否為空
          'If objTransactionDetailDataInfo.Unit = "" Then
          '  ret_strResultMsg = "Unit is empty"
          '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '  Return False
          'End If
          '檢查QTY是否為空
          If objTransferOwnerDetailDataInfo.CheckQty = "" Then
            ret_strResultMsg = "CheckQty is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          'Dim QTY = CInt(objTransferOwnerDetailDataInfo.CheckQty)
          If IsNumeric(objTransferOwnerDetailDataInfo.CheckQty) = False Then
            ret_strResultMsg = "CheckQty 不為數字"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        Next
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '新增資料或得到要更新的資料
  Private Function Get_UpdateData(ByRef ret_PO_ID As String, ByRef ret_PO_TYPE As String, ByRef DocTypeCode As String, ByVal objPO_Data As MSG_SendTransferOwnerData,
                                                                  ByRef ret_strResultMsg As String,
                                                                  ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                                                                  ByRef ret_dicAdd_PO As Dictionary(Of String, clsPO),
                                                                  ByRef ret_dicUpdate_PO As Dictionary(Of String, clsPO),
                                                                  ByRef ret_dicAdd_POLine As Dictionary(Of String, clsPO_LINE),
                                                                  ByRef ret_dicDelete_POLine As Dictionary(Of String, clsPO_LINE),
                                                                  ByRef ret_dicUpdate_POLine As Dictionary(Of String, clsPO_LINE),
                                                                  ByRef ret_dicAdd_PO_DTL As Dictionary(Of String, clsPO_DTL),
                                                                  ByRef ret_dicDelete_PO_DTL As Dictionary(Of String, clsPO_DTL),
                                                                  ByRef ret_dicUpdate_PO_DTL As Dictionary(Of String, clsPO_DTL),
                                                                  ByRef ret_dicAdd_PO_DTL_TRANSACTION As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION),
                                                                  ByRef ret_dicDelete_PO_DTL_TRANSACTION As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION),
                                                                  ByRef ret_dicUpdate_PO_DTL_TRANSACTION As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION),
                                                                  ByRef ret_dicGluePO_DTL As Dictionary(Of String, clsPO_DTL)) As Boolean
    Try
      Dim dicAdd_SKU As New Dictionary(Of String, clsSKU)
      'Dim dicUpdate_SKU As New Dictionary(Of String, clsSKU)
      '取得所有的PO單號
      Dim tmp_dicPOID As New Dictionary(Of String, String)
      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim tmp_dicPO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
      Dim User_ID As String = objPO_Data.WebService_ID
      Dim Event_ID As String = objPO_Data.EventID
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim LotManagement = "N"
      'Dim Companyid = objSendWorkData.Companyid
      Dim Create_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim PickingData_TYPE = ""
      Dim Action As String = "Create"
      'Dim bln_DocTypeisNormal = False
      For Each objTransactionOwnerDataInfo In objPO_Data.TransferOwnerDataList.TransferOwnerDataInfo
        Dim PO_ID As String = objTransactionOwnerDataInfo.POType & "_" & objTransactionOwnerDataInfo.POId
        If tmp_dicPOID.ContainsKey(PO_ID) = False Then
          tmp_dicPOID.Add(PO_ID, PO_ID)
        End If
      Next

      '使用dicPO取得資料庫裡的PO資料

      If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(tmp_dicPOID, tmp_dicPO) = True Then
        Action = "Modify"
      End If
      If Action = "Modify" Then
        '使用dicPO取得資料庫裡的PO_Line資料
        If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicPOID, tmp_dicPO_Line) = False Then
          ret_strResultMsg = "Get DB WMS_T_PO_Line Data Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '使用dicPO取得資料庫裡的PO_DTL資料
        If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicPOID, tmp_dicPO_DTL) = False Then
          ret_strResultMsg = "Get DB WMS_T_PO_DTL Data Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '使用dicPO取得資料庫裡的PO_DTLTRANSACTION資料
        If gMain.objHandling.O_GetDB_dicPODTLTRANSACTIONBydicPO_ID(tmp_dicPOID, tmp_dicPO_DTL_TRANSACTION) = False Then
          ret_strResultMsg = "Get DB WMS_T_PO_DTL_Transaction Data Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      End If




      For Each objTransactionOwnerDataInfo In objPO_Data.TransferOwnerDataList.TransferOwnerDataInfo
        Dim PO_KEY1 = objTransactionOwnerDataInfo.POType
        Dim PO_KEY2 = objTransactionOwnerDataInfo.POId
        Dim PO_KEY3 = ""
        Dim PO_KEY4 = ""
        Dim PO_KEY5 = ""

        Dim PO_ID = PO_KEY1 & "_" & PO_KEY2
        ret_PO_ID = objTransactionOwnerDataInfo.POId
        Dim PO_TYPE1 = enuPOType_1.Transaction
        Dim PO_TYPE2 = enuPOType_2.transaction_account
        Dim PO_TYPE3 = enuPOType_3.None
        Dim WO_TYPE = enuWOType.Transform
        Dim H_PO_ORDER_TYPE = enuOrderType.transaction_account


        Dim PRIORITY = 50
        'Dim CREATE_TIME = ""
        Dim START_TIME = ""
        Dim FINISH_TIME = ""
        'Dim USER_ID = ""
        Dim CUSTOMER_NO = ""
        Dim SUPPLIER_NO = ""
        Dim CLASS_NO = ""
        Dim SHIPPING_NO = ""
        Dim WRITE_OFF_NO = ""
        Dim PO_STATUS = enuPOStatus.Queued
        Dim AUTO_BOUND = False
        Dim H_PO_CREATE_TIME = ""
        Dim H_PO_FINISH_TIME = ""
        Dim H_PO_STEP_NO = enuStepNo.Queue
        'Dim H_PO_ORDER_TYPE = enuOrderType.semiSKU_out
        Dim H_PO1 = objTransactionOwnerDataInfo.POType             '單別
        ret_PO_TYPE = H_PO1
        Dim H_PO2 = objTransactionOwnerDataInfo.TransferOwnerDateTime    '單據日期
        Dim H_PO3 = objTransactionOwnerDataInfo.FactoryId          '廠別代號
        Dim H_PO4 = ""    '生產線別
        Dim H_PO5 = ""              '單據性質
        Dim H_PO6 = ""       '確認碼
        Dim H_PO7 = objTransactionOwnerDataInfo.TransferOutOwner '轉出庫
        Dim H_PO8 = objTransactionOwnerDataInfo.TransferInOwner  '轉入庫
        Dim H_PO9 = ""
        Dim H_PO10 = ""
        Dim H_PO11 = ""
        Dim H_PO12 = ""
        Dim H_PO13 = ""
        Dim H_PO14 = ""
        Dim H_PO15 = ""
        Dim H_PO16 = ""
        Dim H_PO17 = ""
        Dim H_PO18 = ""
        Dim H_PO19 = ""
        Dim H_PO20 = ""


        For Each objPO_DTL In objTransactionOwnerDataInfo.TransferOwnerDetailDataList.TransferOwnerDetailDataInfo
          '檢查料品主檔是否存在
          Dim dicSKU As New Dictionary(Of String, clsSKU)
          '判斷是否是膠塊
          Dim blnGlue As Boolean = False
          'If gMain.objHandling.O_GetDB_dicSKUBySKUNo(objPO_DTL.SKU, dicSKU) = True Then
          '  If dicSKU.Any Then
          '    Dim objSKU = dicSKU.First.Value
          '    LotManagement = objSKU.SKU_COMMON9

          '    '將需要另外處理的紀錄起來
          '    If objSKU.SKU_TYPE2 = "1" Then  '膠塊
          '      blnGlue = True
          '    End If
          '  Else
          '    ret_strResultMsg = "料品不存在無法取得庫別資訊, PO_ID=" & PO_ID & ", SKU_NO=" & objPO_DTL.SKU & " 請先建立品號資料"
          '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '    Return False
          '  End If
          'End If
          Dim dicExistSKU As New Dictionary(Of String, clsSKU)
          gMain.objHandling.O_GetDB_dicSKUBySKU_ID1_SKU_ID2(objPO_DTL.SKU, objTransactionOwnerDataInfo.TransferOutOwner, dicExistSKU)

          Dim SKU_NO = "" '
          If dicExistSKU.Any = False Then
            ' SKU_NO = dicExistSKU.First.Value.SKU_NO
          Else
            SKU_NO = dicExistSKU.First.Value.SKU_NO
          End If

          Dim Serial_ID As Integer = 0
          Dim PO_LINE_NO = objPO_DTL.SerialId

          Dim QTY As Integer = CInt(objPO_DTL.CheckQty)
          QTY = Abs(QTY)  '取絕對值
          'If SetQTYByPackeUnit(objPO_DTL.SKU, objPO_DTL.Qty, objPO_DTL.Unit, QTY, ret_strResultMsg) = False Then
          '  Return False
          'End If
          'If bln_DocTypeisNormal = True Then
          '  '一般單據
          '  If IntegerCheckPositive(objPO_DTL.CheckQty) = True Then
          '    '正數 入庫
          '    PO_TYPE1 = enuPOType_1.Combination_in
          '    PO_TYPE2 = enuPOType_2.normal_in
          '    WO_TYPE = enuWOType.Receipt
          '    H_PO_ORDER_TYPE = enuOrderType.normal_in
          '  Else
          '    '負數 出庫
          '    PO_TYPE1 = enuPOType_1.Picking_out
          '    PO_TYPE2 = enuPOType_2.normal_out
          '    WO_TYPE = enuWOType.Discharge
          '    H_PO_ORDER_TYPE = enuOrderType.normal_out
          '  End If
          'Else


          Dim QTY_FINISH = 0
          Dim H_QTY_PROCESS = 0
          Dim H_POL1 = ""
          Dim H_POL2 = ""
          Dim H_POL3 = ""
          Dim H_POL4 = ""
          Dim H_POL5 = ""
          Dim PO_Line_Key = clsPO_LINE.Get_Combination_Key(PO_ID, PO_LINE_NO)
          If tmp_dicPO_Line.ContainsKey(PO_Line_Key) = True Then  '單據已經存在
            Dim objNewPO_Line = tmp_dicPO_Line.Item(PO_Line_Key).Clone()
            With objNewPO_Line
              .QTY = QTY
              .QTY_FINISH = QTY_FINISH
              .H_QTY_PROCESS = H_QTY_PROCESS
              .H_POL1 = H_POL1
              .H_POL2 = H_POL2
              .H_POL3 = H_POL3
              .H_POL4 = H_POL4
              .H_POL5 = H_POL5
            End With
            If ret_dicUpdate_POLine.ContainsKey(objNewPO_Line.gid) = False Then
              ret_dicUpdate_POLine.Add(objNewPO_Line.gid, objNewPO_Line)
            End If
          Else
            Dim objNewPO_Line = New clsPO_LINE(PO_ID, PO_LINE_NO, QTY, QTY_FINISH, H_QTY_PROCESS, H_POL1, H_POL2, H_POL3, H_POL4, H_POL5)
            If ret_dicAdd_POLine.ContainsKey(objNewPO_Line.gid) = False Then
              ret_dicAdd_POLine.Add(objNewPO_Line.gid, objNewPO_Line)
            End If
          End If

          Dim PO_SERIAL_NO = objPO_DTL.SerialId

          'If PickingData_TYPE = "aimt324" Then
          '  If objPO_DTL.NoticeSerialId Mod 2 = 1 Then  '奇數
          '    Serial_ID = CInt(objPO_DTL.NoticeSerialId)
          '    Serial_ID = Fix(Serial_ID / 2)
          '    Serial_ID = Serial_ID + 1
          '  Else                                        '偶數
          '    Serial_ID = CInt(objPO_DTL.NoticeSerialId) / 2
          '  End If
          '  PO_SERIAL_NO = Serial_ID.ToString
          '  PO_SERIAL_NO = PO_SERIAL_NO.PadLeft(4, "0")
          'Else
          '  PO_SERIAL_NO = PO_SERIAL_NO.PadLeft(4, "0")
          'End If
          Dim WORKING_TYPE = ""
          Dim WORKING_SERIAL_NO = ""
          Dim WORKING_SERIAL_SEQ = ""
          'Dim SKU_NO = ""
          If LotManagement = "Y" Then
            If objPO_DTL.LotId = "" Then
              ret_strResultMsg = "此品號:" & SKU_NO & " 需有批號"
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If
          End If
          Dim LOT_NO = objPO_DTL.LotId.Trim
          'Dim QTY = ""
          Dim QTY_PROCESS = 0
          'Dim QTY_FINISH = 0
          Dim PODTL_STATUS = enuPODTLStatus.Queued
          Dim PACKAGE_ID = objPO_DTL.SN.Trim
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
          Dim COMMENTS = ""
          Dim EXPIRED_DATE = ""
          Dim STORAGE_TYPE = enuStorageType.Store
          Dim BND = enuBND.NB
          Dim QC_STATUS = enuQCStatus.NULL
          Dim FROM_OWNER_ID = ""
          Dim FROM_SUB_OWNER_ID = ""
          Dim TO_OWNER_ID = ""
          Dim TO_SUB_OWNER_ID = ""
          If PO_TYPE1 = enuPOType_1.Combination_in Then
            TO_OWNER_ID = objTransactionOwnerDataInfo.FactoryId
          Else
            FROM_OWNER_ID = objTransactionOwnerDataInfo.FactoryId
          End If
          Dim FACTORY_ID = ""
          Dim DEST_AREA_ID = ""
          Dim DEST_LOCATION_ID = ""
          Dim CLOSE_ABLE = 1
          Dim H_POD_STEP_NO = enuStepNo.Queue
          Dim H_POD_MOVE_TYPE = ""
          Dim H_POD_FINISH_TIME = ""
          Dim H_POD_BILLING_DATE = ""
          Dim H_POD_CREATE_TIME = ""
          Dim H_POD1 = ""             '單位
          Dim H_POD2 = ""
          Dim H_POD3 = ""
          Dim H_POD4 = ""
          Dim H_POD5 = ""
          Dim H_POD6 = ""
          Dim H_POD7 = ""
          Dim H_POD8 = "" '轉出庫
          Dim H_POD9 = "" '轉入庫
          Dim H_POD10 = "" '轉出儲位
          Dim H_POD11 = "" '轉入儲位
          Dim H_POD12 = ""
          Dim H_POD13 = ""
          Dim H_POD14 = ""
          Dim H_POD15 = ""
          Dim H_POD16 = ""
          Dim H_POD17 = ""
          Dim H_POD18 = ""
          Dim H_POD19 = ""
          Dim H_POD20 = ""
          Dim H_POD21 = ""
          Dim H_POD22 = ""
          Dim H_POD23 = ""
          Dim H_POD24 = ""
          Dim H_POD25 = ""

          Dim TRANCATION_TYPE As enuTransaction_Type = enuTransaction_Type.Transaction_T



          Dim PO_DTL_Key = clsPO_DTL.Get_Combination_Key(PO_ID, PO_SERIAL_NO)
          Dim PO_DTL_TRANSACTION_Key = clsWMS_T_PO_DTL_TRANSACTION.Get_Combination_Key(PO_ID, PO_SERIAL_NO)


          '一般入出庫單據
#Region "一般單據的PO_DTL"
          If tmp_dicPO_DTL.ContainsKey(PO_DTL_Key) = True Then  '單據存在
            Dim objNewPO_DTL = tmp_dicPO_DTL.Item(PO_DTL_Key).Clone()
            With objNewPO_DTL
              .SKU_NO = SKU_NO
              .LOT_NO = LOT_NO
              .QTY = QTY
              .PACKAGE_ID = PACKAGE_ID
              .FROM_OWNER_ID = objTransactionOwnerDataInfo.TransferOutOwner
              .SORT_ITEM_COMMON5 = SORT_ITEM_COMMON5
              .H_POD1 = H_POD1
              .H_POD2 = H_POD2
              .H_POD3 = H_POD3
              .H_POD4 = H_POD4
              .H_POD5 = H_POD5
              .H_POD6 = H_POD6
              .H_POD7 = H_POD7
              .H_POD8 = H_POD8
              .H_POD9 = H_POD9
              .H_POD10 = H_POD10
              .H_POD11 = H_POD11
              .H_POD12 = H_POD12
              .H_POD13 = H_POD13
              .H_POD14 = H_POD14
            End With
            If ret_dicUpdate_PO_DTL.ContainsKey(objNewPO_DTL.gid) = False Then
              ret_dicUpdate_PO_DTL.Add(objNewPO_DTL.gid, objNewPO_DTL)
            End If
          Else
            Dim objNewPO_DTL = New clsPO_DTL(PO_ID, PO_LINE_NO, PO_SERIAL_NO, WORKING_TYPE, WORKING_SERIAL_NO, WORKING_SERIAL_SEQ, SKU_NO, LOT_NO, QTY, QTY_PROCESS, QTY_FINISH,
                                             COMMENTS, PACKAGE_ID, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5,
                                             ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10,
                                             SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, STORAGE_TYPE, BND, QC_STATUS, FROM_OWNER_ID,
                                             FROM_SUB_OWNER_ID, TO_OWNER_ID, TO_SUB_OWNER_ID, FACTORY_ID, DEST_AREA_ID, DEST_LOCATION_ID, H_POD_STEP_NO,
                                             H_POD_MOVE_TYPE, H_POD_FINISH_TIME, H_POD_BILLING_DATE, H_POD_CREATE_TIME,
                                             H_POD1, H_POD2, H_POD3, H_POD4, H_POD5, H_POD6, H_POD7, H_POD8, H_POD9, H_POD10,
                                             H_POD11, H_POD12, H_POD13, H_POD14, H_POD15, H_POD16, H_POD17, H_POD18, H_POD19, H_POD20,
                                             H_POD21, H_POD22, H_POD23, H_POD24, H_POD25, PODTL_STATUS, CLOSE_ABLE)
            If ret_dicAdd_PO_DTL.ContainsKey(objNewPO_DTL.gid) = False Then
              ret_dicAdd_PO_DTL.Add(objNewPO_DTL.gid, objNewPO_DTL)
            End If
            objNewPO_DTL.FROM_OWNER_ID = objTransactionOwnerDataInfo.TransferOutOwner

            '紀錄是膠塊的項次
            If blnGlue = True AndAlso PO_TYPE1 = enuPOType_1.Combination_in Then
              If ret_dicGluePO_DTL.ContainsKey(objNewPO_DTL.gid) = False Then
                ret_dicGluePO_DTL.Add(objNewPO_DTL.gid, objNewPO_DTL)
              End If
            End If
          End If
#End Region
#Region "PO_DTL_Transation"
          If tmp_dicPO_DTL_TRANSACTION.ContainsKey(PO_DTL_TRANSACTION_Key) = True Then
            Dim objNewPO_DTL_TRANSACTION = tmp_dicPO_DTL_TRANSACTION.Item(PO_DTL_TRANSACTION_Key).Clone()
            With objNewPO_DTL_TRANSACTION
              .SKU_NO = SKU_NO
              .LOT_NO = LOT_NO
              .QTY = QTY
              .PACKAGE_ID = PACKAGE_ID
              .TO_OWNER_ID = objTransactionOwnerDataInfo.TransferInOwner
              .SORT_ITEM_COMMON5 = SORT_ITEM_COMMON5
              .H_POD1 = H_POD1
              .H_POD2 = H_POD2
              .H_POD3 = H_POD3
              .H_POD4 = H_POD4
              .H_POD5 = H_POD5
              .H_POD6 = H_POD6
              .H_POD7 = H_POD7
              .H_POD8 = H_POD8
              .H_POD9 = H_POD9
              .H_POD10 = H_POD10
              .H_POD11 = H_POD11
              .H_POD12 = H_POD12
              .H_POD13 = H_POD13
              .H_POD14 = H_POD14
            End With
            If ret_dicUpdate_PO_DTL_TRANSACTION.ContainsKey(objNewPO_DTL_TRANSACTION.gid) = False Then
              ret_dicUpdate_PO_DTL_TRANSACTION.Add(objNewPO_DTL_TRANSACTION.gid, objNewPO_DTL_TRANSACTION)
            End If
          Else
            Dim objNewPO_DTL_TRANSACTION = New clsWMS_T_PO_DTL_TRANSACTION(PO_ID, PO_SERIAL_NO, TRANCATION_TYPE, SKU_NO, LOT_NO, QTY, PACKAGE_ID, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5,
                                                                           ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4,
                                                                           SORT_ITEM_COMMON5, STORAGE_TYPE, BND, QC_STATUS, FROM_OWNER_ID, FROM_SUB_OWNER_ID, TO_OWNER_ID, TO_SUB_OWNER_ID, FACTORY_ID, DEST_AREA_ID, DEST_LOCATION_ID,
                                                                           H_POD1, H_POD2, H_POD3, H_POD4, H_POD5, H_POD6, H_POD7, H_POD8, H_POD9, H_POD10, H_POD11, H_POD12, H_POD13, H_POD14, H_POD15, H_POD16, H_POD17, H_POD18, H_POD19,
                                                                           H_POD20, H_POD21, H_POD22, H_POD23, H_POD24, H_POD25)
            If ret_dicAdd_PO_DTL_TRANSACTION.ContainsKey(objNewPO_DTL_TRANSACTION.gid) = False Then
              ret_dicAdd_PO_DTL_TRANSACTION.Add(objNewPO_DTL_TRANSACTION.gid, objNewPO_DTL_TRANSACTION)
            End If
            objNewPO_DTL_TRANSACTION.TO_OWNER_ID = objTransactionOwnerDataInfo.TransferInOwner

          End If
#End Region


        Next

        '調整PO
        If ret_dicAdd_PO.ContainsKey(PO_ID) = False And ret_dicUpdate_PO.ContainsKey(PO_ID) = False Then
          If tmp_dicPO.ContainsKey(PO_ID) = True Then '單據已經存在
            Dim obj_PO As clsPO = tmp_dicPO.Item(PO_ID)
            '先檢查PO的狀態還有類型是否正確
            If obj_PO.PO_Status <> enuPOStatus.Queued Then
              ret_strResultMsg = "單據已執行"
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If

            '更新PO的資料
            Dim objNewPO As clsPO = tmp_dicPO.Item(PO_ID).Clone
            objNewPO.H_PO1 = H_PO1
            objNewPO.H_PO3 = H_PO3
            objNewPO.H_PO4 = H_PO4
            objNewPO.H_PO5 = H_PO5
            objNewPO.H_PO8 = H_PO8
            objNewPO.H_PO9 = H_PO9
            objNewPO.H_PO10 = H_PO10
            ret_dicUpdate_PO.Add(objNewPO.gid, objNewPO)
          Else  '單據不存在
            '建立新的PO資料
            Dim objNewPO = New clsPO(PO_ID, PO_TYPE1, PO_TYPE2, PO_TYPE3, WO_TYPE, PRIORITY, Now_Time, START_TIME, FINISH_TIME, User_ID, CUSTOMER_NO, CLASS_NO, SHIPPING_NO,
                                     PO_STATUS, WRITE_OFF_NO, AUTO_BOUND, H_PO_CREATE_TIME, H_PO_FINISH_TIME, H_PO_STEP_NO, H_PO_ORDER_TYPE,
                                     H_PO1, H_PO2, H_PO3, H_PO4, H_PO5, H_PO6, H_PO7, H_PO8, H_PO9, H_PO10, H_PO11, H_PO12, H_PO13, H_PO14, H_PO15,
                                     H_PO16, H_PO17, H_PO18, H_PO19, H_PO20, SUPPLIER_NO, PO_KEY1, PO_KEY2, PO_KEY3, PO_KEY4, PO_KEY5)
            'Dim objNewPO = New clsPO(PO_ID, PO_TYPE1, PO_TYPE2, "", WO_TYPE, 50, CREATE_TIME, "", "", USER_ID, "", "", "", PO_STATUS, PO_ID, False, CREATE_TIME, "", H_PO_STEP_NO,
            '                         H_PO_ORDER_TYPE, H_PO1, "", H_PO3, H_PO4, H_PO5, "", "", H_PO8, H_PO9, H_PO10, "", "", "", "", "", "", "", "", "", "")
            If ret_dicAdd_PO.ContainsKey(objNewPO.gid) = False Then
              ret_dicAdd_PO.Add(objNewPO.gid, objNewPO)
            End If
          End If
        End If

        If ret_dicAdd_PO.Any Then
          If Module_Send_WMSMessage.Send_T5F5U1_TransactionOederManagement_to_WMS(ret_strResultMsg, ret_dicAdd_PO, ret_dicAdd_POLine, ret_dicAdd_PO_DTL, ret_dicAdd_PO_DTL_TRANSACTION, Host_Command, "Create") = False Then
            Return False
          End If
        End If
        If ret_dicUpdate_PO.Any Then
          If Module_Send_WMSMessage.Send_T5F5U1_TransactionOederManagement_to_WMS(ret_strResultMsg, ret_dicUpdate_PO, ret_dicUpdate_POLine, ret_dicUpdate_PO_DTL, ret_dicUpdate_PO_DTL_TRANSACTION, Host_Command, "Modify") = False Then
            Return False
          End If
        End If




      Next

      'For Each objTransactionDataInfo In objPO_Data.TransactionDataList.TransactionDataInfo
      '  Dim PO_ID As String = objTransactionDataInfo.POId
      '  Dim PO_TYPE As String = objTransactionDataInfo.POType
      '  If objTransactionDataInfo.DocTypeCode = "11" Then
      '    SendTransactionData_Normal(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)
      '  Else
      '    SendTransactionData(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)
      '  End If
      'Next
      'ret_dicDelete_POLine = tmp_dicPO_Line
      'ret_dicDelete_PO_DTL = tmp_dicPO_DTL
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  'SQL
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
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
      '取得要刪除的PO_DTL SQL
      'For Each obj In ret_dicDelete_PO_DTL.Values
      '  If obj.O_Add_Delete_SQLString(lstSql) = False Then
      '    ret_strResultMsg = "Get Delete PO_DTL SQL Failed"
      '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
      ''取得要刪除的PO_Line SQL
      'For Each obj In ret_dicDelete_POLine.Values
      '  If obj.O_Add_Delete_SQLString(lstSql) = False Then
      '    ret_strResultMsg = "Get Delete PO_LINE SQL Failed"
      '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
      ''取得要Update的PO SQL
      'For Each obj In ret_dicUpdate_PO.Values
      '  If obj.O_Add_Update_SQLString(lstSql) = False Then
      '    ret_strResultMsg = "Get Update PO SQL Failed"
      '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
      ''取得要Insert的PO SQL
      'For Each obj In ret_dicAdd_PO.Values
      '  If obj.O_Add_Insert_SQLString(lstSql) = False Then
      '    ret_strResultMsg = "Get Insert PO SQL Failed"
      '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
      ''取得要新增的PO_Line SQL
      'For Each obj In ret_dicAdd_POLine.Values
      '  If obj.O_Add_Insert_SQLString(lstSql) = False Then
      '    ret_strResultMsg = "Get Insert PO_LINE SQL Failed"
      '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
      ''取得要新增的PO_DTL SQL
      'For Each obj In ret_dicAdd_PO_DTL.Values
      '  If obj.O_Add_Insert_SQLString(lstSql) = False Then
      '    ret_strResultMsg = "Get Insert PO_DTL SQL Failed"
      '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
      Return True
    Catch ex As Exception
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
  Private Function Get_PO_TO_WO(ByVal dicPO As Dictionary(Of String, clsPO),
                            ByRef ret_strResultMsg As String,
                            ByRef dic_PO_DTL As Dictionary(Of String, clsPO_DTL),
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                            ByRef ret_Wait_UUID As String) As Boolean
    Try
      Dim User_ID = ""

      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      '先進行資料邏輯檢查
      For Each objPOInfo In dicPO.Values
        '資料檢查
        'Dim IN_PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        User_ID = objPOInfo.User_ID
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


        'Dim COMMENTS As String = objPOInfo.COMMENTS
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


      'Dim SHIPPING_NO = dicPO.Body.WOInfo.SHIPPING_NO

      '寫入Command給WMS
      If Send_Command_to_WMS(ret_strResultMsg, dic_PO_DTL, objUUID, Host_Command, ret_Wait_UUID, User_ID) = False Then
        Return False
      End If




      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Send_Command_to_WMS(ByRef Result_Message As String, ByVal dicUpdate_PO_DTL As Dictionary(Of String, clsPO_DTL), ByRef objUUID As clsUUID,
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                                       ByRef ret_Wait_UUID As String, ByVal User_ID As String, Optional ByVal WOInfo As MSG_T5F1U11_POExecution.clsWOInfo = Nothing) As Boolean
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
      If WOInfo IsNot Nothing Then
        WO_Info.WO_ID = WOInfo.WO_ID
        'WO_Info.RECEIPT_ENABLE = WOInfo.
        WO_Info.COMMENTS = WOInfo.COMMENTS
        WO_Info.SOURCE_AREA_NO = WOInfo.SOURCE_AREA_NO
        WO_Info.SOURCE_LOCATION_NO = WOInfo.SOURCE_LOCATION_NO
        WO_Info.SHIPPING_NO = WOInfo.SHIPPING_NO
        WO_Info.SHIPPING_PRIORITY = WOInfo.SHIPPING_PRIORITY
      Else
        WO_Info.WO_ID = ""
        'WO_Info.RECEIPT_ENABLE = WOInfo.
        WO_Info.COMMENTS = ""
        WO_Info.SOURCE_AREA_NO = ""
        WO_Info.SOURCE_LOCATION_NO = ""
        WO_Info.SHIPPING_NO = ""
        WO_Info.SHIPPING_PRIORITY = ""
      End If



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
