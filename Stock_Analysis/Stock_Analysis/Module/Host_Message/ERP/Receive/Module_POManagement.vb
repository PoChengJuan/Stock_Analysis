'20200608
'V1.0.0
'Vito
'接收到ERP的所有單據

Imports eCA_TransactionMessage
Imports eCA_HostObject

Module Module_POManagement
  Public Function O_POManagement(ByRef objPO As MSG_eWMSMessage,
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
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '檢查資料
      If Check_Data(objPO, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(objPO, ret_strResultMsg, Host_Command, dicAdd_PO, dicUpdate_PO, dicAdd_PO_Line, dicDelete_PO_Line, dicUpdate_PO_Line, dicAdd_PO_DTL, dicDelete_PO_DTL, dicUpdate_PO_DTL, dicAdd_PO_DTL_TRANSACTION, dicDelete_PO_DTL_TRANSACTION, dicUpdate_PO_DTL_TRANSACTION) = False Then
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
        Return False
      End If
      '執行SQL與更新物件
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

  Private Function Check_Data(ByVal objPO As MSG_eWMSMessage,
                                                          ByRef ret_strResultMsg As String) As Boolean
    Try
      If objPO.EventID = "" Then
        ret_strResultMsg = "EventID is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      For Each objNoticeDataInfo In objPO.NoticeDataList.NoticeDataInfo
        '檢查NoticeType是否為空
        If objNoticeDataInfo.NoticeType = "" Then
          ret_strResultMsg = "NoticeType is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查NoticeId是否為空
        If objNoticeDataInfo.NoticeId = "" Then
          ret_strResultMsg = "NoticeId is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        For Each objNoticeDetailDataInfo In objNoticeDataInfo.NoticeDetailDataList.NoticeDetailDataInfo
          '檢查NoticeSerialId是否為空
          If objNoticeDetailDataInfo.NoticeSerialId = "" Then
            ret_strResultMsg = "NoticeSerialId is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          '檢查TransferType是否為空
          If objNoticeDetailDataInfo.TransferType = "" Then
            ret_strResultMsg = "TransferType is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          '檢查SKU是否為空
          If objNoticeDetailDataInfo.SKU = "" Then
            ret_strResultMsg = "SKU is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          '檢查QTY是否為空
          If objNoticeDetailDataInfo.QTY = "" Then
            ret_strResultMsg = "QTY is empty"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          '非數字做CINT會直接進EXCEPTION
          'Dim QTY = CInt(objNoticeDetailDataInfo.QTY)
          If IsNumeric(objNoticeDetailDataInfo.QTY) = False Then
            ret_strResultMsg = "QTY 不為數字"
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
  Private Function Get_UpdateData(ByVal objPO_Data As MSG_eWMSMessage,
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
                                                                  ByRef ret_dicUpdate_PO_DTL_TRANSACTION As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)) As Boolean
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
      'Dim Companyid = objSendWorkData.Companyid
      Dim Create_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim NOTICE_TYPE = ""
      For Each objNoticeDataInfo In objPO_Data.NoticeDataList.NoticeDataInfo
        Dim PO_ID As String = objNoticeDataInfo.NoticeId
        If tmp_dicPOID.ContainsKey(PO_ID) = False Then
          tmp_dicPOID.Add(PO_ID, PO_ID)
        End If
      Next
      '使用dicPO取得資料庫裡的PO資料
      If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(tmp_dicPOID, tmp_dicPO) = False Then
        ret_strResultMsg = "WMS get PO data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '使用dicPO取得資料庫裡的PO_Line資料
      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicPOID, tmp_dicPO_Line) = False Then
        ret_strResultMsg = "WMS get PO_Line data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '使用dicPO取得資料庫裡的PO_DTL資料
      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicPOID, tmp_dicPO_DTL) = False Then
        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '使用dicPO取得資料庫裡的PO_DTLTRANSACTION資料
      If gMain.objHandling.O_GetDB_dicPODTLTRANSACTIONBydicPO_ID(tmp_dicPOID, tmp_dicPO_DTL_TRANSACTION) = False Then
        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      For Each objNoticeDataInfo In objPO_Data.NoticeDataList.NoticeDataInfo
        Dim obj = objNoticeDataInfo.NoticeDetailDataList.NoticeDetailDataInfo.First

        Dim PO_ID = objNoticeDataInfo.NoticeId  '製令單號_製令單別(ERP使用)
        Dim PO_TYPE1 = ""
        Dim PO_TYPE2 = ""
        Dim PO_TYPE3 = ""
        Dim WO_TYPE = ""
        Dim H_PO_ORDER_TYPE = ""
        NOTICE_TYPE = objNoticeDataInfo.NoticeType
        'Select Case NOTICE_TYPE
        '  Case "asfi511"    '工單成套發料單維護作業
        '    PO_TYPE1 = enuPOType_1.Picking_out
        '    PO_TYPE2 = enuPOType_2.material_out
        '    WO_TYPE = enuWOType.Discharge
        '    H_PO_ORDER_TYPE = enuOrderType.material_out
        '  Case "apmt720"     '採購入庫異動維護作業
        '    PO_TYPE1 = enuPOType_1.Combination_in
        '    PO_TYPE2 = enuPOType_2.material_in
        '    WO_TYPE = enuWOType.Receipt
        '    H_PO_ORDER_TYPE = enuOrderType.material_in
        '  Case "asft620"    '工單完工入庫維護作業
        '    PO_TYPE1 = enuPOType_1.Produce_in
        '    PO_TYPE2 = enuPOType_2.SKU_in
        '    WO_TYPE = enuWOType.Receipt
        '    H_PO_ORDER_TYPE = enuOrderType.SKU_in
        '  Case "aimt324"    '倉庫間直接調撥作業 (多行)
        '    PO_TYPE1 = enuPOType_1.Transaction
        '    PO_TYPE2 = enuPOType_2.transaction_account
        '    WO_TYPE = enuWOType.Transform
        '    H_PO_ORDER_TYPE = enuOrderType.transaction_account
        '  Case "aimt301"    '倉庫雜項發料作業
        '    PO_TYPE1 = enuPOType_1.Picking_out
        '    PO_TYPE2 = enuPOType_2.semiSKU_out
        '    WO_TYPE = enuWOType.Discharge
        '    H_PO_ORDER_TYPE = enuOrderType.semiSKU_out
        '  Case "aimt302"    '倉庫雜項收料作業
        '    PO_TYPE1 = enuPOType_1.Combination_in
        '    PO_TYPE2 = enuPOType_2.other_in
        '    WO_TYPE = enuWOType.Receipt
        '    H_PO_ORDER_TYPE = enuOrderType.other_in
        '  Case "asfi526"    '工單成套退料單維護作業
        '    PO_TYPE1 = enuPOType_1.Combination_in
        '    PO_TYPE2 = enuPOType_2.asfi526
        '    WO_TYPE = enuWOType.Receipt
        '    H_PO_ORDER_TYPE = enuOrderType.asfi526
        '  Case "asfi512"    '工單超領發料單維護作業
        '    PO_TYPE1 = enuPOType_1.Picking_out
        '    PO_TYPE2 = enuPOType_2.asfi512
        '    WO_TYPE = enuWOType.Discharge
        '    H_PO_ORDER_TYPE = enuOrderType.asfi512
        '  Case "asfi513"    '工單欠料補料單維護作業
        '    PO_TYPE1 = enuPOType_1.Combination_in
        '    PO_TYPE2 = enuPOType_2.asfi513
        '    WO_TYPE = enuWOType.Receipt
        '    H_PO_ORDER_TYPE = enuOrderType.asfi513
        '  Case "asfi514"    '工單領料維護作業
        '    PO_TYPE1 = enuPOType_1.Picking_out
        '    PO_TYPE2 = enuPOType_2.asfi514
        '    WO_TYPE = enuWOType.Discharge
        '    H_PO_ORDER_TYPE = enuOrderType.asfi514
        '  Case "asfi527"    '工單超領退料單維護作業
        '    PO_TYPE1 = enuPOType_1.Combination_in
        '    PO_TYPE2 = enuPOType_2.asfi527
        '    WO_TYPE = enuWOType.Receipt
        '    H_PO_ORDER_TYPE = enuOrderType.asfi527
        '  Case "asfi528"    '工單一般退料單維護作業
        '    PO_TYPE1 = enuPOType_1.Combination_in
        '    PO_TYPE2 = enuPOType_2.back_in
        '    WO_TYPE = enuWOType.Receipt
        '    H_PO_ORDER_TYPE = enuOrderType.back_in
        '  Case "asfi529"    '工單領退料維護作業
        '    PO_TYPE1 = enuPOType_1.Combination_in
        '    PO_TYPE2 = enuPOType_2.asfi529
        '    WO_TYPE = enuWOType.Receipt
        '    H_PO_ORDER_TYPE = enuOrderType.asfi529
        '  Case Else
        '    ret_strResultMsg = "單別未定義"
        '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '    Return False
        'End Select

        'If obj.TransferType = "I" Then
        '  PO_TYPE1 = enuPOType_1.Combination_in
        'ElseIf obj.TransferType = "O" Then
        '  PO_TYPE1 = enuPOType_1.Picking_out
        'End If

        Dim PRIORITY = "50"
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
        Dim H_PO1 = objNoticeDataInfo.NoticeDate
        Dim H_PO2 = objNoticeDataInfo.NoticeTime
        Dim H_PO3 = objNoticeDataInfo.Spare1
        Dim H_PO4 = objNoticeDataInfo.Spare2
        Dim H_PO5 = objNoticeDataInfo.Spare3
        Dim H_PO6 = objNoticeDataInfo.Spare4
        Dim H_PO7 = objNoticeDataInfo.Spare5
        Dim H_PO8 = ""
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
        Dim PO_KEY1 = ""
        Dim PO_KEY2 = ""
        Dim PO_KEY3 = ""
        Dim PO_KEY4 = ""
        Dim PO_KEY5 = ""

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
            objNewPO.PO_Type1 = PO_TYPE1
            objNewPO.User_ID = User_ID
            objNewPO.Write_Off_No = PO_ID
            objNewPO.H_PO_STEP_NO = H_PO_STEP_NO
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

        If NOTICE_TYPE = "aimt324" Then
          '檢查入出的數量是否一致
          Dim Odd_QTY = 0
          Dim Odd_Serial = ""
          Dim Even_QTY = 0
          Dim Even_Serial = ""
          For Each objPO_DTL In objNoticeDataInfo.NoticeDetailDataList.NoticeDetailDataInfo
            If objPO_DTL.NoticeSerialId Mod 2 = 1 Then
              '取奇數位數量
              Odd_QTY = objPO_DTL.QTY
              Odd_Serial = objPO_DTL.NoticeSerialId
            Else
              '取偶數位數量
              Even_QTY = objPO_DTL.QTY
              Even_Serial = objPO_DTL.NoticeSerialId
            End If
            If Odd_QTY <> 0 And Even_QTY <> 0 Then
              If Odd_QTY <> Even_QTY Then
                ret_strResultMsg = "PO_ID：" & PO_ID & ",序號：" & Odd_Serial & "的數量：" & Odd_QTY & " ,與序號：" & Even_Serial & "的數量：" & Even_QTY & "不批配"
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If
              Odd_QTY = 0
              Even_QTY = 0
              Odd_Serial = ""
              Even_Serial = ""
            End If
          Next
        End If

        For Each objPO_DTL In objNoticeDataInfo.NoticeDetailDataList.NoticeDetailDataInfo
          '檢查料品主檔是否存在
          Dim dicSKU As New Dictionary(Of String, clsSKU)
          If gMain.objHandling.O_GetDB_dicSKUBySKUNo(objPO_DTL.SKU, dicSKU) = True Then
            If dicSKU.Any Then

            Else

            End If
            'ret_strResultMsg = "料品不存在無法取得庫別資訊, PO_ID=" & PO_ID & ", SKU_NO=" & objPO_DTL.SKU & " 請先建立品號資料"
            'SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            'Return False
          End If

          Dim SKU_NO = objPO_DTL.SKU
          If dicSKU.Any = False Then
            '新增料品主檔
            Dim SKU_ID1 = objPO_DTL.SKU
            Dim SKU_ID2 = ""
            Dim SKU_ID3 = ""
            Dim SKU_ALIS1 = objPO_DTL.SKUName
            Dim SKU_ALIS2 = ""
            Dim SKU_DESC = ""
            Dim SKU_CATALOG = enuSKU_CATALOG.NULL
            Dim SKU_TYPE1 = ""
            Dim SKU_TYPE2 = ""
            Dim SKU_TYPE3 = ""
            Dim SKU_COMMON1 = ""
            Dim SKU_COMMON2 = ""
            Dim SKU_COMMON3 = ""
            Dim SKU_COMMON4 = ""
            Dim SKU_COMMON5 = ""
            Dim SKU_COMMON6 = ""
            Dim SKU_COMMON7 = ""
            Dim SKU_COMMON8 = ""
            Dim SKU_COMMON9 = ""
            Dim SKU_COMMON10 = ""
            Dim SKU_L = 0
            Dim SKU_W = 0
            Dim SKU_H = 0
            Dim SKU_WEIGHT = 0
            Dim SKU_VALUE = 0
            Dim SKU_UNIT = ""
            Dim INBOUND_UNIT = ""
            Dim OUTBOUND_UNIT = ""
            Dim HIGH_WATER = 0
            Dim LOW_WATER = 0
            Dim AVAILABLE_DAYS = 0
            Dim SAVE_DAYS = ""
            'Dim CREATE_TIME = ""
            Dim UPDATE_TIME = ""
            Dim WEIGHT_DIFFERENCE = 0
            Dim ENABLE = True
            Dim EFFECTIVE_DATE = ""
            Dim FAILURE_DATE = ""
            Dim QC_METHOD = ""
            Dim RECEIPT_DAYS = ""
            Dim DISCHARGE_DAYS = ""
            Dim RETURN_DAYS = ""
            Dim ASSIGN_AREA_NO = ""
            'Dim COMMENTS = ""

            Dim objNewSKU = New clsSKU(SKU_NO, SKU_ID1, SKU_ID2, SKU_ID3, SKU_ALIS1, SKU_ALIS2, SKU_DESC, SKU_CATALOG, SKU_TYPE1, SKU_TYPE2,
                                       SKU_TYPE3, SKU_COMMON1, SKU_COMMON2, SKU_COMMON3, SKU_COMMON4, SKU_COMMON5, SKU_COMMON6, SKU_COMMON7, SKU_COMMON8,
                                       SKU_COMMON9, SKU_COMMON10, SKU_L, SKU_W, SKU_H, SKU_WEIGHT, SKU_VALUE, SKU_UNIT, INBOUND_UNIT, OUTBOUND_UNIT, HIGH_WATER,
                                       LOW_WATER, AVAILABLE_DAYS, SAVE_DAYS, Create_Time, UPDATE_TIME, WEIGHT_DIFFERENCE, ENABLE, EFFECTIVE_DATE,
                                       FAILURE_DATE, QC_METHOD, "")
            If dicAdd_SKU.ContainsKey(objNewSKU.gid) = False Then
              dicAdd_SKU.Add(objNewSKU.gid, objNewSKU)
            End If
            'Dim objSKU = dicSKU.First
            'OWNER_NO = objSKU.SKU_COMMON4 '主要庫別
          End If
          Dim Serial_ID As Integer = 0
          Dim PO_LINE_NO = objPO_DTL.NoticeSerialId
          If NOTICE_TYPE = "aimt324" Then
            If objPO_DTL.NoticeSerialId Mod 2 = 1 Then  '奇數
              'Serial_ID = (CInt(objPO_DTL.NoticeSerialId) / 2)
              'Serial_ID = Serial_ID + 1

              Serial_ID = CInt(objPO_DTL.NoticeSerialId)
              Serial_ID = Fix(Serial_ID / 2)
              Serial_ID = Serial_ID + 1
            Else                                        '偶數
              Serial_ID = CInt(objPO_DTL.NoticeSerialId) / 2
            End If
            PO_LINE_NO = Serial_ID.ToString

            PO_LINE_NO = PO_LINE_NO.PadLeft(4, "0")
          End If
          Dim QTY = 1
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

          Dim PO_SERIAL_NO = objPO_DTL.NoticeSerialId

          If NOTICE_TYPE = "aimt324" Then
            If objPO_DTL.NoticeSerialId Mod 2 = 1 Then  '奇數
              Serial_ID = CInt(objPO_DTL.NoticeSerialId)
              Serial_ID = Fix(Serial_ID / 2)
              Serial_ID = Serial_ID + 1
            Else                                        '偶數
              Serial_ID = CInt(objPO_DTL.NoticeSerialId) / 2
            End If
            PO_SERIAL_NO = Serial_ID.ToString
            PO_SERIAL_NO = PO_SERIAL_NO.PadLeft(4, "0")
          Else
            PO_SERIAL_NO = PO_SERIAL_NO.PadLeft(4, "0")
          End If
          Dim WORKING_TYPE = ""
          Dim WORKING_SERIAL_NO = ""
          Dim WORKING_SERIAL_SEQ = ""
          'Dim SKU_NO = ""
          Dim LOT_NO = objPO_DTL.LotId
          'Dim QTY = ""
          Dim QTY_PROCESS = 0
          'Dim QTY_FINISH = 0
          Dim PODTL_STATUS = enuPODTLStatus.Queued
          Dim PACKAGE_ID = ""
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
          Dim BND = enuBND.None
          Dim QC_STATUS = enuQCStatus.NULL
          Dim FROM_OWNER_ID = ""
          Dim FROM_SUB_OWNER_ID = ""
          Dim TO_OWNER_ID = ""
          Dim TO_SUB_OWNER_ID = ""
          Dim FACTORY_ID = ""
          Dim DEST_AREA_ID = ""
          Dim DEST_LOCATION_ID = ""
          Dim CLOSE_ABLE = 1
          Dim H_POD_STEP_NO = enuStepNo.Queue
          Dim H_POD_MOVE_TYPE = ""
          Dim H_POD_FINISH_TIME = ""
          Dim H_POD_BILLING_DATE = ""
          Dim H_POD_CREATE_TIME = ""
          Dim H_POD1 = objPO_DTL.Unit
          Dim H_POD2 = objPO_DTL.WH
          Dim H_POD3 = objPO_DTL.Location
          Dim H_POD4 = objPO_DTL.SKUName
          Dim H_POD5 = CDbl(objPO_DTL.QTY) 'objPO_DTL.WEIGHT
          Dim H_POD6 = objPO_DTL.LENGTH
          Dim H_POD7 = objPO_DTL.WIDTH
          Dim H_POD8 = objPO_DTL.WO
          Dim H_POD9 = objPO_DTL.GRADE
          Dim H_POD10 = objPO_DTL.Spare1
          Dim H_POD11 = objPO_DTL.Spare2
          Dim H_POD12 = objPO_DTL.Spare3
          Dim H_POD13 = objPO_DTL.Spare4
          Dim H_POD14 = objPO_DTL.Spare5
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





          Dim PO_DTL_Key = clsPO_DTL.Get_Combination_Key(PO_ID, PO_SERIAL_NO)
          Dim PO_DTL_TRANSACTION_Key = clsWMS_T_PO_DTL_TRANSACTION.Get_Combination_Key(PO_ID, PO_SERIAL_NO)


          '判斷是否為調播單
          If NOTICE_TYPE = "aimt324" Then
            If objPO_DTL.TransferType = "O" Then
#Region "轉播單出"
              If tmp_dicPO_DTL.ContainsKey(PO_DTL_Key) = True Then  '單據存在
                Dim objNewPO_DTL = tmp_dicPO_DTL.Item(PO_DTL_Key).Clone()
                With objNewPO_DTL
                  .SKU_NO = SKU_NO
                  .LOT_NO = LOT_NO
                  .QTY = QTY
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
              End If
#End Region

            ElseIf objPO_DTL.TransferType = "I" Then
#Region "轉播單入"
              If tmp_dicPO_DTL_TRANSACTION.ContainsKey(PO_DTL_TRANSACTION_Key) = True Then
                '單據存在
                Dim objNewPO_DTL_TRANSACTION = tmp_dicPO_DTL_TRANSACTION.Item(PO_DTL_TRANSACTION_Key).Clone()
                With objNewPO_DTL_TRANSACTION
                  .SKU_NO = SKU_NO
                  .LOT_NO = LOT_NO
                  .QTY = QTY
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
                Dim objNewPO_DTL_TRANSACTION = New clsWMS_T_PO_DTL_TRANSACTION(PO_ID, PO_SERIAL_NO, enuTransaction_Type.Transaction_IN, SKU_NO, LOT_NO, QTY, PACKAGE_ID,
                                               ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8,
                                               ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5,
                                                STORAGE_TYPE, BND, QC_STATUS, FROM_OWNER_ID, FROM_SUB_OWNER_ID, TO_OWNER_ID, TO_SUB_OWNER_ID, FACTORY_ID, DEST_AREA_ID,
                                                DEST_LOCATION_ID, H_POD1, H_POD2, H_POD3, H_POD4, H_POD5, H_POD6, H_POD7, H_POD8, H_POD9, H_POD10, H_POD11, H_POD12, H_POD13,
                                                H_POD14, H_POD15, H_POD16, H_POD17, H_POD18, H_POD19, H_POD20, H_POD21, H_POD22, H_POD23, H_POD24, H_POD25)
                If ret_dicAdd_PO_DTL_TRANSACTION.ContainsKey(objNewPO_DTL_TRANSACTION.gid) = False Then
                  ret_dicAdd_PO_DTL_TRANSACTION.Add(objNewPO_DTL_TRANSACTION.gid, objNewPO_DTL_TRANSACTION)
                End If
              End If
#End Region

            End If

          Else '一般入出庫單據
#Region "一般單據的PO_DTL"
            If tmp_dicPO_DTL.ContainsKey(PO_DTL_Key) = True Then  '單據存在
              Dim objNewPO_DTL = tmp_dicPO_DTL.Item(PO_DTL_Key).Clone()
              With objNewPO_DTL
                .SKU_NO = SKU_NO
                .LOT_NO = LOT_NO
                .QTY = QTY
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
            End If
#End Region

          End If

        Next

        If dicAdd_SKU.Any Then
          If Module_Send_WMSMessage.Send_T2F3U1_SKUManagement_to_WMS(ret_strResultMsg, dicAdd_SKU, Host_Command, "Create") = False Then      'Vito_20203
            Return False
          End If
        End If

        If NOTICE_TYPE = "aimt324" Then
          If ret_dicAdd_PO.Any Then
            If Module_Send_WMSMessage.Send_T5F5U1_TransactionOederManagement_to_WMS(ret_strResultMsg, ret_dicAdd_PO, ret_dicAdd_POLine, ret_dicAdd_PO_DTL, ret_dicAdd_PO_DTL_TRANSACTION, Host_Command, "Create") Then

            End If
          End If
          If ret_dicUpdate_PO.Any Then
            If Module_Send_WMSMessage.Send_T5F5U1_TransactionOederManagement_to_WMS(ret_strResultMsg, ret_dicUpdate_PO, ret_dicUpdate_POLine, ret_dicUpdate_PO_DTL, ret_dicUpdate_PO_DTL_TRANSACTION, Host_Command, "Modify") Then

            End If
          End If
        Else
          If ret_dicAdd_PO.Any Then
            If Module_Send_WMSMessage.Send_T5F1U1_POManagement_to_WMS(ret_strResultMsg, ret_dicAdd_PO, ret_dicAdd_POLine, ret_dicAdd_PO_DTL, Host_Command, "Create") = False Then
              Return False
            End If
          End If
          If ret_dicUpdate_PO.Any Then
            If Module_Send_WMSMessage.Send_T5F1U1_POManagement_to_WMS(ret_strResultMsg, ret_dicUpdate_PO, ret_dicUpdate_POLine, ret_dicUpdate_PO_DTL, Host_Command, "Modify") = False Then
              Return False
            End If
          End If
        End If




      Next
      ret_dicDelete_POLine = tmp_dicPO_Line
      ret_dicDelete_PO_DTL = tmp_dicPO_DTL
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
End Module
