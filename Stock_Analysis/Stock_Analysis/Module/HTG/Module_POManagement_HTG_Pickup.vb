'20220830
'V1.0.0
'Vito
'接收到ERP的領料單據

Imports eCA_TransactionMessage
Imports eCA_HostObject

Module Module_POManagement_HTG_Pickup
  Public Function O_POManagement_HTG_Pickup(ByVal dicPO_ID As Dictionary(Of String, String),
                                          ByRef dicMOCXD As Dictionary(Of String, clsMOCXD),
                                          ByRef ret_strResultMsg As String) As Boolean

    Try

      For Each PO_ID In dicPO_ID.Values
        Dim str_PO_ID As String() = PO_ID.Split("_")
        Dim XD001 = str_PO_ID(0)  '單別
        Dim XD002 = str_PO_ID(1)  '單號
        Dim XD011 = str_PO_ID(2)  '更新碼(增、刪、修)
        'Dim XD015 = str_PO_ID(3)  '更新時間
        Dim dicMOCXD_Transfer As New Dictionary(Of String, clsMOCXD)
        For Each objMOCXD In dicMOCXD.Values
          If objMOCXD.XD001 <> XD001 Or objMOCXD.XD002 <> XD002 Or objMOCXD.XD011 <> XD011 Then
            Continue For
          End If
          If dicMOCXD_Transfer.ContainsKey(objMOCXD.gid) = False Then
            dicMOCXD_Transfer.Add(objMOCXD.gid, objMOCXD)
          End If
        Next
        SendMessageToLog("單別：" & XD001 & "，單號：" & XD002 & "，動作：" & XD011, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        If dicMOCXD_Transfer.Any Then
          '要變更的資料
          Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
          Dim dicAdd_PO As New Dictionary(Of String, clsPO)
          Dim dicDelete_PO As New Dictionary(Of String, clsPO)
          Dim dicUpdate_PO As New Dictionary(Of String, clsPO)
          Dim dicAdd_PO_Line As New Dictionary(Of String, clsPO_LINE)
          Dim dicDelete_PO_Line As New Dictionary(Of String, clsPO_LINE)
          Dim dicUpdate_PO_Line As New Dictionary(Of String, clsPO_LINE)
          Dim dicAdd_PO_DTL As New Dictionary(Of String, clsPO_DTL)
          Dim dicDelete_PO_DTL As New Dictionary(Of String, clsPO_DTL)
          Dim dicUpdate_PO_DTL As New Dictionary(Of String, clsPO_DTL)

          Dim dicUpdateMOCXD As New Dictionary(Of String, clsMOCXD)
          'Dim PO_ID = ""
          Dim PO_TYPE = ""
          '儲存要更新的SQL，進行一次性更新
          Dim lstSql As New List(Of String)
          Dim lstSql_ERP As New List(Of String)

          '檢查資料
          If Check_Data(dicMOCXD_Transfer, ret_strResultMsg) = False Then
            Return False
          End If
          '進行資料調整
          If Get_UpdateData(dicMOCXD_Transfer, ret_strResultMsg, Host_Command, dicUpdateMOCXD, dicAdd_PO, dicDelete_PO, dicUpdate_PO, dicAdd_PO_Line, dicDelete_PO_Line, dicUpdate_PO_Line, dicAdd_PO_DTL, dicDelete_PO_DTL, dicUpdate_PO_DTL) = False Then
            'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
            Return False
          End If
          '取得SQL
          If Get_SQL(ret_strResultMsg, Host_Command, dicUpdateMOCXD, lstSql, lstSql_ERP) = False Then
            'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
            Return False
          End If
          '執行SQL與更新物件
          If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
            'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
            Return False
          End If
          'SendPurchaserData(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)

        End If
      Next

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Check_Data(ByRef ret_objMOCXD As Dictionary(Of String, clsMOCXD),
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      For Each objMOCXD In ret_objMOCXD.Values
        If objMOCXD.XD001 = "" Then
          ret_strResultMsg = "ERP端 領退料單別為空"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        If objMOCXD.XD002 = "" Then
          ret_strResultMsg = "ERP端 領退料單號為空"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        If objMOCXD.XD005 = "" Then
          ret_strResultMsg = "ERP端 領退料單品號為空"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        If IsNumeric(objMOCXD.XD007) = False Then
          ret_strResultMsg = "ERP端 領退料單數量不為數字"
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
  '新增資料或得到要更新的資料
  Private Function Get_UpdateData(ByRef ret_dicMOCXD As Dictionary(Of String, clsMOCXD),
                                  ByRef ret_strResultMsg As String,
                                  ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                                  ByRef ret_dicUpdateMOCXD As Dictionary(Of String, clsMOCXD),
                                  ByRef ret_dicAdd_PO As Dictionary(Of String, clsPO),
                                  ByRef ret_dicDelete_PO As Dictionary(Of String, clsPO),
                                  ByRef ret_dicUpdate_PO As Dictionary(Of String, clsPO),
                                  ByRef ret_dicAdd_POLine As Dictionary(Of String, clsPO_LINE),
                                  ByRef ret_dicDelete_POLine As Dictionary(Of String, clsPO_LINE),
                                  ByRef ret_dicUpdate_POLine As Dictionary(Of String, clsPO_LINE),
                                  ByRef ret_dicAdd_PO_DTL As Dictionary(Of String, clsPO_DTL),
                                  ByRef ret_dicDelete_PO_DTL As Dictionary(Of String, clsPO_DTL),
                                  ByRef ret_dicUpdate_PO_DTL As Dictionary(Of String, clsPO_DTL)) As Boolean
    Try
      Dim dicAdd_SKU As New Dictionary(Of String, clsSKU)
      'Dim dicUpdate_SKU As New Dictionary(Of String, clsSKU)
      '取得所有的PO單號
      Dim tmp_dicPOID As New Dictionary(Of String, String)
      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      'Dim tmp_dicPO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
      Dim User_ID As String = ""
      'Dim Event_ID As String = objPO_Data.EventID
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim LotManagement = "N"
      'Dim Companyid = objSendWorkData.Companyid
      Dim Create_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim PickingData_TYPE = ""

      For Each objMOCXD In ret_dicMOCXD.Values
        Dim PO_ID As String = objMOCXD.XD001 & "_" & objMOCXD.XD002
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


      Dim bln_POUpdate = False
      If ret_dicMOCXD.First.Value.XD011 = "7" Or
         ret_dicMOCXD.First.Value.XD011 = "9" Then
        bln_POUpdate = True
      End If
      If bln_POUpdate = False Then
        For Each objMOCXD In ret_dicMOCXD.Values
          'Dim obj = objNoticeDataInfo.NoticeDetailDataList.NoticeDetailDataInfo.First


          '以下是填固定對應欄位
          Dim PO_ID = objMOCXD.XD001 & "_" & objMOCXD.XD002  '領/退料單號
          Dim PO_TYPE1 = enuPOType_1.Picking_out
          Dim PO_TYPE2 = enuPOType_2.Material_Out
          Dim PO_TYPE3 = ""
          Dim WO_TYPE = enuWOType.Discharge
          Dim H_PO_ORDER_TYPE = enuOrderType.Material_Out
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
          Dim H_PO_CREATE_TIME = Now_Time
          Dim H_PO_FINISH_TIME = ""
          Dim H_PO_STEP_NO = enuStepNo.Queue
          'Dim H_PO_ORDER_TYPE = enuOrderType.semiSKU_out
          Dim H_PO1 = ""
          Dim H_PO2 = ""
          Dim H_PO3 = ""
          Dim H_PO4 = ""
          Dim H_PO5 = ""
          Dim H_PO6 = ""
          Dim H_PO7 = ""
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
          Dim PO_KEY1 = objMOCXD.XD001
          Dim PO_KEY2 = objMOCXD.XD002
          Dim PO_KEY3 = ""
          Dim PO_KEY4 = ""
          Dim PO_KEY5 = ""

          '調整PO
          If ret_dicAdd_PO.ContainsKey(PO_ID) = False And ret_dicUpdate_PO.ContainsKey(PO_ID) = False Then
            If tmp_dicPO.ContainsKey(PO_ID) = True Then '單據已經存在
              'Dim obj_PO As clsPO = tmp_dicPO.Item(PO_ID)
              ''先檢查PO的狀態還有類型是否正確
              'If obj_PO.PO_Status <> enuPOStatus.Queued Then
              '  ret_strResultMsg = "單據已執行"
              '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              '  Return False
              'End If

              ''更新PO的資料
              'Dim objNewPO As clsPO = tmp_dicPO.Item(PO_ID).Clone
              'objNewPO.PO_Type1 = PO_TYPE1
              'objNewPO.User_ID = User_ID
              'objNewPO.Write_Off_No = PO_ID
              'objNewPO.H_PO_STEP_NO = H_PO_STEP_NO
              'objNewPO.H_PO1 = H_PO1
              'objNewPO.H_PO3 = H_PO3
              'objNewPO.H_PO4 = H_PO4
              'objNewPO.H_PO5 = H_PO5
              'objNewPO.H_PO8 = H_PO8
              'objNewPO.H_PO9 = H_PO9
              'objNewPO.H_PO10 = H_PO10
              'ret_dicUpdate_PO.Add(objNewPO.gid, objNewPO)
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

          'For Each objPURTD In ret_dicPURTD.Values
          'Dim ASRSPart = "N"
          '檢查料品主檔是否存在
          Dim SKU_NO = objMOCXD.XD005.Trim

          Dim dicSKU As New Dictionary(Of String, clsSKU)
          'If gMain.objHandling.O_GetDB_dicSKUBySKUNo(objMOCXD.XD005, dicSKU) = True Then
          If gMain.objHandling.O_GetDB_dicSKUBySKUNo(SKU_NO, dicSKU) = True Then
            If dicSKU.Any Then
              'ASRSPart = dicSKU.First.Value.SKU_COMMON1
            Else
              'ret_strResultMsg = "料品不存在無法取得庫別資訊, PO_ID=" & PO_ID & ", SKU_NO=" & objMOCXD.XD004 & " 請先建立品號資料"
              ret_strResultMsg = "料品不存在無法取得庫別資訊, PO_ID=" & PO_ID & ", SKU_NO=" & SKU_NO & " 請先建立品號資料"
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Continue For
            End If
          End If

          'Dim SKU_NO = objMOCXD.XD005

          Dim PO_LINE_NO = objMOCXD.XD013
          Dim QTY = CDbl(objMOCXD.XD007)

          'If SetQTYByPackeUnit(objPO_DTL.SKU, objPO_DTL.CheckQty, objPO_DTL.Unit, QTY, ret_strResultMsg) = False Then
          '  Continue For
          'End If
          Dim QTY_FINISH = 0
          Dim H_QTY_PROCESS = 0
          Dim H_POL1 = ""
          Dim H_POL2 = ""
          Dim H_POL3 = ""
          Dim H_POL4 = ""
          Dim H_POL5 = ""
          Dim PO_Line_Key = clsPO_LINE.Get_Combination_Key(PO_ID, PO_LINE_NO)
          If tmp_dicPO_Line.ContainsKey(PO_Line_Key) = True Then  '單據已經存在
            'Dim objNewPO_Line = tmp_dicPO_Line.Item(PO_Line_Key).Clone()
            'With objNewPO_Line
            '  .QTY = QTY
            '  .QTY_FINISH = QTY_FINISH
            '  .H_QTY_PROCESS = H_QTY_PROCESS
            '  .H_POL1 = H_POL1
            '  .H_POL2 = H_POL2
            '  .H_POL3 = H_POL3
            '  .H_POL4 = H_POL4
            '  .H_POL5 = H_POL5
            'End With
            'If ret_dicUpdate_POLine.ContainsKey(objNewPO_Line.gid) = False Then
            '  ret_dicUpdate_POLine.Add(objNewPO_Line.gid, objNewPO_Line)
            'End If
          Else
            Dim objNewPO_Line = New clsPO_LINE(PO_ID, PO_LINE_NO, QTY, QTY_FINISH, H_QTY_PROCESS, H_POL1, H_POL2, H_POL3, H_POL4, H_POL5)
            If ret_dicAdd_POLine.ContainsKey(objNewPO_Line.gid) = False Then
              ret_dicAdd_POLine.Add(objNewPO_Line.gid, objNewPO_Line)
            End If
          End If

          Dim PO_SERIAL_NO = objMOCXD.XD013
          Dim WORKING_TYPE = ""
          Dim WORKING_SERIAL_NO = ""
          Dim WORKING_SERIAL_SEQ = ""
          'Dim SKU_NO = ""
          'If LotManagement = "Y" Then
          '  If objPO_DTL.LotId = "" Then
          '    ret_strResultMsg = "此品號:" & SKU_NO & " 需有批號"
          '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          '    Continue For
          '  End If
          'End If

          '客戶信件提到料號中用 "-" 切割，取中間那段文字為客戶編號，若沒有 "-" 代表該料號沒有客戶編號
          Dim customerNo = String.Empty
          Dim strSplit() = SKU_NO.Split("-")

          If strSplit.Length = 3 Then
            customerNo = strSplit(1)
          End If

          Dim LOT_NO = "" 'objPO_DTL.LotId
          'Dim QTY = ""
          Dim QTY_PROCESS = 0
          'Dim QTY_FINISH = 0
          Dim PODTL_STATUS = enuPODTLStatus.Queued
          Dim PACKAGE_ID = EMPTYKey
          Dim ITEM_COMMON1 = EMPTYKey
          Dim ITEM_COMMON2 = EMPTYKey
          Dim ITEM_COMMON3 = EMPTYKey
          Dim ITEM_COMMON4 = EMPTYKey
          Dim ITEM_COMMON5 = EMPTYKey
          Dim ITEM_COMMON6 = EMPTYKey
          Dim ITEM_COMMON7 = EMPTYKey
          Dim ITEM_COMMON8 = EMPTYKey
          Dim ITEM_COMMON9 = EMPTYKey
          Dim ITEM_COMMON10 = EMPTYKey
          Dim SORT_ITEM_COMMON1 = EMPTYKey
          Dim SORT_ITEM_COMMON2 = EMPTYKey
          Dim SORT_ITEM_COMMON3 = EMPTYKey
          Dim SORT_ITEM_COMMON4 = EMPTYKey
          Dim SORT_ITEM_COMMON5 = EMPTYKey
          Dim COMMENTS = ""
          COMMENTS = COMMENTS.Replace("'", "''")
          Dim EXPIRED_DATE = ""
          Dim STORAGE_TYPE = enuStorageType.Store
          Dim BND = enuBND.NB
          Dim QC_STATUS = enuQCStatus.NULL
          'Dim FROM_OWNER_ID = ""
          Dim FROM_OWNER_ID = objMOCXD.XD006
          Dim FROM_SUB_OWNER_ID = ""
          Dim TO_OWNER_ID = "" 'objInDataInfo.Owner
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
          Dim H_POD1 = objMOCXD.XD003 '單據日期
          Dim H_POD2 = objMOCXD.XD004 '廠別
          Dim H_POD3 = objMOCXD.XD006 '領退料單庫別
          Dim H_POD4 = objMOCXD.XD008 '單位
          Dim H_POD5 = objMOCXD.XD009 '製令單別
          Dim H_POD6 = objMOCXD.XD010 '製令單號
          Dim H_POD7 = objMOCXD.XD012 '單據性質
          Dim H_POD8 = ""
          Dim H_POD9 = ""
          Dim H_POD10 = ""
          Dim H_POD11 = ""
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

          Dim PO_DTL_Key = clsPO_DTL.Get_Combination_Key(PO_ID, PO_SERIAL_NO)

          '一般入出庫單據

          If tmp_dicPO_DTL.ContainsKey(PO_DTL_Key) = True Then  '單據存在
            '            Dim objNewPO_DTL = ret_tmp_dicPO_DTL.Item(PO_DTL_Key).Clone()
            '            With objNewPO_DTL
            '              .SKU_NO = SKU_NO
            '              .LOT_NO = LOT_NO
            '              .QTY = QTY
            '              .H_POD1 = H_POD1
            '              .H_POD2 = H_POD2
            '              .H_POD3 = H_POD3
            '              .H_POD4 = H_POD4
            '              .H_POD5 = H_POD5
            '              .H_POD6 = H_POD6
            '              .H_POD7 = H_POD7
            '              .H_POD8 = H_POD8
            '              .H_POD9 = H_POD9
            '              .H_POD10 = H_POD10
            '              .H_POD11 = H_POD11
            '              .H_POD12 = H_POD12
            '              .H_POD13 = H_POD13
            '              .H_POD14 = H_POD14
            '            End With


            '            If ret_dicUpdate_PO_DTL.ContainsKey(objNewPO_DTL.gid) = False Then
            '              ret_dicUpdate_PO_DTL.Add(objNewPO_DTL.gid, objNewPO_DTL)
            '            End If
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

          Dim objUpdateMOCXD As clsMOCXD = objMOCXD.Clone
          objUpdateMOCXD.XD011 = "1"
          ret_dicUpdateMOCXD.Add(objUpdateMOCXD.gid, objUpdateMOCXD)
        Next

#Region "處理MSG並送出執行"
        If ret_dicAdd_PO.Any Then
          If Module_Send_WMSMessage.Send_T5F1U1_POManagement_to_WMS(ret_strResultMsg, ret_dicAdd_PO, ret_dicAdd_POLine, ret_dicAdd_PO_DTL, Host_Command, enuAction.Create.ToString) = False Then
            Return False
          End If
        End If

#End Region
        'Next

      Else
        If tmp_dicPO.Count = 0 Then
          '一次應該只會有一個PO_ID，此MODULE每張單據都會各自呼叫
          For Each PO_ID In tmp_dicPOID.Values
            ret_strResultMsg = $"PO_ID:{PO_ID} 查無 PO 單據"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Continue For
          Next
        End If

        Dim objUpdateMOCXD As clsMOCXD = ret_dicMOCXD.First.Value.Clone 'objMOCXD.Clone

        '改單
        If objUpdateMOCXD.XD011 = "7" Then

          'PO
          Dim objUpdatePO = tmp_dicPO.First.Value.Clone
          If ret_dicUpdate_PO.ContainsKey(objUpdatePO.gid) = False Then
            ret_dicUpdate_PO.Add(objUpdatePO.gid, objUpdatePO)
          End If

          'PO_LINE
          For Each objMOCXD In ret_dicMOCXD.Values

            For Each tmp_objUpdatePO_Line In tmp_dicPO_Line.Values
              If objMOCXD.XD013 = tmp_objUpdatePO_Line.PO_LINE_NO Then
                Dim objUpdatePO_Line = tmp_objUpdatePO_Line.Clone
                objUpdatePO_Line.QTY = objMOCXD.XD007
                If ret_dicUpdate_POLine.ContainsKey(objUpdatePO_Line.gid) = False Then
                  ret_dicUpdate_POLine.Add(objUpdatePO_Line.gid, objUpdatePO_Line)
                End If
                Continue For
              End If
            Next

            'PO_DTL
            For Each tmp_objUpdatePO_DTL In tmp_dicPO_DTL.Values
              If objMOCXD.XD013 = tmp_objUpdatePO_DTL.PO_LINE_NO Then
                Dim objUpdatePO_DTL = tmp_objUpdatePO_DTL.Clone
                objUpdatePO_DTL.QTY = objUpdateMOCXD.XD007
                If ret_dicUpdate_PO_DTL.ContainsKey(objUpdatePO_DTL.gid) = False Then
                  ret_dicUpdate_PO_DTL.Add(objUpdatePO_DTL.gid, objUpdatePO_DTL)
                End If
              End If
            Next
          Next

          objUpdateMOCXD.XD011 = "8"
          ret_dicUpdateMOCXD.Add(objUpdateMOCXD.gid, objUpdateMOCXD)
          If ret_dicUpdate_PO.Any Then
            If Module_Send_WMSMessage.Send_T5F1U1_POManagement_to_WMS(ret_strResultMsg, ret_dicUpdate_PO, ret_dicUpdate_POLine, ret_dicUpdate_PO_DTL, Host_Command, enuAction.Modify.ToString) = False Then
              Return False
            End If
          End If
        ElseIf objUpdateMOCXD.XD011 = "9" Then
          '刪單
          If tmp_dicPO.Any Then
            ret_dicDelete_PO = tmp_dicPO
          End If
          If tmp_dicPO_Line.Any Then
            ret_dicDelete_POLine = tmp_dicPO_Line
          End If
          If tmp_dicPO_DTL.Any Then
            ret_dicDelete_PO_DTL = tmp_dicPO_DTL
          End If
          objUpdateMOCXD.XD011 = "A"
          ret_dicUpdateMOCXD.Add(objUpdateMOCXD.gid, objUpdateMOCXD)
          If ret_dicDelete_PO.Any Then
            If Module_Send_WMSMessage.Send_T5F1U1_POManagement_to_WMS(ret_strResultMsg, ret_dicDelete_PO, ret_dicDelete_POLine, ret_dicDelete_PO_DTL, Host_Command, enuAction.Delete.ToString) = False Then
              Return False
            End If
          End If
        End If


      End If


      'For Each objBuyDataInfo In objPO_Data.BuyDataList.BuyDataInfo
      '  Dim PO_ID As String = objBuyDataInfo.POId
      '  Dim PO_TYPE As String = objBuyDataInfo.POType
      '  SendPurchaserData(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)
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
                           ByRef ret_dicUpdateMOCXD As Dictionary(Of String, clsMOCXD),
                          ByRef lstSql As List(Of String),
                           ByRef lstSql_ERP As List(Of String)) As Boolean
    Try
      '取得Host_Command的SQL
      For Each _Host_COMMAND In Host_Command.Values
        If _Host_COMMAND.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert HOST_T_WMS_Command SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      'ERP的修改更新碼的動作改到等WMS回傳RESULT是成功時才去更新 2023.01.15
      'For Each obj In ret_dicUpdateMOCXD.Values
      '  If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
      '    ret_strResultMsg = "Get Update MOCXD SQL Failed"
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
  Private Function Execute_DataUpdate_ERP(ByRef ret_strResultMsg As String,
                                          ByRef lstSql_ERP As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If ERP_DBManagement.BatchUpdate(lstSql_ERP) = False Then
        '更新DB失敗則回傳False
        ret_strResultMsg = "ERP Update DB Failed"
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

'Module Module_POManagement_HTG_Pickup
'  Public Function O_POManagement_HTG_Pickup(ByVal dicPO_ID As Dictionary(Of String, String),
'                                           ByRef dicMOCXD As Dictionary(Of String, clsMOCXD),
'                                          ByRef ret_strResultMsg As String) As Boolean

'    Try
'      For Each PO_ID In dicPO_ID.Values
'        Dim str_PO_ID As String() = PO_ID.Split("_")
'        Dim XD001 = str_PO_ID(0)
'        Dim XD002 = str_PO_ID(1)
'        Dim dicMOCXD_Transfer As New Dictionary(Of String, clsMOCXD)
'        For Each objMOCXD In dicMOCXD.Values
'          If objMOCXD.XD001 <> XD001 Or objMOCXD.XD002 <> XD002 Then
'            Continue For
'          End If
'          If dicMOCXD_Transfer.ContainsKey(objMOCXD.gid) = False Then
'            dicMOCXD_Transfer.Add(objMOCXD.gid, objMOCXD)
'          End If
'        Next

'        '要變更的資料
'        Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
'        Dim dicAdd_PO As New Dictionary(Of String, clsPO)
'        Dim dicUpdate_PO As New Dictionary(Of String, clsPO)
'        Dim dicAdd_PO_Line As New Dictionary(Of String, clsPO_LINE)
'        Dim dicDelete_PO_Line As New Dictionary(Of String, clsPO_LINE)
'        Dim dicUpdate_PO_Line As New Dictionary(Of String, clsPO_LINE)
'        Dim dicAdd_PO_DTL As New Dictionary(Of String, clsPO_DTL)
'        Dim dicDelete_PO_DTL As New Dictionary(Of String, clsPO_DTL)
'        Dim dicUpdate_PO_DTL As New Dictionary(Of String, clsPO_DTL)
'        Dim dicAdd_PO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
'        Dim dicDelete_PO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
'        Dim dicUpdate_PO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
'        'Dim PO_ID = ""
'        Dim PO_TYPE = ""
'        '儲存要更新的SQL，進行一次性更新
'        Dim lstSql As New List(Of String)

'        '檢查資料
'        If Check_Data(dicMOCXD_Transfer, ret_strResultMsg) = False Then
'          Return False
'        End If
'        '進行資料調整
'        If Get_UpdateData(dicMOCXD_Transfer, ret_strResultMsg, Host_Command, dicAdd_PO, dicUpdate_PO, dicAdd_PO_Line, dicDelete_PO_Line, dicUpdate_PO_Line, dicAdd_PO_DTL, dicDelete_PO_DTL, dicUpdate_PO_DTL, dicAdd_PO_DTL_TRANSACTION, dicDelete_PO_DTL_TRANSACTION, dicUpdate_PO_DTL_TRANSACTION) = False Then
'          'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
'          Return False
'        End If
'        '取得SQL
'        If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
'          'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
'          Return False
'        End If
'        '執行SQL與更新物件
'        If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
'          'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
'          Return False
'        End If
'        'SendPurchaserData(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)
'      Next

'      Return True
'    Catch ex As Exception
'      ret_strResultMsg = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  Private Function Check_Data(ByVal dicMOCXD_Transfer As Dictionary(Of String, clsMOCXD),
'                              ByRef ret_strResultMsg As String) As Boolean
'    Try



'      Return True
'    Catch ex As Exception
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'  '新增資料或得到要更新的資料
'  Private Function Get_UpdateData(ByVal dicMOCXD_Transfer As Dictionary(Of String, clsMOCXD),
'                                  ByRef ret_strResultMsg As String,
'                                  ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
'                                  ByRef ret_dicAdd_PO As Dictionary(Of String, clsPO),
'                                  ByRef ret_dicUpdate_PO As Dictionary(Of String, clsPO),
'                                  ByRef ret_dicAdd_POLine As Dictionary(Of String, clsPO_LINE),
'                                  ByRef ret_dicDelete_POLine As Dictionary(Of String, clsPO_LINE),
'                                  ByRef ret_dicUpdate_POLine As Dictionary(Of String, clsPO_LINE),
'                                  ByRef ret_dicAdd_PO_DTL As Dictionary(Of String, clsPO_DTL),
'                                  ByRef ret_dicDelete_PO_DTL As Dictionary(Of String, clsPO_DTL),
'                                  ByRef ret_dicUpdate_PO_DTL As Dictionary(Of String, clsPO_DTL),
'                                  ByRef ret_dicAdd_PO_DTL_TRANSACTION As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION),
'                                  ByRef ret_dicDelete_PO_DTL_TRANSACTION As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION),
'                                  ByRef ret_dicUpdate_PO_DTL_TRANSACTION As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)) As Boolean
'    Try
'      Dim dicAdd_SKU As New Dictionary(Of String, clsSKU)
'      'Dim dicUpdate_SKU As New Dictionary(Of String, clsSKU)
'      '取得所有的PO單號
'      Dim tmp_dicPOID As New Dictionary(Of String, String)
'      Dim tmp_dicpo As New Dictionary(Of String, clsPO)
'      Dim tmp_dicpo_line As New Dictionary(Of String, clsPO_LINE)
'      Dim tmp_dicpo_dtl As New Dictionary(Of String, clsPO_DTL)
'      'Dim tmp_dicPO_DTL_TRANSACTION As New Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION)
'      Dim User_ID As String = ""
'      'Dim Event_ID As String = objPO_Data.EventID
'      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
'      Dim LotManagement = "N"
'      'Dim Companyid = objSendWorkData.Companyid
'      Dim Create_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
'      Dim PickingData_TYPE = ""
'      For Each objMOCXD In dicMOCXD_Transfer.Values
'        Dim PO_ID As String = objMOCXD.XD001 & "_" & objMOCXD.XD002
'        If tmp_dicPOID.ContainsKey(PO_ID) = False Then
'          tmp_dicPOID.Add(PO_ID, PO_ID)
'        End If
'      Next
'      '使用dicPO取得資料庫裡的PO資料
'      If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(tmp_dicPOID, tmp_dicpo) = False Then
'        ret_strResultMsg = "WMS get PO data From DB Failed"
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If
'      '使用dicPO取得資料庫裡的PO_Line資料
'      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicPOID, tmp_dicpo_line) = False Then
'        ret_strResultMsg = "WMS get PO_Line data From DB Failed"
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If
'      '使用dicPO取得資料庫裡的PO_DTL資料
'      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicPOID, tmp_dicpo_dtl) = False Then
'        ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If

'      Dim bln_POUpdate = True
'      If bln_POUpdate = False Then
'        For Each objMOCXD In dicMOCXD_Transfer.Values
'          'Dim obj = objNoticeDataInfo.NoticeDetailDataList.NoticeDetailDataInfo.First


'          '以下是填固定對應欄位
'          Dim PO_ID = objMOCXD.XD001 & "_" & objMOCXD.XD002  '領/退料單號
'          Dim PO_TYPE1 = enuPOType_1.Combination_in
'          Dim PO_TYPE2 = enuPOType_2.Inbound_Data
'          Dim PO_TYPE3 = ""
'          Dim WO_TYPE = enuWOType.Receipt
'          Dim H_PO_ORDER_TYPE = enuOrderType.Inbound_Data
'          Dim PRIORITY = "50"
'          'Dim CREATE_TIME = ""
'          Dim START_TIME = ""
'          Dim FINISH_TIME = ""
'          'Dim USER_ID = ""
'          Dim CUSTOMER_NO = ""
'          Dim SUPPLIER_NO = ""
'          Dim CLASS_NO = ""
'          Dim SHIPPING_NO = ""
'          Dim WRITE_OFF_NO = ""
'          Dim PO_STATUS = enuPOStatus.Queued
'          Dim AUTO_BOUND = False
'          Dim H_PO_CREATE_TIME = ""
'          Dim H_PO_FINISH_TIME = ""
'          Dim H_PO_STEP_NO = enuStepNo.Queue
'          'Dim H_PO_ORDER_TYPE = enuOrderType.semiSKU_out
'          Dim H_PO1 = ""
'          Dim H_PO2 = ""
'          Dim H_PO3 = ""
'          Dim H_PO4 = ""
'          Dim H_PO5 = ""
'          Dim H_PO6 = ""
'          Dim H_PO7 = ""
'          Dim H_PO8 = ""
'          Dim H_PO9 = ""
'          Dim H_PO10 = ""
'          Dim H_PO11 = ""
'          Dim H_PO12 = ""
'          Dim H_PO13 = ""
'          Dim H_PO14 = ""
'          Dim H_PO15 = ""
'          Dim H_PO16 = ""
'          Dim H_PO17 = ""
'          Dim H_PO18 = ""
'          Dim H_PO19 = ""
'          Dim H_PO20 = ""
'          Dim PO_KEY1 = objMOCXD.XD001
'          Dim PO_KEY2 = objMOCXD.XD002
'          Dim PO_KEY3 = ""
'          Dim PO_KEY4 = ""
'          Dim PO_KEY5 = ""

'          '調整PO
'          If ret_dicAdd_PO.ContainsKey(PO_ID) = False And ret_dicUpdate_PO.ContainsKey(PO_ID) = False Then
'            If tmp_dicpo.ContainsKey(PO_ID) = True Then '單據已經存在
'              'Dim obj_PO As clsPO = tmp_dicPO.Item(PO_ID)
'              ''先檢查PO的狀態還有類型是否正確
'              'If obj_PO.PO_Status <> enuPOStatus.Queued Then
'              '  ret_strResultMsg = "單據已執行"
'              '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'              '  Return False
'              'End If

'              ''更新PO的資料
'              'Dim objNewPO As clsPO = tmp_dicPO.Item(PO_ID).Clone
'              'objNewPO.PO_Type1 = PO_TYPE1
'              'objNewPO.User_ID = User_ID
'              'objNewPO.Write_Off_No = PO_ID
'              'objNewPO.H_PO_STEP_NO = H_PO_STEP_NO
'              'objNewPO.H_PO1 = H_PO1
'              'objNewPO.H_PO3 = H_PO3
'              'objNewPO.H_PO4 = H_PO4
'              'objNewPO.H_PO5 = H_PO5
'              'objNewPO.H_PO8 = H_PO8
'              'objNewPO.H_PO9 = H_PO9
'              'objNewPO.H_PO10 = H_PO10
'              'ret_dicUpdate_PO.Add(objNewPO.gid, objNewPO)
'            Else  '單據不存在
'              '建立新的PO資料
'              Dim objNewPO = New clsPO(PO_ID, PO_TYPE1, PO_TYPE2, PO_TYPE3, WO_TYPE, PRIORITY, Now_Time, START_TIME, FINISH_TIME, User_ID, CUSTOMER_NO, CLASS_NO, SHIPPING_NO,
'                                       PO_STATUS, WRITE_OFF_NO, AUTO_BOUND, H_PO_CREATE_TIME, H_PO_FINISH_TIME, H_PO_STEP_NO, H_PO_ORDER_TYPE,
'                                       H_PO1, H_PO2, H_PO3, H_PO4, H_PO5, H_PO6, H_PO7, H_PO8, H_PO9, H_PO10, H_PO11, H_PO12, H_PO13, H_PO14, H_PO15,
'                                       H_PO16, H_PO17, H_PO18, H_PO19, H_PO20, SUPPLIER_NO, PO_KEY1, PO_KEY2, PO_KEY3, PO_KEY4, PO_KEY5)
'              'Dim objNewPO = New clsPO(PO_ID, PO_TYPE1, PO_TYPE2, "", WO_TYPE, 50, CREATE_TIME, "", "", USER_ID, "", "", "", PO_STATUS, PO_ID, False, CREATE_TIME, "", H_PO_STEP_NO,
'              '                         H_PO_ORDER_TYPE, H_PO1, "", H_PO3, H_PO4, H_PO5, "", "", H_PO8, H_PO9, H_PO10, "", "", "", "", "", "", "", "", "", "")
'              If ret_dicAdd_PO.ContainsKey(objNewPO.gid) = False Then
'                ret_dicAdd_PO.Add(objNewPO.gid, objNewPO)
'              End If
'            End If
'          End If

'          'For Each objPURTD In ret_dicPURTD.Values
'          'Dim ASRSPart = "N"
'          '檢查料品主檔是否存在
'          Dim dicSKU As New Dictionary(Of String, clsSKU)
'          If gMain.objHandling.O_GetDB_dicSKUBySKUNo(objMOCXD.XD005, dicSKU) = True Then
'            If dicSKU.Any Then
'              'ASRSPart = dicSKU.First.Value.SKU_COMMON1
'            Else
'              ret_strResultMsg = "料品不存在無法取得庫別資訊, PO_ID=" & PO_ID & ", SKU_NO=" & objMOCXD.XD005 & " 請先建立品號資料"
'              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'              Return False
'            End If
'          End If

'          Dim SKU_NO = objMOCXD.XD005

'          Dim Serial_ID As Integer = 0
'          Dim PO_LINE_NO = objMOCXD.XD013
'          Dim QTY = CDbl(objMOCXD.XD007)

'          'If SetQTYByPackeUnit(objPO_DTL.SKU, objPO_DTL.CheckQty, objPO_DTL.Unit, QTY, ret_strResultMsg) = False Then
'          '  Return False
'          'End If
'          Dim QTY_FINISH = 0
'          Dim H_QTY_PROCESS = 0
'          Dim H_POL1 = ""
'          Dim H_POL2 = ""
'          Dim H_POL3 = ""
'          Dim H_POL4 = ""
'          Dim H_POL5 = ""
'          Dim PO_Line_Key = clsPO_LINE.Get_Combination_Key(PO_ID, PO_LINE_NO)
'          If tmp_dicpo_line.ContainsKey(PO_Line_Key) = True Then  '單據已經存在
'            'Dim objNewPO_Line = tmp_dicPO_Line.Item(PO_Line_Key).Clone()
'            'With objNewPO_Line
'            '  .QTY = QTY
'            '  .QTY_FINISH = QTY_FINISH
'            '  .H_QTY_PROCESS = H_QTY_PROCESS
'            '  .H_POL1 = H_POL1
'            '  .H_POL2 = H_POL2
'            '  .H_POL3 = H_POL3
'            '  .H_POL4 = H_POL4
'            '  .H_POL5 = H_POL5
'            'End With
'            'If ret_dicUpdate_POLine.ContainsKey(objNewPO_Line.gid) = False Then
'            '  ret_dicUpdate_POLine.Add(objNewPO_Line.gid, objNewPO_Line)
'            'End If
'          Else
'            Dim objNewPO_Line = New clsPO_LINE(PO_ID, PO_LINE_NO, QTY, QTY_FINISH, H_QTY_PROCESS, H_POL1, H_POL2, H_POL3, H_POL4, H_POL5)
'            If ret_dicAdd_POLine.ContainsKey(objNewPO_Line.gid) = False Then
'              ret_dicAdd_POLine.Add(objNewPO_Line.gid, objNewPO_Line)
'            End If
'          End If

'          Dim PO_SERIAL_NO = objMOCXD.XD013
'          Dim WORKING_TYPE = ""
'          Dim WORKING_SERIAL_NO = ""
'          Dim WORKING_SERIAL_SEQ = ""
'          'Dim SKU_NO = ""
'          'If LotManagement = "Y" Then
'          '  If objPO_DTL.LotId = "" Then
'          '    ret_strResultMsg = "此品號:" & SKU_NO & " 需有批號"
'          '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          '    Return False
'          '  End If
'          'End If
'          Dim LOT_NO = "" 'objPO_DTL.LotId
'          'Dim QTY = ""
'          Dim QTY_PROCESS = 0
'          'Dim QTY_FINISH = 0
'          Dim PODTL_STATUS = enuPODTLStatus.Queued
'          Dim PACKAGE_ID = ""
'          Dim ITEM_COMMON1 = ""
'          Dim ITEM_COMMON2 = ""
'          Dim ITEM_COMMON3 = ""
'          Dim ITEM_COMMON4 = ""
'          Dim ITEM_COMMON5 = ""
'          Dim ITEM_COMMON6 = ""
'          Dim ITEM_COMMON7 = ""
'          Dim ITEM_COMMON8 = ""
'          Dim ITEM_COMMON9 = ""
'          Dim ITEM_COMMON10 = ""
'          Dim SORT_ITEM_COMMON1 = ""
'          Dim SORT_ITEM_COMMON2 = ""
'          Dim SORT_ITEM_COMMON3 = ""
'          Dim SORT_ITEM_COMMON4 = ""
'          Dim SORT_ITEM_COMMON5 = ""
'          Dim COMMENTS = ""
'          COMMENTS = COMMENTS.Replace("'", "''")
'          Dim EXPIRED_DATE = ""
'          Dim STORAGE_TYPE = enuStorageType.Store
'          Dim BND = enuBND.None
'          Dim QC_STATUS = enuQCStatus.NULL
'          Dim FROM_OWNER_ID = ""
'          Dim FROM_SUB_OWNER_ID = ""
'          Dim TO_OWNER_ID = "" 'objInDataInfo.Owner
'          Dim TO_SUB_OWNER_ID = ""
'          Dim FACTORY_ID = ""
'          Dim DEST_AREA_ID = ""
'          Dim DEST_LOCATION_ID = ""
'          Dim CLOSE_ABLE = 1
'          Dim H_POD_STEP_NO = enuStepNo.Queue
'          Dim H_POD_MOVE_TYPE = ""
'          Dim H_POD_FINISH_TIME = ""
'          Dim H_POD_BILLING_DATE = ""
'          Dim H_POD_CREATE_TIME = ""
'          Dim H_POD1 = objMOCXD.XD001
'          Dim H_POD2 = objMOCXD.XD002
'          Dim H_POD3 = objMOCXD.XD003
'          Dim H_POD4 = objMOCXD.XD004
'          Dim H_POD5 = objMOCXD.XD005
'          Dim H_POD6 = objMOCXD.XD006
'          Dim H_POD7 = objMOCXD.XD007
'          Dim H_POD8 = objMOCXD.XD008
'          Dim H_POD9 = objMOCXD.XD009
'          Dim H_POD10 = objMOCXD.XD010
'          Dim H_POD11 = objMOCXD.XD011
'          Dim H_POD12 = objMOCXD.XD012
'          Dim H_POD13 = objMOCXD.XD013
'          Dim H_POD14 = ""
'          Dim H_POD15 = ""
'          Dim H_POD16 = ""
'          Dim H_POD17 = ""
'          Dim H_POD18 = ""
'          Dim H_POD19 = ""
'          Dim H_POD20 = ""
'          Dim H_POD21 = ""
'          Dim H_POD22 = ""
'          Dim H_POD23 = ""
'          Dim H_POD24 = ""
'          Dim H_POD25 = ""

'          Dim PO_DTL_Key = clsPO_DTL.Get_Combination_Key(PO_ID, PO_SERIAL_NO)

'          '一般入出庫單據
'          If tmp_dicpo_dtl.ContainsKey(PO_DTL_Key) = True Then  '單據存在
'            '            Dim objNewPO_DTL = ret_tmp_dicPO_DTL.Item(PO_DTL_Key).Clone()
'            '            With objNewPO_DTL
'            '              .SKU_NO = SKU_NO
'            '              .LOT_NO = LOT_NO
'            '              .QTY = QTY
'            '              .H_POD1 = H_POD1
'            '              .H_POD2 = H_POD2
'            '              .H_POD3 = H_POD3
'            '              .H_POD4 = H_POD4
'            '              .H_POD5 = H_POD5
'            '              .H_POD6 = H_POD6
'            '              .H_POD7 = H_POD7
'            '              .H_POD8 = H_POD8
'            '              .H_POD9 = H_POD9
'            '              .H_POD10 = H_POD10
'            '              .H_POD11 = H_POD11
'            '              .H_POD12 = H_POD12
'            '              .H_POD13 = H_POD13
'            '              .H_POD14 = H_POD14
'            '            End With

'            '            If ret_dicUpdate_PO_DTL.ContainsKey(objNewPO_DTL.gid) = False Then
'            '              ret_dicUpdate_PO_DTL.Add(objNewPO_DTL.gid, objNewPO_DTL)
'            '            End If
'          Else
'            Dim objNewPO_DTL = New clsPO_DTL(PO_ID, PO_LINE_NO, PO_SERIAL_NO, WORKING_TYPE, WORKING_SERIAL_NO, WORKING_SERIAL_SEQ, SKU_NO, LOT_NO, QTY, QTY_PROCESS, QTY_FINISH,
'                                             COMMENTS, PACKAGE_ID, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5,
'                                             ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10,
'                                             SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, STORAGE_TYPE, BND, QC_STATUS, FROM_OWNER_ID,
'                                             FROM_SUB_OWNER_ID, TO_OWNER_ID, TO_SUB_OWNER_ID, FACTORY_ID, DEST_AREA_ID, DEST_LOCATION_ID, H_POD_STEP_NO,
'                                             H_POD_MOVE_TYPE, H_POD_FINISH_TIME, H_POD_BILLING_DATE, H_POD_CREATE_TIME,
'                                             H_POD1, H_POD2, H_POD3, H_POD4, H_POD5, H_POD6, H_POD7, H_POD8, H_POD9, H_POD10,
'                                             H_POD11, H_POD12, H_POD13, H_POD14, H_POD15, H_POD16, H_POD17, H_POD18, H_POD19, H_POD20,
'                                             H_POD21, H_POD22, H_POD23, H_POD24, H_POD25, PODTL_STATUS, CLOSE_ABLE)
'            If ret_dicAdd_PO_DTL.ContainsKey(objNewPO_DTL.gid) = False Then
'              ret_dicAdd_PO_DTL.Add(objNewPO_DTL.gid, objNewPO_DTL)
'            End If
'          End If

'          'Next

'#Region "處理MSG並送出執行"
'          If ret_dicAdd_PO.Any Then
'            If Module_Send_WMSMessage.Send_T5F1U1_POManagement_to_WMS(ret_strResultMsg, ret_dicAdd_PO, ret_dicAdd_POLine, ret_dicAdd_PO_DTL, Host_Command, enuAction.Create.ToString) = False Then
'              Return False
'            End If
'          End If

'#End Region

'        Next
'      Else
'        'PO
'        'Dim objUpdatePO = ret_tmp_dicPO.First.Value.Clone
'        'If ret_dicUpdate_PO.ContainsKey(objUpdatePO.gid) = False Then
'        '  ret_dicUpdate_PO.Add(objUpdatePO.gid, objUpdatePO)
'        'End If
'        ''PO_LINE
'        'Dim objUpdatePO_Line = ret_tmp_dicPO_Line.First.Value.Clone
'        'If ret_dicUpdate_POLine.ContainsKey(objUpdatePO_Line.gid) = False Then
'        '  ret_dicUpdate_POLine.Add(objUpdatePO_Line.gid, objUpdatePO_Line)
'        'End If
'        ''PO_DTL
'        'Dim objUpdatePO_DTL = ret_tmp_dicPO_DTL.First.Value.Clone
'        'objUpdatePO_DTL.QTY = ret_dicPURTF.First.Value.TF009
'        'If ret_dicUpdate_PO_DTL.ContainsKey(objUpdatePO_DTL.gid) = False Then
'        '  ret_dicUpdate_PO_DTL.Add(objUpdatePO_DTL.gid, objUpdatePO_DTL)
'        'End If

'        If ret_dicUpdate_PO.Any Then
'          If Module_Send_WMSMessage.Send_T5F1U1_POManagement_to_WMS(ret_strResultMsg, ret_dicUpdate_PO, ret_dicUpdate_POLine, ret_dicUpdate_PO_DTL, Host_Command, enuAction.Modify.ToString) = False Then
'            Return False
'          End If
'        End If
'      End If


'      'For Each objBuyDataInfo In objPO_Data.BuyDataList.BuyDataInfo
'      '  Dim PO_ID As String = objBuyDataInfo.POId
'      '  Dim PO_TYPE As String = objBuyDataInfo.POType
'      '  SendPurchaserData(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)
'      'Next

'      'ret_dicDelete_POLine = tmp_dicPO_Line
'      'ret_dicDelete_PO_DTL = tmp_dicPO_DTL
'      Return True
'    Catch ex As Exception
'      ret_strResultMsg = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'  'SQL
'  Private Function Get_SQL(ByRef ret_strResultMsg As String,
'                           ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
'                          ByRef lstSql As List(Of String)) As Boolean
'    Try
'      '取得Host_Command的SQL
'      For Each _Host_COMMAND In Host_Command.Values
'        If _Host_COMMAND.O_Add_Insert_SQLString(lstSql) = False Then
'          ret_strResultMsg = "Get Insert HOST_T_WMS_Command SQL Failed"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next

'      Return True
'    Catch ex As Exception
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'  '執行刪除和新增的SQL語句，並進行記憶體資料更新
'  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
'                                      ByRef lstSql As List(Of String)) As Boolean
'    Try
'      '更新所有的SQL
'      If Common_DBManagement.BatchUpdate(lstSql) = False Then
'        '更新DB失敗則回傳False
'        ret_strResultMsg = "WMS Update DB Failed"
'        Return False
'      End If
'      Return True
'    Catch ex As Exception
'      ret_strResultMsg = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'End Module
