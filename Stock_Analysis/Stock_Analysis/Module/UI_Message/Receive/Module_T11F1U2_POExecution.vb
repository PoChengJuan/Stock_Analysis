'20180629
'V1.0.0
'Jerry

'執行PO單 送PO to WO 給WMS

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T11F1U2_POExecution

  Public Function O_T11F1U2_POExecution(ByVal Receive_Msg As MSG_T11F1U2_POExecution,
                                        ByRef ret_strResultMsg As String,
                                        ByRef ret_Wait_UUID As String) As Boolean
    Try
      ''儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '要變更的資料
      Dim dic_PO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim dic_Item_Label As New Dictionary(Of String, clsItemLabel)
      Dim dicUpdate_Item_Label As New Dictionary(Of String, clsItemLabel)

      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If

      '進行資料處理
      If Get_Data(Receive_Msg, ret_strResultMsg, dic_PO_DTL, dic_Item_Label, dicUpdate_Item_Label, Host_Command, ret_Wait_UUID) = False Then
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_T11F1U2_POExecution,
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
        '檢查PO_Type1是否正確
        If H_PO_ORDER_TYPE = "" Then
          ret_strResultMsg = "H_PO_ORDER_TYPE is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        ElseIf ModuleHelpFunc.CheckValueInEnum(Of enuOrderType)(H_PO_ORDER_TYPE) = False Then
          ret_strResultMsg = "H_PO_ORDER_TYPE 不存在于定义中"
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


  '資料處理
  Private Function Get_Data(ByVal Receive_Msg As MSG_T11F1U2_POExecution,
                            ByRef ret_strResultMsg As String,
                            ByRef dic_PO_DTL As Dictionary(Of String, clsPO_DTL),
                            ByRef ret_dic_Item_Label As Dictionary(Of String, clsItemLabel),
                            ByRef ret_dicUpdate_Item_Label As Dictionary(Of String, clsItemLabel),
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                            ByRef ret_Wait_UUID As String) As Boolean
    Try
      Dim User_ID = Receive_Msg.Header.ClientInfo.UserID
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim Now_Date As String = ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyMMdd")

      Dim dicUUID As New Dictionary(Of String, clsUUID)
      Dim objUUID As clsUUID = Nothing
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
        If gMain.objHandling.O_Get_dicPOByPO_ID_ORDER_TYPE(PO_ID, H_PO_ORDER_TYPE, tmp_dicPO) = False Then
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


        For Each objPO_DTL In tmp_dicPO_DTL.Values
          ' QTY += objPO_DTL.Value.QTY
          If dic_PO_DTL.ContainsKey(objPO_DTL.gid) = False Then
            dic_PO_DTL.Add(objPO_DTL.gid, objPO_DTL)
            If tmp_dicPO.First.Value.PO_Type2 = enuPOType_2.Inbound_Data Then
#Region "建立標籤資訊"
              'For Each objItemLabelInfo In Receive_Msg.Body.ItemLabelList.ItemLabelInfo
              Dim SKU_NO = objPO_DTL.SKU_NO
              Dim dicSKU As New Dictionary(Of String, clsSKU)
              Dim objSKU As clsSKU = Nothing
              gMain.objHandling.O_GetDB_dicSKUBySKUNo(SKU_NO, dicSKU)
              If dicSKU.Any = False Then
                ret_strResultMsg = "無法取得料號：" & SKU_NO
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If
              objSKU = dicSKU.First.Value.Clone

              If objSKU.SKU_TYPE1 <> 1 Then
                ret_strResultMsg = "入庫單據料品需為原料"
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If
              Dim dicSKUPackeStructure As New Dictionary(Of String, clsMSKUPackeStructure)
              Dim objSKUPackeStructure As clsMSKUPackeStructure
              gMain.objHandling.O_GetDB_dicSKUPackeStructureBySKU_NO(objSKU.SKU_NO, dicSKUPackeStructure)
              If dicSKUPackeStructure.Any Then
                objSKUPackeStructure = dicSKUPackeStructure.First.Value
              Else
                ret_strResultMsg = "無法取得料號：" & SKU_NO & "的包裝結構"
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If

              If gMain.objHandling.O_Get_UUID(enuUUID_No.PACKAGE_ID.ToString, dicUUID) = False Then
                ret_strResultMsg = "Get UUID False"
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If
              If dicUUID.Any = False Then
                ret_strResultMsg = "Get UUID False"
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If
              objUUID = dicUUID.Values(0)

              Dim Serial_No_Cnt = 0
              If CDec(objPO_DTL.QTY) >= CDec(objSKUPackeStructure.PACKE_WEIGHT) Then
                Serial_No_Cnt = CDbl(CDec(objPO_DTL.QTY) / CDec(objSKUPackeStructure.PACKE_WEIGHT))
                Serial_No_Cnt = Math.Ceiling(Serial_No_Cnt)
              Else
                ret_strResultMsg = "料號：" & SKU_NO & "的包裝結構設定錯誤，單據：" & objPO_DTL.PO_ID & "，項次：" & objPO_DTL.PO_SERIAL_NO & "，單據重量：" & objPO_DTL.QTY & "，包裝結構淨重：" & objSKUPackeStructure.PACKE_WEIGHT
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If



              For Serial_no As Integer = 1 To Serial_No_Cnt
                Dim SEQ = objUUID.Get_NewUUID
                If objSKU.SKU_TYPE2 = "1" Then
                  '膠塊
                  I_Get_dicItemLabel(ret_strResultMsg, SEQ, objSKU.SKU_TYPE2, Now_Date, objPO_DTL, objSKUPackeStructure.PACKE_WEIGHT, objSKU, ret_dic_Item_Label)
                  I_Get_dicItemLabelByPiece(ret_strResultMsg, SEQ, Now_Date, objPO_DTL, objSKU, ret_dic_Item_Label)
                Else
                  '非膠塊

                  I_Get_dicItemLabel(ret_strResultMsg, SEQ, objSKU.SKU_TYPE2, Now_Date, objPO_DTL, objSKUPackeStructure.PACKE_WEIGHT, objSKU, ret_dic_Item_Label)

                End If
              Next

#End Region
            ElseIf tmp_dicPO.First.Value.PO_Type2 = "" Then 'enuPOType_2.transaction_in Then
              Dim Package_ID = objPO_DTL.PACKAGE_ID
              Dim dicItemLabel As New Dictionary(Of String, clsItemLabel)
              gMain.objHandling.O_GetDB_dicItemLabelByPackage_ID(Package_ID, dicItemLabel)
              If dicItemLabel.Any Then
                Dim objUpdateItemLabel = dicItemLabel.First.Value
                objUpdateItemLabel.PO_ID = PO_ID
                If ret_dicUpdate_Item_Label.ContainsKey(objUpdateItemLabel.gid) = False Then
                  ret_dicUpdate_Item_Label.Add(objUpdateItemLabel.gid, objUpdateItemLabel)
                End If
              Else
                Dim str = "單號：" & PO_ID & ",取不到對應標籤，ITEM_LABEL_ID：" & Package_ID
                SendMessageToLog(str, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If

            End If
            'Next


          End If
        Next
      Next

      If ret_dicUpdate_Item_Label.Any Then
        If Module_Send_WMSMessage.Send_T6F5U1_ItemLabelManagement_to_WMS(ret_strResultMsg, ret_dicUpdate_Item_Label, Host_Command, enuAction.Modify.ToString) = False Then
          Return False
        End If
      End If
      If ret_dic_Item_Label.Any Then
        If Module_Send_WMSMessage.Send_T6F5U1_ItemLabelManagement_to_WMS(ret_strResultMsg, ret_dic_Item_Label, Host_Command, enuAction.Create.ToString) = False Then
          Return False
        End If
      End If

      '取得流水號
      dicUUID = New Dictionary(Of String, clsUUID)
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
      objUUID = dicUUID.Values(0)


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

  'SQL
  Private Function _Get_SQL(ByRef Result_Message As String,
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                            ByRef lstSql As List(Of String)) As Boolean
    Try
      ''取得要刪除的PO_DTL SQL
      'For Each obj In dicUpdate_PO_DTL.Values
      '  If obj.O_Add_Update_SQLString(lstSql) = False Then
      '    Result_Message = "Get Update PO_DTL SQL Failed"
      '    SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
      ''取得要刪除的PO_Line SQL
      'For Each obj In dicUpdate_PO_Line.Values
      '  If obj.O_Add_Update_SQLString(lstSql) = False Then
      '    Result_Message = "Get Update PO_LINE SQL Failed"
      '    SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
      ''取得要Delete的PO SQL
      'For Each obj In dicUpdate_PO.Values
      '  If obj.O_Add_Update_SQLString(lstSql) = False Then
      '    Result_Message = "Get Update PO SQL Failed"
      '    SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '    Return False
      '  End If
      'Next
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
                                       ByRef ret_Wait_UUID As String, ByVal User_ID As String) As Boolean
    Try
      Dim UUID = objUUID.Get_NewUUID

      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim dicPOtoWO As New MSG_T5F3U23_POToWO
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

      dicPOtoWO.Body = New MSG_T5F3U23_POToWO.clsBody
      dicPOtoWO.Body.Action = "Create"
      dicPOtoWO.Body.AutoFlag = "1"


      Dim PO_ID = ""
      Dim POList As New MSG_T5F3U23_POToWO.clsBody.clsPOList
      For Each PO_DTL In dicUpdate_PO_DTL.Values
        Dim lstPOInfo As New MSG_T5F3U23_POToWO.clsBody.clsPOList.clsPOInfo
        PO_ID = PO_DTL.PO_ID
        lstPOInfo.PO_ID = PO_DTL.PO_ID
        lstPOInfo.PO_SERIAL_NO = PO_DTL.PO_SERIAL_NO
        lstPOInfo.QTY = PO_DTL.QTY
        POList.POInfo.Add(lstPOInfo)
      Next

      Dim WO_Info As New MSG_T5F3U23_POToWO.clsBody.clsWOInfo
      WO_Info.WO_ID = "" 'PO_ID
      WO_Info.SHIPPING_NO = ""

      dicPOtoWO.Body.WOInfo = WO_Info
      dicPOtoWO.Body.POList = POList '資料填寫完成

      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_T5F3U3_POToWO(strXML, dicPOtoWO, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(MSG_T5F3U3_POToWO)"
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
  Private Function I_Get_dicItemLabel(ByRef Result_Message As String,
                                      ByVal SEQ As String,
                                      ByVal Type As String,
                                      ByVal Now_Date As String,
                                      ByRef objPO_DTL As clsPO_DTL,
                                      ByVal Weight As String,
                                      ByVal objSKU As clsSKU,
                                      ByRef ret_dic_Item_Label As Dictionary(Of String, clsItemLabel)) As Boolean
    Try
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim Now_Date_yyyyMMdd As String = ModuleHelpFunc.GetNewDate_DBFormat_yyyyMMdd()

      Dim dicSKUPackeStructure As New Dictionary(Of String, clsMSKUPackeStructure)
      Dim objSKUPackeStructure As clsMSKUPackeStructure = Nothing
      gMain.objHandling.O_GetDB_dicSKUPackeStructureBySKU_NO(objPO_DTL.SKU_NO, dicSKUPackeStructure)

      'If dicSKUPackeStructure.Any = False Then
      '  Result_Message = "查無料品：" & objPO_DTL.SKU_NO & "之包裝結構。"
      '  SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '  Return False
      'Else
      '  objSKUPackeStructure = dicSKUPackeStructure.First.Value.Clone
      'End If

      Dim ITEM_LABEL_ID = I_Get_Label_ID(SEQ, "")
      Dim ITEM_LABEL_TYPE = "3"

      If Type = "1" Then
        ITEM_LABEL_TYPE = "1"
      ElseIf Type = "5" Then
        ITEM_LABEL_TYPE = "4"
      End If
      Dim PO_ID = objPO_DTL.PO_ID
      Dim TAG1 = objPO_DTL.PO_SERIAL_NO                        '快速輸入碼
      Dim TAG2 = objSKU.SKU_NO
      Dim str_Date = Now_Date_yyyyMMdd
      Dim TAG3 = str_Date.Replace("/", "") 'objPO_DTL.LOT_NO
      Dim TAG4 = Weight 'objPO_DTL.QTY
      Dim TAG5 = objPO_DTL.ITEM_COMMON3
      Dim TAG6 = CDec(CInt(Weight) * 1000).ToString
      'If objSKU.SKU_TYPE2 = "1" Then
      '  TAG6 = Weight
      'Else
      '  TAG6 = objSKU.SKU_WEIGHT
      'End If
      Dim TAG7 = DateAdd(DateInterval.Day, CInt(objSKU.SAVE_DAYS), CDate(str_Date)).ToString(DBDate_IDFormat_TaiYin) 'Now_Date_yyyyMMdd + objSKU.SAVE_DAYS
      Dim TAG8 = I_Get_Label_Data(ITEM_LABEL_ID, objSKU.SKU_ID1, Now_Date, objSKU.SKU_WEIGHT, TAG7, "")
      Dim TAG9 = ""
      Dim TAG10 = objSKU.SKU_ID1
      Dim TAG11 = objSKU.SKU_ALIS1
      Dim TAG12 = ""
      Dim TAG13 = ""
      Dim TAG14 = ""
      Dim TAG15 = ""
      Dim TAG16 = ""
      Dim TAG17 = ""
      Dim TAG18 = ""
      Dim TAG19 = ""
      Dim TAG20 = ""
      Dim TAG21 = ""                       '正常狀態為 N，Y:表示被供應商標示刪除，參考用，避免供應商印出標籤貼上後，誤刪資料。
      Dim TAG22 = ""                       '面料品項規格
      Dim TAG23 = ""
      Dim TAG24 = ""
      Dim TAG25 = ""
      Dim TAG26 = ""
      Dim TAG27 = ""
      Dim TAG28 = ""
      Dim TAG29 = ""
      Dim TAG30 = ""
      Dim TAG31 = ""
      Dim TAG32 = ""
      Dim TAG33 = ""
      Dim TAG34 = ""
      Dim TAG35 = ""
      Dim PRINTED = "0"
      Dim CREATE_USER = "WMS"
      Dim FIRST_PRINT_TIME = ""
      Dim LAST_PRINT_TIME = ""
      Dim UPDATE_TIME = ""
      Dim CREATE_TIME = Now_Time

      Dim objNewItemLabel = New clsItemLabel(ITEM_LABEL_ID, ITEM_LABEL_TYPE, PO_ID, TAG1, TAG2, TAG3, TAG4, TAG5, TAG6, TAG7, TAG8, TAG9,
                                   TAG10, TAG11, TAG12, TAG13, TAG14, TAG15, TAG16, TAG17, TAG18, TAG19, TAG20, TAG21, TAG22, TAG23, TAG24, TAG25,
                                   TAG26, TAG27, TAG28, TAG29, TAG30, TAG31, TAG32, TAG33, TAG34, TAG35, PRINTED, CREATE_USER, FIRST_PRINT_TIME, LAST_PRINT_TIME, UPDATE_TIME, CREATE_USER)

      If ret_dic_Item_Label.ContainsKey(objNewItemLabel.gid) = False Then
        ret_dic_Item_Label.Add(objNewItemLabel.gid, objNewItemLabel)
      End If

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function I_Get_dicItemLabelByPiece(ByRef Result_Message As String, ByVal SEQ As String, ByVal Now_Date As String, ByRef objPO_DTL As clsPO_DTL, ByVal objSKU As clsSKU, ByRef ret_dic_Item_Label As Dictionary(Of String, clsItemLabel)) As Boolean
    Try
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim Now_Date_yyyyMMdd As String = ModuleHelpFunc.GetNewDate_DBFormat_yyyyMMdd()

      Dim dicSKUPackeStructure As New Dictionary(Of String, clsMSKUPackeStructure)
      Dim objSKUPackeStructure As clsMSKUPackeStructure = Nothing
      gMain.objHandling.O_GetDB_dicSKUPackeStructureBySKU_NO(objPO_DTL.SKU_NO, dicSKUPackeStructure)

      If dicSKUPackeStructure.Any = False Then
        Result_Message = "查無料品：" & objPO_DTL.SKU_NO & "之包裝結構。"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      Else
        objSKUPackeStructure = dicSKUPackeStructure.First.Value.Clone
      End If
      Dim Serial_No_Cnt = CDbl(CDec(objSKUPackeStructure.PACKE_WEIGHT) / CDec(objSKU.SKU_WEIGHT))
      Serial_No_Cnt = Math.Ceiling(Serial_No_Cnt)

      For Serial_no As Integer = 1 To Serial_No_Cnt
        Dim ITEM_LABEL_ID = I_Get_Label_ID(SEQ, Serial_no)          'QR Code
        Dim ITEM_LABEL_TYPE = "2"
        Dim PO_ID = objPO_DTL.PO_ID
        Dim TAG1 = objPO_DTL.PO_SERIAL_NO
        Dim TAG2 = objSKU.SKU_NO
        Dim str_Date = Now_Date_yyyyMMdd
        Dim TAG3 = str_Date.Replace("/", "")
        Dim TAG4 = objSKU.SKU_WEIGHT 'objPO_DTL.QTY
        Dim TAG5 = objPO_DTL.ITEM_COMMON3
        Dim TAG6 = CDec(CInt(objSKU.SKU_WEIGHT) * 1000).ToString 'objSKU.SKU_WEIGHT
        Dim TAG7 = DateAdd(DateInterval.Day, CInt(objSKU.SAVE_DAYS), CDate(str_Date)).ToString(DBDate_IDFormat_TaiYin)
        Dim TAG8 = I_Get_Label_Data(ITEM_LABEL_ID, objSKU.SKU_ID1, Now_Date, objSKU.SKU_WEIGHT, TAG7, Serial_no)
        Dim TAG9 = SEQ
        Dim TAG10 = objSKU.SKU_ID1
        Dim TAG11 = objSKU.SKU_ALIS1      'Alan_20210917
        Dim TAG12 = ""
        Dim TAG13 = ""
        Dim TAG14 = ""
        Dim TAG15 = ""
        Dim TAG16 = ""
        Dim TAG17 = ""
        Dim TAG18 = ""
        Dim TAG19 = ""
        Dim TAG20 = ""
        Dim DESC = ""
        Dim COLOR = ""
        Dim COMPN = ""
        Dim TAG21 = ""                       '正常狀態為 N，Y:表示被供應商標示刪除，參考用，避免供應商印出標籤貼上後，誤刪資料。
        Dim TAG22 = ""                       '面料品項規格
        Dim TAG23 = ""
        Dim TAG24 = ""
        Dim TAG25 = ""
        Dim TAG26 = ""
        Dim TAG27 = ""
        Dim TAG28 = ""
        Dim TAG29 = ""
        Dim TAG30 = ""
        Dim TAG31 = ""
        Dim TAG32 = ""
        Dim TAG33 = ""
        Dim TAG34 = ""
        Dim TAG35 = ""
        Dim PRINTED = "0"
        Dim CREATE_USER = "WMS"
        Dim FIRST_PRINT_TIME = ""
        Dim LAST_PRINT_TIME = ""
        Dim UPDATE_TIME = ""
        Dim CREATE_TIME = Now_Time

        Dim objNewItemLabel = New clsItemLabel(ITEM_LABEL_ID, ITEM_LABEL_TYPE, PO_ID, TAG1, TAG2, TAG3, TAG4, TAG5, TAG6, TAG7, TAG8, TAG9,
                                     TAG10, TAG11, TAG12, TAG13, TAG14, TAG15, TAG16, TAG17, TAG18, TAG19, TAG20, TAG21, TAG22, TAG23, TAG24, TAG25,
                                     TAG26, TAG27, TAG28, TAG29, TAG30, TAG31, TAG32, TAG33, TAG34, TAG35, PRINTED, CREATE_USER, FIRST_PRINT_TIME, LAST_PRINT_TIME, UPDATE_TIME, CREATE_USER)

        If ret_dic_Item_Label.ContainsKey(objNewItemLabel.gid) = False Then
          ret_dic_Item_Label.Add(objNewItemLabel.gid, objNewItemLabel)
        End If
      Next


      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function I_Get_Label_Data(ByVal SEQ As String,
                                  ByVal SKU_NO As String,
                                  ByVal strDate As String,
                                  ByVal Weight As String,
                                  ByVal Valied_Date As String,
                                  ByVal Serial_No As String) As String
    Try
      Dim Label_DATA = ""
      Weight = (CInt(Weight) * 1000).ToString
      If Serial_No = "" Then
        Label_DATA = SEQ & "@" & SKU_NO & "@" & strDate & "@" & Weight & "@" & Valied_Date
      Else
        Label_DATA = SEQ & "@" & SKU_NO & "@" & strDate & "@" & Weight & "@" & Valied_Date & "_" & Serial_No
      End If
      Return Label_DATA
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function I_Get_Label_ID(ByVal SEQ As String,
                                  ByVal Serial_No As String) As String
    Try
      Dim Label_ID = ""
      If Serial_No = "" Then
        Label_ID = SEQ
      Else
        Label_ID = SEQ & "_" & Serial_No.PadLeft(3, "0")
      End If
      Return Label_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
