'20200327
'V1.0.0
'Vito

'執行PO單 送PO to WO 給WMS

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T6F5U2_ItemLabelPrint

  Public Function O_T6F5U2_ItemLabelPrint(ByVal Receive_Msg As MSG_T6F5U2_ItemLabelPrint,
                                        ByRef ret_strResultMsg As String,
                                        ByRef ret_Wait_UUID As String) As Boolean
    Try
      ''儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '要變更的資料
      Dim dic_Item_Label As New Dictionary(Of String, clsItemLabel)
      Dim dic_Print_Item_Label As New Dictionary(Of String, clsItemLabel)
      'Dim dic_PO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If

      '進行資料處理
      If Get_Data(Receive_Msg, ret_strResultMsg, dic_Item_Label, dic_Print_Item_Label, Host_Command, ret_Wait_UUID) = False Then
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_T6F5U2_ItemLabelPrint,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      For Each objItemLabelInfo In Receive_Msg.Body.ItemLabelList.ItemLabelInfo
        '資料檢查
        Dim ITEM_LABEL_ID As String = objItemLabelInfo.ITEM_LABEL_ID
        Dim ITEM_LABEL_TYPE As String = objItemLabelInfo.ITEM_LABEL_TYPE
        Dim PO_ID As String = objItemLabelInfo.PO_ID

        '檢查ITEM_LABEL_ID是否為空
        'If ITEM_LABEL_ID = "" Then
        '  ret_strResultMsg = "ITEM_LABEL_ID is empty"
        '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  Return False
        'End If
        '檢查ITEM_LABEL_TYPE是否為空
        If ITEM_LABEL_TYPE = "" Then
          ret_strResultMsg = "ITEM_LABEL_TYPE is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        Else
          If ITEM_LABEL_TYPE <> "1" And ITEM_LABEL_TYPE <> "2" And ITEM_LABEL_TYPE <> "3" Then
            ret_strResultMsg = "ITEM_LABEL_TYPE 未定義"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        End If
        '檢查PO_ID是否為空
        If PO_ID = "" Then
          ret_strResultMsg = "PO_ID is empty"
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
  Private Function Get_Data(ByVal Receive_Msg As MSG_T6F5U2_ItemLabelPrint,
                            ByRef ret_strResultMsg As String,
                            ByRef dic_Item_Label As Dictionary(Of String, clsItemLabel),
                            ByRef dic_Print_Item_Label As Dictionary(Of String, clsItemLabel),
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                            ByRef ret_Wait_UUID As String) As Boolean
    Try
      Dim Action = Receive_Msg.Body.Action
      Dim User_ID = Receive_Msg.Header.ClientInfo.UserID
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      '先進行資料邏輯檢查
      For Each objItemLabelInfo In Receive_Msg.Body.ItemLabelList.ItemLabelInfo
        Dim GUID As Guid = Guid.NewGuid
        Dim ITEM_LABEL_ID = GUID.ToString          'QR Code
        Dim ITEM_LABEL_TYPE = objItemLabelInfo.ITEM_LABEL_TYPE
        Dim PO_ID = objItemLabelInfo.PO_ID
        Dim TAG1 = objItemLabelInfo.TAG1                        '快速輸入碼
        Dim TAG2 = objItemLabelInfo.TAG2
        Dim TAG3 = objItemLabelInfo.TAG3
        Dim TAG4 = objItemLabelInfo.TAG4
        Dim TAG5 = objItemLabelInfo.TAG5
        Dim TAG6 = objItemLabelInfo.TAG6
        Dim TAG7 = objItemLabelInfo.TAG7
        Dim TAG8 = objItemLabelInfo.TAG8                        '出口批次號
        Dim TAG9 = objItemLabelInfo.TAG9                        '供應商代碼
        Dim TAG10 = objItemLabelInfo.TAG10                       '料號
        Dim TAG11 = objItemLabelInfo.TAG11                       '內部訂單號前6碼
        Dim TAG12 = objItemLabelInfo.TAG12                       'Bom表PK
        Dim TAG13 = objItemLabelInfo.TAG13                       '面料長度
        Dim TAG14 = objItemLabelInfo.TAG14                       '淨重
        Dim TAG15 = objItemLabelInfo.TAG15                       '毛重
        Dim TAG16 = objItemLabelInfo.TAG16                       '包裝尺寸：深
        Dim TAG17 = objItemLabelInfo.TAG17                       '包裝尺寸：寬
        Dim TAG18 = objItemLabelInfo.TAG18                       '包裝尺寸：高
        Dim TAG19 = objItemLabelInfo.TAG19                       '捲號
        Dim TAG20 = objItemLabelInfo.TAG20                       '缸號
        Dim DESC = ""
        Dim COLOR = ""
        Dim COMPN = ""
        'Dim ShellSpec As String() = objItemLabel.ShellSpec.Split(vbCrLf)
        Dim TAG21 = objItemLabelInfo.TAG21                       '正常狀態為 N，Y:表示被供應商標示刪除，參考用，避免供應商印出標籤貼上後，誤刪資料。
        Dim TAG22 = objItemLabelInfo.TAG22                       '面料品項規格
        Dim TAG23 = objItemLabelInfo.TAG23
        Dim TAG24 = objItemLabelInfo.TAG24
        Dim TAG25 = objItemLabelInfo.TAG25
        Dim TAG26 = objItemLabelInfo.TAG26
        Dim TAG27 = objItemLabelInfo.TAG27
        Dim TAG28 = objItemLabelInfo.TAG28
        Dim TAG29 = objItemLabelInfo.TAG29
        Dim TAG30 = objItemLabelInfo.TAG30
        Dim TAG31 = objItemLabelInfo.TAG31
        Dim TAG32 = objItemLabelInfo.TAG32
        Dim TAG33 = objItemLabelInfo.TAG33
        Dim TAG34 = objItemLabelInfo.TAG34
        Dim TAG35 = objItemLabelInfo.TAG35
        Dim PRINTED = "1"                         '建完會馬上印
        Dim CREATE_USER = "Other"
        Dim FIRST_PRINT_TIME = ""
        Dim LAST_PRINT_TIME = ""
        Dim UPDATE_TIME = ""
        Dim CREATE_TIME = Now_Time

        Dim objNewItemLabel = New clsItemLabel(ITEM_LABEL_ID, ITEM_LABEL_TYPE, PO_ID, TAG1, TAG2, TAG3, TAG4, TAG5, TAG6, TAG7, TAG8, TAG9,
                                   TAG10, TAG11, TAG12, TAG13, TAG14, TAG15, TAG16, TAG17, TAG18, TAG19, TAG20, TAG21, TAG22, TAG23, TAG24, TAG25,
                                   TAG26, TAG27, TAG28, TAG29, TAG30, TAG31, TAG32, TAG33, TAG34, TAG35, PRINTED, CREATE_USER, FIRST_PRINT_TIME, LAST_PRINT_TIME, UPDATE_TIME, CREATE_USER)

        If dic_Item_Label.ContainsKey(objNewItemLabel.gid) = False Then
          dic_Item_Label.Add(objNewItemLabel.gid, objNewItemLabel)
        End If

        If dic_Print_Item_Label.ContainsKey(objNewItemLabel.gid) = False Then
          dic_Print_Item_Label.Add(objNewItemLabel.gid, objNewItemLabel)
        End If
      Next


      If dic_Item_Label.Any Then
        '列印標籤前先建立標籤
        If Action = "PRINT" Then
          If Module_Send_WMSMessage.Send_T6F5U1_ItemLabelManagement_to_WMS(ret_strResultMsg, dic_Item_Label, Host_Command, enuAction.Create.ToString, ret_Wait_UUID) = False Then
            Return False
          End If
          'If Module_Send_WMSMessage.Send_T6F5U2_ItemLabelPrint_to_WMS(ret_strResultMsg, dic_Print_Item_Label, Host_Command, "PRINT") = False Then
          '  Return False
          'End If
        End If
        End If


      '取得流水號
      'Dim dicUUID As New Dictionary(Of String, clsUUID)
      'If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
      '  ret_strResultMsg = "Get UUID False"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      'If dicUUID.Any = False Then
      '  ret_strResultMsg = "Get UUID False"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      'Dim objUUID = dicUUID.Values(0)


      '寫入Command給WMS
      'If Send_Command_to_WMS(ret_strResultMsg, dic_PO_DTL, objUUID, Host_Command, ret_Wait_UUID, User_ID) = False Then
      '  Return False
      'End If




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

End Module
