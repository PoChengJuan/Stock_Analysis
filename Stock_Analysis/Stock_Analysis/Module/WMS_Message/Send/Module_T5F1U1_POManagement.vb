'20190117
'V1.0.0
'Mark
'WMS回覆訂單轉工單的結果
'狀態:Open

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T5F1U1_POManagement
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~發送~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~接收回傳的結果~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  Public Function O_CheckMessageResult(ByVal Receive_Msg As MSG_T5F1U1_PO_Management,
                                       ByRef ret_strResultMsg As String,
                                       ByVal strRejectReason As String,
                                       ByRef ret_Wait_UUID As String,
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
      If Get_Data(Receive_Msg, ret_strResultMsg, blnResult, Host_Command, ret_Wait_UUID, strRejectReason) = False Then
        Return False
      End If
      '取得要更新到DB的SQL
      If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
        Return False
      End If

#Region "將膠塊的部份另外自動執行"

      ''將要送給WMS的COMMAND先清除
      'Host_Command.Clear()
      'lstSql.Clear()
      ''膠塊需要另外處理
      'If dicGluePO_DTL.Any AndAlso Send_T5F2U62_AutoInbound_to_WMS(ret_strResultMsg, Host_Command, dicGluePO_DTL) = False Then
      '  Return False
      'End If
      ''取得SQL
      'If Get_SQL_Host_Command(ret_strResultMsg, Host_Command, lstSql) = False Then

      '  Return False
      'End If
      ''執行SQL與更新物件
      'If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then

      '  Return False
      'End If
#End Region

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T5F1U1_PO_Management,
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
  Private Function Get_Data(ByVal Receive_Msg As MSG_T5F1U1_PO_Management,
                            ByRef ret_strResultMsg As String,
                            ByRef blnResult As Boolean,
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                            ByRef ret_Wait_UUID As String,
                            ByVal strRejectReason As String) As Boolean
    Try
      Dim Hist_UUID As String = GetNewTime_ByDataTimeFormat(DBFullTimeUUIDFormat)
      Dim UUID As String = Receive_Msg.Header.UUID
      Dim Now_Time As String = GetNewTime_DBFormat()
      '取出所有PO_ID
      Dim tmp_dicPO_ID As New Dictionary(Of String, String)
      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      'For Each POInfo In Receive_Msg.Body.POInfo.PO_ID
      Dim PO_ID As String = Receive_Msg.Body.POInfo.PO_ID
      If tmp_dicPO_ID.ContainsKey(PO_ID) = False Then
        tmp_dicPO_ID.Add(PO_ID, PO_ID)
      End If
      'Next
      '抓取資料庫PO的資料
      'Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(tmp_dicPO_ID, tmp_dicPO) = True Then
        For Each objPO As clsPO In tmp_dicPO.Values
#Region "111"
          '資料檢查
          'Dim IN_PO_ID As String = objPOInfo.PO_ID
          Dim H_PO_ORDER_TYPE As String = objPO.H_PO_ORDER_TYPE
          Dim User_ID = objPO.User_ID
          PO_ID = objPO.PO_ID
          If ExcutePO.ContainsKey(PO_ID) = False Then ExcutePO.Add(PO_ID, PO_ID) '排除時間差問題 20190628

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

          'If tmp_dicPO.First.Value.PO_Type2 <> enuPOType_2.transaction_in Then
          '  Return True
          'End If

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
          'For Each objPO_DTL In tmp_dicPO_DTL
          '  ' QTY += objPO_DTL.Value.QTY
          '  If dic_PO_DTL.ContainsKey(objPO_DTL.Key) = False Then
          '    dic_PO_DTL.Add(objPO_DTL.Key, objPO_DTL.Value)
          '  End If
          'Next

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

#End Region


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
          If Send_Command_to_WMS(ret_strResultMsg, tmp_dicPO_DTL, objUUID, Host_Command, ret_Wait_UUID, User_ID) = False Then
            Return False
          End If


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
                           ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得 SQL
      For Each obj As clsFromHostCommand In Host_Command.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert HostHandler Command SQL Failed"
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

  '
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
