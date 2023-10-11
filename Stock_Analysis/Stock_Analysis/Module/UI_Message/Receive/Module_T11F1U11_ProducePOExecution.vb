'20190115
'V1.0.0
'Mark

'狀態:Open
'執行生產線工單

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T11F1U11_ProducePOExecution
  Public Function O_Process_Message(ByVal Receive_Msg As MSG_T11F1U11_ProducePOExecution,
                                    ByRef ret_strResultMsg As String,
                                    ByRef ret_Wait_UUID As String) As Boolean
    Try
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)
      '要變更的資料
      Dim dic_PO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim Host_Command As clsFromHostCommand = Nothing

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '取得要新增的Carrier和Carrier_Status的資料
      If Get_Data(Receive_Msg, ret_strResultMsg, dic_PO_DTL, Host_Command, ret_Wait_UUID) = False Then
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
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T11F1U11_ProducePOExecution,
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

  Private Function Get_Data(ByVal Receive_Msg As MSG_T11F1U11_ProducePOExecution,
                            ByRef ret_strResultMsg As String,
                            ByRef dic_PO_DTL As Dictionary(Of String, clsPO_DTL),
                            ByRef Host_Command As clsFromHostCommand,
                            ByRef ret_Wait_UUID As String) As Boolean
    Try
      Dim User_ID = Receive_Msg.Header.ClientInfo.UserID
      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      Dim tmp_dicPO_Line As New Dictionary(Of String, clsPO_LINE)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim SOURCE_AREA_NO As String = ""
      Dim SOURCE_LOCATION_NO As String = ""
      '先進行資料邏輯檢查
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        Dim COMMENTS As String = objPOInfo.COMMENTS
        SOURCE_AREA_NO = objPOInfo.SOURCE_AREA_NO
        SOURCE_LOCATION_NO = objPOInfo.SOURCE_LOCATION_NO
        '取得單據
        If gMain.objHandling.O_GetDB_dicPOByPOID_OrderType(PO_ID, H_PO_ORDER_TYPE, tmp_dicPO) = False Then
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
        For Each objPO_DTL In tmp_dicPO_DTL
          If dic_PO_DTL.ContainsKey(objPO_DTL.Key) = False Then
            dic_PO_DTL.Add(objPO_DTL.Key, objPO_DTL.Value)
          End If
        Next
      Next

      '取得流水號
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_GetDB_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
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


      '寫入Command給WMS
      If Send_Command_to_WMS(ret_strResultMsg, dic_PO_DTL, objUUID, Host_Command, ret_Wait_UUID, User_ID, SOURCE_AREA_NO, SOURCE_LOCATION_NO) = False Then
        Return False
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
                           ByRef Host_Command As clsFromHostCommand,
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得要送給WMS的CMD
      If Host_Command.O_Add_Insert_SQLString(lstSql) = False Then
        ret_strResultMsg = "Get Insert HOST_T_COMMAND SQL Failed"
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

  Private Function Send_Command_to_WMS(ByRef Result_Message As String,
                                       ByVal dicUpdate_PO_DTL As Dictionary(Of String, clsPO_DTL), ByRef objUUID As clsUUID,
                                       ByRef Host_Command As clsFromHostCommand,
                                       ByRef ret_Wait_UUID As String,
                                       ByVal User_ID As String,
                                       ByVal SOURCE_AREA_NO As String,
                                       ByVal SOURCE_LOCATION_NO As String) As Boolean
    Try
      Dim UUID = objUUID.Get_NewUUID
      '將單據發並送給WMS 取得回復為OK後才將單據更新
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
        lstPOInfo.PO_SERIAL_NO = PO_DTL.PO_Serial_No
        lstPOInfo.QTY = PO_DTL.QTY
        lstPOInfo.SOURCE_AREA_NO = SOURCE_AREA_NO '"MOSA01"       '暫時先寫死
        lstPOInfo.SOURCE_LOCATION_NO = SOURCE_LOCATION_NO '"MOSA_MOSA01_MGV01"   '暫時先寫死
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
      Host_Command = New clsFromHostCommand(UUID, enuSystemType.HostHandler, enuSystemType.WMS, "T5F3U23_POToWO", 1, "", "", "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "", "")
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
