'20190319
'V1.0.0
'Jerry

'狀態:Open
'回報ERP已放行 通知WMS執行盤點單

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T11F1U12_StocktakingExecution
  Public Function O_Process_Message(ByVal Receive_Msg As MSG_T11F1U12_StocktakingExecution,
                                    ByRef ret_strResultMsg As String,
                                    ByRef ret_Wait_UUID As String) As Boolean
    'Try
    '  '儲存要更新的SQL，進行一次性更新
    '  Dim lstSql As New List(Of String)
    '  '要變更的資料
    '  Dim Host_Command As clsFromHostCommand = Nothing

    '  '先進行資料邏輯檢查
    '  If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
    '    Return False
    '  End If
    '  '取得要新增的Carrier和Carrier_Status的資料
    '  If Get_Data(Receive_Msg, ret_strResultMsg, Host_Command) = False Then
    '    Return False
    '  End If
    '  '取得要更新到DB的SQL
    '  If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
    '    Return False
    '  End If
    '  '執行資料更新
    '  If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
    '    Return False
    '  End If
    '  Return True
    'Catch ex As Exception
    '  ret_strResultMsg = ex.ToString
    '  SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    '  Return False
    'End Try
    Return True
  End Function

  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T11F1U12_StocktakingExecution,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      For Each objStocktakingInfo In Receive_Msg.Body.StocktakingList.StocktakingInfo
        '資料檢查
        Dim STOCKTAKING_ID As String = objStocktakingInfo.STOCKTAKING_ID
        If STOCKTAKING_ID = "" Then
          ret_strResultMsg = "H_PO_ORDER_TYPE is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查 STOCKTAKING_TYPE3 是否正確
        Dim STOCKTAKING_TYPE3 As String = objStocktakingInfo.STOCKTAKING_TYPE3
        If STOCKTAKING_TYPE3 = "" Then
          ret_strResultMsg = "STOCKTAKING_TYPE3 is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        ElseIf CheckValueInEnum(Of enuStockTaking_Type3)(STOCKTAKING_TYPE3) = False Then
          ret_strResultMsg = "STOCKTAKING_TYPE3 is not defined"
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

  Private Function Get_Data(ByVal Receive_Msg As MSG_T11F1U12_StocktakingExecution,
                            ByRef ret_strResultMsg As String,
                            ByRef Host_Command As clsFromHostCommand) As Boolean
    'Try
    '  Dim User_ID = Receive_Msg.Header.ClientInfo.UserID
    '  Dim tmp_dicStocktacking_id As New Dictionary(Of String, String)
    '  '先進行資料邏輯檢查
    '  For Each objStocktakingInfo In Receive_Msg.Body.StocktakingList.StocktakingInfo
    '    '資料檢查
    '    Dim STOCKTAKING_ID As String = objStocktakingInfo.STOCKTAKING_ID
    '    Dim STOCKTAKING_TYPE1 As String = objStocktakingInfo.STOCKTAKING_TYPE1
    '    Dim STOCKTAKING_TYPE2 As String = objStocktakingInfo.STOCKTAKING_TYPE2
    '    Dim STOCKTAKING_TYPE3 As String = objStocktakingInfo.STOCKTAKING_TYPE3
    '    Dim dicStocktakingID As New Dictionary(Of String, String)
    '    If dicStocktakingID.ContainsKey(STOCKTAKING_ID) = False Then
    '      dicStocktakingID.Add(STOCKTAKING_ID, STOCKTAKING_ID)
    '    End If
    '    If tmp_dicStocktacking_id.ContainsKey(STOCKTAKING_ID) = False Then
    '      tmp_dicStocktacking_id.Add(STOCKTAKING_ID, STOCKTAKING_ID)
    '    End If

    '    '檢查盤點單是否存在
    '    Dim dicStocktaking As New Dictionary(Of String, clsTSTOCKTAKING)
    '    If gMain.objHandling.O_GetDB_dicStocktakingBydicStocktakingID(dicStocktakingID, dicStocktaking) = False Then
    '      ret_strResultMsg = "無法取得盤點單資訊"
    '      SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
    '      Return False
    '    End If
    '    If dicStocktaking.Any = False Then
    '      ret_strResultMsg = "不存在盤點單，單號:" & STOCKTAKING_ID
    '      SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
    '      Return False
    '    End If
    '    Dim objStocktaking As clsStocktaking = dicStocktaking.First.Value
    '    If objStocktaking.STATUS <> enuSTOCKTAKING_STATUS.Queued Then
    '      ret_strResultMsg = "盤點單單號:" & STOCKTAKING_ID & " 不為未執行"
    '      SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
    '      Return False
    '    End If
    '    If objStocktaking.STOCKTAKING_TYPE3 <> enuStockTaking_Type3.ERP Then
    '      ret_strResultMsg = "盤點單單號:" & STOCKTAKING_ID & " 不屬於ERP單據，請確認事件接收方。"
    '      SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
    '      Return False
    '    End If
    '    '通知ERP已放行
    '    INVTE(STOCKTAKING_ID)

    '  Next

    '  '取得流水號
    '  Dim dicUUID As New Dictionary(Of String, clsUUID)
    '  If gMain.objHandling.O_GetDB_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
    '    ret_strResultMsg = "Get UUID False"
    '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
    '    Return False
    '  End If
    '  If dicUUID.Any = False Then
    '    ret_strResultMsg = "Get UUID False"
    '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
    '    Return False
    '  End If
    '  Dim objUUID = dicUUID.Values(0)


    '  '寫入Command給WMS
    '  If Send_Command_to_WMS(ret_strResultMsg, tmp_dicStocktacking_id, Host_Command, objUUID, User_ID) = False Then
    '    Return False
    '  End If
    '  Return True
    'Catch ex As Exception
    '  ret_strResultMsg = ex.ToString
    '  SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    '  Return False
    'End Try
    Return True
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
                                       ByVal tmp_dicStocktacking_id As Dictionary(Of String, String),
                                       ByRef Host_Command As clsFromHostCommand, ByRef objUUID As clsUUID,
                                       ByVal User_ID As String) As Boolean
    Try
      Dim UUID = objUUID.Get_NewUUID
      '將單據發並送給WMS 
      Dim dicStockTakingExecute As New MSG_T10F2U2_StocktakingExecute
      dicStockTakingExecute.Header = New clsHeader
      'ret_Wait_UUID = UUID
      dicStockTakingExecute.Header.UUID = UUID
      dicStockTakingExecute.Header.EventID = "T10F2U2_StocktakingExecute"
      dicStockTakingExecute.Header.Direction = "Primary"

      dicStockTakingExecute.Header.ClientInfo = New clsHeader.clsClientInfo
      dicStockTakingExecute.Header.ClientInfo.ClientID = "Handler"
      dicStockTakingExecute.Header.ClientInfo.UserID = User_ID
      dicStockTakingExecute.Header.ClientInfo.IP = ""
      dicStockTakingExecute.Header.ClientInfo.MachineID = ""

      dicStockTakingExecute.Body = New MSG_T10F2U2_StocktakingExecute.clsBody

      For Each StocktakingID In tmp_dicStocktacking_id.Values
        Dim StocktakingInfo As New MSG_T10F2U2_StocktakingExecute.clsBody.clsStocktakingList.clsStocktakingInfo
        StocktakingInfo.STOCKTAKING_ID = StocktakingID
        dicStockTakingExecute.Body.StocktakingList.StocktakingInfo.Add(StocktakingInfo)
      Next


      'dicStockTakingExecute.Body.Action = "Create"
      'dicStockTakingExecute.Body.AutoFlag = "1"

      'Dim PO_ID = ""
      'Dim POList As New MSG_T5F3U23_POToWO.clsBody.clsPOList
      'For Each PO_DTL In dicUpdate_PO_DTL.Values
      '  Dim lstPOInfo As New MSG_T5F3U23_POToWO.clsBody.clsPOList.clsPOInfo
      '  PO_ID = PO_DTL.PO_ID
      '  lstPOInfo.PO_ID = PO_DTL.PO_ID
      '  lstPOInfo.PO_SERIAL_NO = PO_DTL.PO_Serial_No
      '  lstPOInfo.QTY = PO_DTL.QTY
      '  lstPOInfo.SOURCE_AREA_NO = SOURCE_AREA_NO '"MOSA01"       '暫時先寫死
      '  lstPOInfo.SOURCE_LOCATION_NO = SOURCE_LOCATION_NO '"MOSA_MOSA01_MGV01"   '暫時先寫死
      '  POList.POInfo.Add(lstPOInfo)
      'Next

      'Dim WO_Info As New MSG_T5F3U23_POToWO.clsBody.clsWOInfo
      'WO_Info.WO_ID = "" 'PO_ID
      'WO_Info.SHIPPING_NO = ""

      'dicStockTakingExecute.Body.WOInfo = WO_Info
      'dicStockTakingExecute.Body.POList = POList '資料填寫完成

      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_T10F2U2_StocktakingExecute(strXML, dicStockTakingExecute, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(MSG_T10F2U2_StocktakingExecute)"
        End If
        Return False
      End If

      '寫Command 
      Host_Command = New clsFromHostCommand(UUID, enuSystemType.HostHandler, enuSystemType.WMS, "T10F2U2_StocktakingExecute", 1, "", "", "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "", "")
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '盤點單放行回報
  Private Sub INVTE(ByVal WorkID As String)
    'Try
    '  Dim STDIN As New MSG_SendTransferDataToERP
    '  STDIN.ProdID = "WMS"
    '  STDIN.Companyid = "MOSA_TEST_LINDA"
    '  STDIN.Userid = "DS"
    '  STDIN.DoAction = "2"
    '  STDIN.Docase = "1"

    '  'Data
    '  Dim Data As New STD_INData
    '  Dim FormHead As New STD_INDataFormHead
    '  FormHead.TableName = "INVTE"

    '  '組成Header
    '  Dim Rocord_Head(0) As STD_INDataFormHeadRecordList
    '  Dim Rocord_Head_Info As New STD_INDataFormHeadRecordList
    '  '<TableName>INVTE</TableName>
    '  '<RecordList>
    '  '  <TE001>盤點底稿編號</TE001>
    '  '  <TE200>WMS接收成功</TE200>                   
    '  '</RecordList>  
    '  Rocord_Head_Info.TE001 = WorkID
    '  Rocord_Head_Info.TE200 = "-1" ' -1:已放行 0:未放行 1:錯誤
    '  Rocord_Head(0) = Rocord_Head_Info
    '  FormHead.RecordList() = Rocord_Head
    '  Data.FormHead = FormHead
    '  STDIN.Result = "success"
    '  STDIN.Data = Data
    '  SendMessageToLog("通知ERP盤點單放行，單號：" & WorkID, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

    '  STD_IN(STDIN)
    'Catch ex As Exception
    '  MsgBox(ex.ToString)
    'End Try
  End Sub

End Module
