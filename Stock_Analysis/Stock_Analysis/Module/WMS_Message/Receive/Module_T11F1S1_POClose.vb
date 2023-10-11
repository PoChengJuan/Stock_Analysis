'20180629
'V1.0.0
'Jerry

'结单

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T11F1S1_POClose
  Public Function O_T11F1S1_POClose(ByVal Receive_Msg As MSG_T11F1S1_POClose,
                                          ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim dicAddPO_Posting As New Dictionary(Of String, clsPO_POSTING)
      Dim dicUpdatePO_Posting As New Dictionary(Of String, clsPO_POSTING)
      Dim dicDeletePO_Posting As New Dictionary(Of String, clsPO_POSTING)
      '儲存要更新的SQL， 進行一次性更新
      Dim lstSql As New List(Of String)
      '儲存要更新的SQL， 進行一次性更新
      Dim lstQueueSql As New List(Of String)

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料處理
      If Process_Data(Receive_Msg, dicAddPO_Posting, dicUpdatePO_Posting, dicDeletePO_Posting, ret_strResultMsg) = False Then
        Return False
      End If
      If Get_SQL(ret_strResultMsg, dicAddPO_Posting, dicUpdatePO_Posting, dicDeletePO_Posting, lstSql, lstQueueSql) = False Then
        Return False
      End If
      If Execute_DataUpdate(ret_strResultMsg, lstSql, lstQueueSql) = False Then
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_T11F1S1_POClose,
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
        ''檢查PO_Type1是否正確
        'If H_PO_ORDER_TYPE = "" Then
        '  ret_strResultMsg = "H_PO_ORDER_TYPE is empty"
        '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  Return False
        'ElseIf ModuleHelpFunc.CheckValueInEnum(Of enuOrderType)(H_PO_ORDER_TYPE) = False Then
        '  ret_strResultMsg = "H_PO_ORDER_TYPE 不存在于定义中"
        '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  Return False
        'End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '資料處理
  Private Function Process_Data(ByVal Receive_Msg As MSG_T11F1S1_POClose,
                                ByRef dicAddPO_Posting As Dictionary(Of String, clsPO_POSTING),
                                ByRef dicUpdatePO_Posting As Dictionary(Of String, clsPO_POSTING),
                                ByRef dicDeletePO_Posting As Dictionary(Of String, clsPO_POSTING),
                                ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim USER_ID = Receive_Msg.Header.ClientInfo.UserID
      Dim UUID = Receive_Msg.Header.UUID
      Dim Forced_Close = Receive_Msg.Body.Forced_Close '强制结单
      If Receive_Msg.Body.POList Is Nothing Or Receive_Msg.Body.POList.POInfo.Count = 0 Then
        SendMessageToLog("WMS 给的结单资讯有缺(POList", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return True
        'ret_strResultMsg = "WMS 给的结单资讯有缺(POList)，无法结单。"
        'Return False
      End If

      Dim dicWO_ID As New Dictionary(Of String, String)
      Dim dicPO_Posting As New Dictionary(Of String, clsPO_POSTING)
      Dim dicPO_ID As New Dictionary(Of String, String)
      Dim dicPO As New Dictionary(Of String, clsPO)
      Dim dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      Dim dicPO_LINE As New Dictionary(Of String, clsPO_LINE)
      '找出對應过帳的紀錄
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        If dicPO_ID.ContainsKey(PO_ID) = False Then
          dicPO_ID.Add(PO_ID, PO_ID)
        End If
        For Each objPO_DTLInfo In objPOInfo.PO_DTLList.PO_DTLInfo
          Dim WO_ID As String = objPO_DTLInfo.WO_ID
          If dicWO_ID.ContainsKey(WO_ID) = False Then
            dicWO_ID.Add(WO_ID, WO_ID)
          End If
        Next
      Next
      If gMain.objHandling.O_Get_dicPO_POSTING_By_dicWO_ID(dicWO_ID, dicPO_Posting) = False Then
        ret_strResultMsg = "Get PO Posting Failed."
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(dicPO_ID, dicPO) = False Then
        ret_strResultMsg = "Get PO Failed."
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(dicPO_ID, dicPO_DTL) = False Then
        ret_strResultMsg = "Get PO DTL Failed."
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(dicPO_ID, dicPO_LINE) = False Then
        ret_strResultMsg = "Get PO LINE Failed."
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If


      '開始對帳 如果沒有資料則建立 如果已有責檢查狀態(已過則無須再過(工單、訂單、訂單項次))
      '1. 檢查資訊 先不過帳 進行dic的建立
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        For Each objPO_DTLInfo In objPOInfo.PO_DTLList.PO_DTLInfo
          Dim WO_ID As String = objPO_DTLInfo.WO_ID
          Dim PO_LINE_NO As String = objPO_DTLInfo.PO_LINE_NO
          Dim PO_SERIAL_NO As String = objPO_DTLInfo.PO_SERIAL_NO
          Dim SKU_NO As String = objPO_DTLInfo.SKU_NO
          Dim LOT_NO As String = objPO_DTLInfo.LOT_NO
          Dim ITEM_COMMON1 As String = objPO_DTLInfo.ITEM_COMMON1
          Dim ITEM_COMMON2 As String = objPO_DTLInfo.ITEM_COMMON2
          Dim ITEM_COMMON3 As String = objPO_DTLInfo.ITEM_COMMON3
          Dim ITEM_COMMON4 As String = objPO_DTLInfo.ITEM_COMMON4
          Dim ITEM_COMMON5 As String = objPO_DTLInfo.ITEM_COMMON5
          Dim ITEM_COMMON6 As String = objPO_DTLInfo.ITEM_COMMON6
          Dim ITEM_COMMON7 As String = objPO_DTLInfo.ITEM_COMMON7
          Dim ITEM_COMMON8 As String = objPO_DTLInfo.ITEM_COMMON8
          Dim ITEM_COMMON9 As String = objPO_DTLInfo.ITEM_COMMON9
          Dim ITEM_COMMON10 As String = objPO_DTLInfo.ITEM_COMMON10
          Dim SORT_ITEM_COMMON1 As String = objPO_DTLInfo.SORT_ITEM_COMMON1
          Dim SORT_ITEM_COMMON2 As String = objPO_DTLInfo.SORT_ITEM_COMMON2
          Dim SORT_ITEM_COMMON3 As String = objPO_DTLInfo.SORT_ITEM_COMMON3
          Dim SORT_ITEM_COMMON4 As String = objPO_DTLInfo.SORT_ITEM_COMMON4
          Dim SORT_ITEM_COMMON5 As String = objPO_DTLInfo.SORT_ITEM_COMMON5
          Dim QTY As String = objPO_DTLInfo.QTY

          '取得各自對應的單據資訊
          Dim objPO As clsPO = Nothing
          Dim objPO_DTL As clsPO_DTL = Nothing
          Dim objPO_LINE As clsPO_LINE = Nothing
          If dicPO.TryGetValue(clsPO.Get_Combination_Key(PO_ID), objPO) = False Then
            ret_strResultMsg = "無法取得單據資訊。PO_ID：" & PO_ID
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If dicPO_DTL.TryGetValue(clsPO_DTL.Get_Combination_Key(PO_ID, PO_SERIAL_NO), objPO_DTL) = False Then
            ret_strResultMsg = "無法取得單據資訊。PO_ID：" & PO_ID & " ,PO_SERIAL_NO：" & PO_SERIAL_NO
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If dicPO_LINE.TryGetValue(clsPO_LINE.Get_Combination_Key(PO_ID, PO_LINE_NO), objPO_LINE) = False Then
            ret_strResultMsg = "無法取得單據資訊。PO_ID：" & PO_ID & " ,PO_LINE_NO：" & PO_LINE_NO
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If

          'If objPO.WO_Type = enuWOType.Transform Then Continue For
          If objPO.PO_Type1 = enuPOType_1.Combination_in Then

          End If
          Dim RESULT = enuPOSTING_RESULT.unPOSTING
          Dim RESULT_MESSAGE = ""
          '如果是強制過帳或無須过帳的單據則狀態修改為完成
          If Forced_Close = 2 Then
            RESULT = enuPOSTING_RESULT.Complete
            RESULT_MESSAGE = "強制過帳不上報。"
          Else
            '如果非強制過帳 則根據單據類型處理
            Select Case H_PO_ORDER_TYPE '環鴻都要過帳
              Case enuOrderType.Inbound_Data '
                SendMessageToLog("单据类型为 ERP入庫", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

              Case enuOrderType.Material_Out
                SendMessageToLog("单据类型为 ERP出庫", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

              Case enuOrderType.m_material_in
                SendMessageToLog("单据类型为 WMS手動入庫單，无需向上位系统过帐", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

              Case enuOrderType.m_material_out
                SendMessageToLog("单据类型为 WMS手動出庫單，无需向上位系统过帐", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

            End Select
            If objPO.WO_Type <> enuWOType.Receipt Then
              RESULT = enuPOSTING_RESULT.Complete
              RESULT_MESSAGE = "不是入庫不上報"
            End If
          End If

          '檢查過帳數量


          Dim H_POP1 = 0 '環鴻紀錄過帳次數
          Dim H_POP2 = ""
          Dim H_POP3 = ""
          Dim H_POP4 = ""
          Dim H_POP5 = ""
          Dim CLOSE_USER_ID = USER_ID
          Dim START_TRANSFER_TIME = ""
          Dim FINISH_TRANSFER_TIME = ""
          Dim ORDER_TYPE = H_PO_ORDER_TYPE
          Dim TKNUM = objPO.H_PO16
          Dim OWNER = IIf(objPO.WO_Type = enuWOType.Discharge, objPO_DTL.FROM_OWNER_ID, objPO_DTL.TO_OWNER_ID)
          Dim SUBOWNER = IIf(objPO.WO_Type = enuWOType.Discharge, objPO_DTL.FROM_SUB_OWNER_ID, objPO_DTL.TO_SUB_OWNER_ID)
          Dim KEY_NO = ""

          Dim objNewPO_Posting As New clsPO_POSTING(PO_ID, PO_LINE_NO, WO_ID, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, QTY, UUID, Now_Time, Now_Time, RESULT, RESULT_MESSAGE, H_POP1, H_POP2, H_POP3, H_POP4, H_POP5, SKU_NO, CLOSE_USER_ID, START_TRANSFER_TIME, FINISH_TRANSFER_TIME, ORDER_TYPE, PO_SERIAL_NO, TKNUM, LOT_NO, OWNER, SUBOWNER)
          Dim objExistPO_Posting As clsPO_POSTING = Nothing
          If dicPO_Posting.TryGetValue(objNewPO_Posting.gid, objExistPO_Posting) = False Then
            '不存在的進行新增
            Dim objTmp_Add_PO_Posting As clsPO_POSTING = Nothing
            If dicAddPO_Posting.TryGetValue(objNewPO_Posting.gid, objTmp_Add_PO_Posting) = True Then
              '如果新增的dic裡有存在，則數量往上加
              objTmp_Add_PO_Posting.QTY += objNewPO_Posting.QTY
            Else
              dicAddPO_Posting.Add(objNewPO_Posting.gid, objNewPO_Posting)
            End If
          Else
            '過帳資訊已存在 '檢查是否需再過帳
            If objExistPO_Posting.RESULT <> enuPOSTING_RESULT.Complete Then
              '如果狀態不是完成 則加入更新的dic
              If dicUpdatePO_Posting.ContainsKey(objExistPO_Posting.gid) = True Then
                '不可能發生 除非人員手動操作資料庫!!!!
                ret_strResultMsg = "T_Posting Exist same data ,Please Check. git=" & objExistPO_Posting.gid
                SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If
              dicUpdatePO_Posting.Add(objExistPO_Posting.gid, objExistPO_Posting)
              If RESULT = enuPOSTING_RESULT.Complete Then '如果單據改為完成 (強制過帳)
                objExistPO_Posting.RESULT = RESULT
              End If
            End If
          End If
        Next
      Next

      'add、Update兩者 進行過帳
      SendMessageToLog("開始進行過帳", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Each objPO_Posting In dicAddPO_Posting.Values
        If objPO_Posting.RESULT <> enuPOSTING_RESULT.Complete Then
          '過帳
        End If
        '環鴻是建立資料後排程去過帳的..
      Next
      For Each objPO_Posting In dicUpdatePO_Posting.Values
        If objPO_Posting.RESULT <> enuPOSTING_RESULT.Complete Then
          '過帳
        End If
        '環鴻是建立資料後排程去過帳的..
      Next
      SendMessageToLog("检查过帐是否全数通过", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      '檢查是否全部過帳完成
      Dim bln_POSTING_Success As Boolean = True '判斷是否全數過帳成功
      Dim str_Posting_Msg As String = "" '過帳返回的資訊
      For Each objPO_Posting In dicAddPO_Posting.Values
        If objPO_Posting.RESULT <> enuPOSTING_RESULT.Complete Then
          '存在過帳失敗
          bln_POSTING_Success = False
        End If
        str_Posting_Msg += "工單：" & objPO_Posting.WO_ID & " ,訂單：" & objPO_Posting.PO_ID & " ,訂單項次：" & objPO_Posting.PO_SERIAL_NO & " ,過帳結果：" & objPO_Posting.RESULT_MESSAGE & ";"
      Next
      For Each objPO_Posting In dicUpdatePO_Posting.Values
        If objPO_Posting.RESULT <> enuPOSTING_RESULT.Complete Then
          '存在過帳失敗
          bln_POSTING_Success = False
        End If
        str_Posting_Msg += "工單：" & objPO_Posting.WO_ID & " ,訂單：" & objPO_Posting.PO_ID & " ,訂單項次：" & objPO_Posting.PO_SERIAL_NO & " ,過帳結果：" & objPO_Posting.RESULT_MESSAGE & ";"
      Next
      If str_Posting_Msg.Length > 2500 Then
        str_Posting_Msg = str_Posting_Msg.Substring(0, 2500) & "..."
      End If

      '客製的部分註解掉
      ''回覆WMS '環鴻客製
      'bln_POSTING_Success = True
      'If bln_POSTING_Success Then
      '  ret_strResultMsg = "成功!!;"
      '  Return True
      'End If

      '正常情況
      If bln_POSTING_Success = True Then
        ret_strResultMsg = str_Posting_Msg
        Return True
      Else
        ret_strResultMsg = "失败!! 尚未全数过账成功;" & str_Posting_Msg
        Return False
      End If

      Return False '不可能到這邊
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByVal dicAddPO_Posting As Dictionary(Of String, clsPO_POSTING),
                           ByVal dicUpdatePO_Posting As Dictionary(Of String, clsPO_POSTING),
                           ByVal dicDeletePO_Posting As Dictionary(Of String, clsPO_POSTING),
                           ByRef lstSql As List(Of String), ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      For Each objPO_POSTING In dicAddPO_Posting.Values
        If objPO_POSTING.O_Add_Insert_SQLString(lstSql, lstQueueSql) = False Then
          Result_Message = "Get Insert PO_POSTING SQL Failed"
          Return False
        End If
      Next
      For Each objPO_POSTING In dicUpdatePO_Posting.Values
        If objPO_POSTING.O_Add_Update_SQLString(lstSql, lstQueueSql) = False Then
          Result_Message = "Get Update PO_POSTING SQL Failed"
          Return False
        End If
      Next
      For Each objPO_POSTING In dicDeletePO_Posting.Values
        If objPO_POSTING.O_Add_Delete_SQLString(lstSql, lstQueueSql) = False Then
          Result_Message = "Get Delete PO_POSTING SQL Failed"
          Return False
        End If
      Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行新增的Carrier和Carrier_Status的SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef Result_Message As String,
                                         ByRef lstSql As List(Of String), ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSql.Any = True Then
        If Common_DBManagement.BatchUpdate(lstSql) = False Then
          '更新DB失敗則回傳False
          Result_Message = "eHOST 更新资料库失败"
          Return False
        End If
      End If
      Common_DBManagement.AddQueued(lstQueueSql)

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
