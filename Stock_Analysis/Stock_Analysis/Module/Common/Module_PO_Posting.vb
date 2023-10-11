Imports System.Threading
Imports eCA_HostObject
Imports eCA_TransactionMessage

'取得所有过帐结果 全数通过则加入Delete行列
Module Module_PO_Posting

  '過帳結果的回覆
  Public Function O_PO_POSTING_Check(ByVal Result_Message As String, ByVal WO_ID As String, ByRef POSTING_FLAG As Boolean, ByRef return_MSG As List(Of String)) As Boolean
    Try
      Dim tmp_dic_PO_POSTING As New Dictionary(Of String, clsPO_POSTING)
      Dim dicDelete_PO_POSTING As New Dictionary(Of String, clsPO_POSTING)
      Dim dicInsert_PO_POSTING_HIST As New Dictionary(Of String, clsPO_POSTING_HIST)

      '儲存要更新的SQL， 進行一次性更新
      Dim lstSql As New List(Of String)
      '儲存要更新的SQL， 進行一次性更新
      Dim lstQueueSql As New List(Of String)


      Dim dicPO_POSTING As New Dictionary(Of String, clsPO_POSTING)
      If gMain.objHandling.O_Get_dicPO_POSTING_By_WO_ID(WO_ID, dicPO_POSTING) = False Then
        SendMessageToLog("O_Get_dicPO_POSTING_By_WO_ID Format Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

      Dim tmp_result_MSG = "工单号:" & WO_ID '记录此笔工单所有过帐结果
      '比对成功笔数 若相符则加入结束行列
      Dim Success_Count = 0
      For Each objPO_POSTING In dicPO_POSTING.Values
        tmp_result_MSG += ";订单号:" & objPO_POSTING.PO_ID & " 项次:" & objPO_POSTING.PO_LINE_NO
        tmp_result_MSG += " ;数量:" & objPO_POSTING.QTY
        'If objPO_POSTING.SORT_ITEM_COMMON1 <> "" Then tmp_result_MSG += " ;库存类型:" & objPO_POSTING.SORT_ITEM_COMMON1
        'If objPO_POSTING.SORT_ITEM_COMMON2 <> "" Then tmp_result_MSG += " ;供应商:" & objPO_POSTING.SORT_ITEM_COMMON2
        ''If objPO_POSTING.SORT_ITEM_COMMON3 <> "" Then tmp_result_MSG += " ;料品类型:" & objPO_POSTING.SORT_ITEM_COMMON3
        ''If objPO_POSTING.SORT_ITEM_COMMON4 <> "" Then tmp_result_MSG += " ;尚未定义:" & objPO_POSTING.SORT_ITEM_COMMON4
        ''If objPO_POSTING.SORT_ITEM_COMMON5 <> "" Then tmp_result_MSG += " ;箱数/容量/尾数:" & objPO_POSTING.SORT_ITEM_COMMON5

        If objPO_POSTING.RESULT = 0 Then '加入有过帐成功的 后续比对笔数用
          Success_Count += 1
          tmp_result_MSG += "; 过帐成功，返回:" & objPO_POSTING.RESULT_MESSAGE
        Else
          tmp_result_MSG += "; 过帐""失败"" :" & objPO_POSTING.RESULT_MESSAGE

        End If
        tmp_result_MSG += " ;;"
      Next

      return_MSG.Add(tmp_result_MSG)

      '比对笔数 若新增的数量等于成功数量则全数通过
      If dicPO_POSTING.Count = Success_Count And dicPO_POSTING.Any = True Then
        dicDelete_PO_POSTING.AddRange(dicPO_POSTING)
        POSTING_FLAG = True
        SendMessageToLog("已全数过帐完成。", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

        For Each objPO_POSTING As clsPO_POSTING In dicPO_POSTING.Values
          Dim PO_ID = objPO_POSTING.PO_ID
          Dim PO_SERIAL_NO = objPO_POSTING.PO_SERIAL_NO
          Dim PO_LINE_NO = objPO_POSTING.PO_LINE_NO
          Dim dicPO As New Dictionary(Of String, clsPO)
          If gMain.objHandling.O_GetDB_dicPOByPOID(PO_ID, dicPO) = False Then
            Continue For
          End If
          '如果是运单 要加进运单为过账成功 才能卡掉重复提单
          If dicPO.First.Value.H_PO_ORDER_TYPE = enuOrderType.None Then
            SendMessageToLog("成品出库，将运单填入历史，以防重复提单。", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Dim NewPO_POSTING As New clsPO_POSTING_HIST(dicPO.First.Value.H_PO16, "", "", "", "", "", "", "", 0, "", ModuleHelpFunc.GetNewTime_DBFormat, "", 0, "此为运单号 防止重复提单用", "", "", "", "", "", ModuleHelpFunc.GetNewTime_DBFormat, "", "WMS", "", "", enuOrderType.None, "", dicPO.First.Value.H_PO16, "", "", "")
            If dicInsert_PO_POSTING_HIST.ContainsKey(NewPO_POSTING.gid) = False Then
              dicInsert_PO_POSTING_HIST.Add(NewPO_POSTING.gid, NewPO_POSTING)
            End If
          End If
        Next
      End If
      '取得要更新到DB的SQL
      If Get_SQL_PO_POSTING_Check(Result_Message, dicDelete_PO_POSTING, dicInsert_PO_POSTING_HIST, lstSql, lstQueueSql) = False Then
        POSTING_FLAG = False
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(Result_Message, lstSql, lstQueueSql) = False Then
        POSTING_FLAG = False
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_PO_POSTING_INIT(ByRef Result_Message As String, ByVal PO_ID As String, ByVal WO_ID As String, ByVal UUID As String,
                                    ByVal objPO_DTLInfo As List(Of MSG_T11F1S1_POClose.BodyData.clsPOList.clsPOInfo.clsPO_DTLList.clsPO_DTLInfo),
                                    ByVal H_PO_ORDER_TYPE As enuOrderType, ByRef lstH_PO16 As List(Of String), ByVal USER_ID As String,
                                    ByRef bln_DeleteOldPosting As Boolean, ByRef bln_First_Init As Boolean,
                                    ByVal Forced_Close As String) As Boolean
    Try
      Dim dicAdd_PO_POSTING As New Dictionary(Of String, clsPO_POSTING)
      Dim dicDelete_PO_POSTING As New Dictionary(Of String, clsPO_POSTING)

      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)
      '儲存要更新的SQL，進行一次性更新
      Dim lstQueueSql As New List(Of String)

      '根据存存在与否决定要不要新增资料
      If Get_Data(Result_Message, PO_ID, WO_ID, UUID, objPO_DTLInfo, dicAdd_PO_POSTING, dicDelete_PO_POSTING, H_PO_ORDER_TYPE, lstH_PO16, USER_ID, bln_DeleteOldPosting, bln_First_Init, Forced_Close) = False Then
        Return False
      End If

      If Get_SQL(Result_Message, dicAdd_PO_POSTING, dicDelete_PO_POSTING, lstSql, lstQueueSql) = False Then
        Return False
      End If

      If Execute_DataUpdate(Result_Message, lstSql, lstQueueSql) = False Then
        Return False
      End If

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_PO_POSTING_Forced_Close(ByRef Result_Message As String, ByVal WO_ID As String) As Boolean
    Try
      Dim dicUpdate_PO_POSTING As New Dictionary(Of String, clsPO_POSTING)

      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)
      '儲存要更新的SQL，進行一次性更新
      Dim lstQueueSql As New List(Of String)

      '根据存存在与否决定要不要新增资料
      If Get_Data_Forced_Close(Result_Message, WO_ID, dicUpdate_PO_POSTING) = False Then
        Return False
      End If

      If Get_SQL_PO_POSTING_Forced_Close(Result_Message, dicUpdate_PO_POSTING, lstSql, lstQueueSql) = False Then
        Return False
      End If

      If Execute_DataUpdate(Result_Message, lstSql, lstQueueSql) = False Then
        Return False
      End If

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '取得要修改的資料
  Private Function Get_Data(ByRef Result_Message As String, ByVal PO_ID As String, ByVal WO_ID As String, ByVal UUID As String,
                            ByVal objPO_DTLInfo As List(Of MSG_T11F1S1_POClose.BodyData.clsPOList.clsPOInfo.clsPO_DTLList.clsPO_DTLInfo),
                            ByRef dicAdd_PO_POSTING As Dictionary(Of String, clsPO_POSTING),
                            ByRef dicDelete_PO_POSTING As Dictionary(Of String, clsPO_POSTING),
                            ByRef H_PO_ORDER_TYPE As enuOrderType, ByRef lstH_PO16 As List(Of String), ByVal USER_ID As String,
                            ByRef bln_DeleteOldPosting As Boolean, ByRef bln_First_Init As Boolean,
                            ByRef Forced_Close As String) As Boolean
    Try
      'Dim objWO As clsWO = Nothing
      Dim Start_Transfer_Time = ""
      Dim Finish_Transfer_Time = ""
      'If gMain.objHandling.O_Get_WO(WO_ID, objWO) Then
      '    SendMessageToLog("取得WO开始搬送、结束时间", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      '    Start_Transfer_Time = IIf(objWO.Start_Transfer_Time = "", objWO.Start_Time, objWO.Start_Transfer_Time)
      '    If Start_Transfer_Time = "" Then
      '        Start_Transfer_Time = objWO.Create_Time
      '    End If
      '    Finish_Transfer_Time = IIf(objWO.Finish_Transfer_Time = "", GetNewTime_DBFormat(), objWO.Finish_Transfer_Time)
      '    'If Finish_Transfer_Time = "" Then
      '    '  Finish_Transfer_Time = GetNewTime_DBFormat()
      '    'End If
      'End If

      Dim dicPO_POSTING As New Dictionary(Of String, clsPO_POSTING)
      If gMain.objHandling.O_Get_dicPO_POSTING_By_WO_ID(WO_ID, dicPO_POSTING) = False Then
        Result_Message = "O_Get_dicPO_POSTING_By_WO_ID Format Error"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      If bln_First_Init = False Then
        bln_DeleteOldPosting = False
      End If
      If bln_First_Init Then
        bln_First_Init = False
      End If

      If dicPO_POSTING.Any = False Then
        '不存在则新增 存在则不处理
        For Each PO_DTLINFO In objPO_DTLInfo

          Dim PO_LINE_NO = PO_DTLINFO.PO_LINE_NO
          Dim PO_SERIAL_NO = PO_DTLINFO.PO_SERIAL_NO
          Dim SORT_ITEM_COMMON1 = IIf(PO_DTLINFO.SORT_ITEM_COMMON1 = "K", PO_DTLINFO.SORT_ITEM_COMMON1, "")
          Dim SORT_ITEM_COMMON2 = PO_DTLINFO.SORT_ITEM_COMMON2
          Dim SORT_ITEM_COMMON3 = PO_DTLINFO.SORT_ITEM_COMMON3
          Dim SORT_ITEM_COMMON4 = PO_DTLINFO.SORT_ITEM_COMMON4
          Dim SORT_ITEM_COMMON5 = PO_DTLINFO.SORT_ITEM_COMMON5
          Dim Result = -1
          Dim Message = "尚未过账"
          Dim QTY = PO_DTLINFO.QTY
          Dim H_POP1 = "0" '過帳次數
          Dim H_POP2 = ""
          Dim H_POP3 = ""
          Dim H_POP4 = ""
          Dim H_POP5 = ""
          Dim SKU_NO = PO_DTLINFO.SKU_NO
          Dim LOT_NO = ""
          Dim OWNER = ""
          Dim SubOwner = ""

          '根據單號、行項取得批號
          Dim dicPODTL As New Dictionary(Of String, clsPO_DTL)
          If gMain.objHandling.O_GetDB_dicPODTLByPOID_POSerialNo(PO_ID, PO_SERIAL_NO, dicPODTL) Then
            If dicPODTL.Any Then
              LOT_NO = dicPODTL.First.Value.LOT_NO
              OWNER = IIf(dicPODTL.First.Value.FROM_OWNER_ID = "", dicPODTL.First.Value.TO_OWNER_ID, dicPODTL.First.Value.FROM_OWNER_ID)
              SubOwner = IIf(dicPODTL.First.Value.FROM_SUB_OWNER_ID = "", dicPODTL.First.Value.TO_SUB_OWNER_ID, dicPODTL.First.Value.FROM_SUB_OWNER_ID)
            End If
          End If
          Dim dicPOLINE As New Dictionary(Of String, clsPO_LINE)
          If gMain.objHandling.O_Get_dicPOLineByPOID_POLineNo(PO_ID, PO_LINE_NO, dicPOLINE) Then
            If dicPOLINE.Any = True Then
              If dicPOLINE.First.Value.QTY <= dicPOLINE.First.Value.H_QTY_PROCESS Then
                Continue For
              End If
            End If
          End If

          Dim dicPO As Dictionary(Of String, clsPO) = Nothing
          Dim objPO As clsPO = Nothing
          If gMain.objHandling.O_Get_dicPOByPO_ID(PO_ID, dicPO) = False Then
            Result_Message = "无法取得订单资讯 PO_ID = " & PO_ID
            Return False
          Else
            If dicPO.Count <> 0 Then
              objPO = dicPO.Values(0)
            Else
              Result_Message = "不存在订单 PO_ID = " & PO_ID
              SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If
          End If


          Dim dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
          If gMain.objHandling.O_Get_dicPODTLByPOID_POLineNo_SERIAL_NO(PO_ID, PO_LINE_NO, PO_SERIAL_NO, dicPO_DTL) = False Then
            Result_Message = "PO_ID: " & PO_ID & " PO_LINE_NO: " & PO_LINE_NO & " 取资料失败"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If dicPO_DTL.Any = False Then
            Result_Message = "PO_ID: " & PO_ID & " PO_LINE_NO: " & PO_LINE_NO & " 不存在资料"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If

          If Forced_Close = 1 And QTY = 0 Then
            Result = enuPOSTING_RESULT.Complete
            Message = "强制结单(需上报)但数量为0"
          Else
            Result = -1
            Message = "尚未过账"
          End If

          Dim New_objPO_POSTING As New clsPO_POSTING(PO_ID, PO_LINE_NO, WO_ID, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4,
                                                     SORT_ITEM_COMMON5, QTY, UUID, ModuleHelpFunc.GetNewTime_DBFormat, ModuleHelpFunc.GetNewTime_DBFormat, Result,
                                                     Message, H_POP1, H_POP2, H_POP3, H_POP4, H_POP5, SKU_NO, USER_ID, Start_Transfer_Time, Finish_Transfer_Time,
                                                     H_PO_ORDER_TYPE, PO_SERIAL_NO, objPO.H_PO16, LOT_NO, OWNER, SubOwner)

          If dicAdd_PO_POSTING.ContainsKey(New_objPO_POSTING.gid) = False Then
            dicAdd_PO_POSTING.Add(New_objPO_POSTING.gid, New_objPO_POSTING)
          Else
            Result_Message = "过账资料存在相同Key"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
          End If
        Next
        'Else
        '  Result_Message = "WO_ID is not exist PO_POSTING, WO_ID = " & WO_ID
        '  SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      Else
        If bln_DeleteOldPosting Then
          dicDelete_PO_POSTING.AddRange(dicPO_POSTING)
        End If

        For Each PO_DTLINFO In objPO_DTLInfo
          Dim PO_LINE_NO = PO_DTLINFO.PO_LINE_NO
          Dim PO_SERIAL_NO = PO_DTLINFO.PO_SERIAL_NO
          Dim SORT_ITEM_COMMON1 = IIf(PO_DTLINFO.SORT_ITEM_COMMON1 = "K", PO_DTLINFO.SORT_ITEM_COMMON1, "")
          Dim SORT_ITEM_COMMON2 = PO_DTLINFO.SORT_ITEM_COMMON2
          Dim SORT_ITEM_COMMON3 = PO_DTLINFO.SORT_ITEM_COMMON3
          Dim SORT_ITEM_COMMON4 = PO_DTLINFO.SORT_ITEM_COMMON4
          Dim SORT_ITEM_COMMON5 = PO_DTLINFO.SORT_ITEM_COMMON5
          Dim Result = -1
          Dim Message = "尚未过账"
          Dim QTY = PO_DTLINFO.QTY
          Dim H_POP1 = 0 '過帳次數
          Dim H_POP2 = ""
          Dim H_POP3 = ""
          Dim H_POP4 = ""
          Dim H_POP5 = ""
          Dim SKU_NO = PO_DTLINFO.SKU_NO
          Dim LOT_NO = ""
          Dim OWNER = ""
          Dim SubOwner = ""

          '根據單號、行項取得批號
          Dim dicPODTL As New Dictionary(Of String, clsPO_DTL)
          If gMain.objHandling.O_GetDB_dicPODTLByPOID_POSerialNo(PO_ID, PO_SERIAL_NO, dicPODTL) Then
            If dicPODTL.Any Then
              LOT_NO = dicPODTL.First.Value.LOT_NO
              OWNER = IIf(dicPODTL.First.Value.FROM_OWNER_ID = "", dicPODTL.First.Value.TO_OWNER_ID, dicPODTL.First.Value.FROM_OWNER_ID)
              SubOwner = IIf(dicPODTL.First.Value.FROM_SUB_OWNER_ID = "", dicPODTL.First.Value.TO_SUB_OWNER_ID, dicPODTL.First.Value.FROM_SUB_OWNER_ID)
            End If
          End If

          Dim dicPOLINE As New Dictionary(Of String, clsPO_LINE)
          If gMain.objHandling.O_Get_dicPOLineByPOID_POLineNo(PO_ID, PO_LINE_NO, dicPOLINE) Then
            If dicPOLINE.Any = True Then
              If dicPOLINE.First.Value.QTY <= dicPOLINE.First.Value.H_QTY_PROCESS Then
                Continue For
              End If
            End If
          End If

          Dim dicPO As Dictionary(Of String, clsPO) = Nothing
          Dim objPO As clsPO = Nothing
          If gMain.objHandling.O_Get_dicPOByPO_ID(PO_ID, dicPO) = False Then
            Result_Message = "无法取得订单资讯 PO_ID = " & PO_ID
            Return False
          Else
            If dicPO.Count <> 0 Then
              objPO = dicPO.Values(0)
            Else
              Result_Message = "不存在订单 PO_ID = " & PO_ID
              SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If
          End If



          Dim dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
          If gMain.objHandling.O_Get_dicPODTLByPOID_POLineNo_SERIAL_NO(PO_ID, PO_LINE_NO, PO_SERIAL_NO, dicPO_DTL) = False Then
            Result_Message = "PO_ID: " & PO_ID & " PO_LINE_NO: " & PO_LINE_NO & " 取资料失败"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If dicPO_DTL.Any = False Then
            Result_Message = "PO_ID: " & PO_ID & " PO_LINE_NO: " & PO_LINE_NO & " 不存在资料"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If H_PO_ORDER_TYPE = enuOrderType.m_material_in Or H_PO_ORDER_TYPE = enuOrderType.m_material_out Or
              H_PO_ORDER_TYPE = enuOrderType.m_semiSKU_in Or H_PO_ORDER_TYPE = enuOrderType.m_SKU_in Or
              H_PO_ORDER_TYPE = enuOrderType.m_general_in Or H_PO_ORDER_TYPE = enuOrderType.m_semiSKU_out Or
              H_PO_ORDER_TYPE = enuOrderType.m_SKU_out Or H_PO_ORDER_TYPE = enuOrderType.m_grneral_out Then
            '成品手工单不用上报
            Dim tmp_objPO_POSTING As New clsPO_POSTING(PO_ID, PO_LINE_NO, WO_ID, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4,
                                                         SORT_ITEM_COMMON5, QTY, UUID, ModuleHelpFunc.GetNewTime_DBFormat, ModuleHelpFunc.GetNewTime_DBFormat,
                                                         0, "WMS手工單", H_POP1, H_POP2, H_POP3, H_POP4, H_POP5, SKU_NO, USER_ID, Start_Transfer_Time,
                                                         Finish_Transfer_Time, H_PO_ORDER_TYPE, PO_SERIAL_NO, objPO.H_PO16, LOT_NO, OWNER, SubOwner)
            If dicAdd_PO_POSTING.ContainsKey(tmp_objPO_POSTING.gid) = False Then
              dicAdd_PO_POSTING.Add(tmp_objPO_POSTING.gid, tmp_objPO_POSTING)
            End If
            Continue For
          End If

          If Forced_Close = 1 And QTY = 0 Then
            Result = enuPOSTING_RESULT.Complete
            Message = "强制结单(需上报)但数量为0"
          Else
            Result = -1
            Message = "尚未过账"
          End If

          Dim New_objPO_POSTING As New clsPO_POSTING(PO_ID, PO_LINE_NO, WO_ID, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4,
                                                     SORT_ITEM_COMMON5, QTY, UUID, ModuleHelpFunc.GetNewTime_DBFormat, ModuleHelpFunc.GetNewTime_DBFormat, Result,
                                                     Message, H_POP1, H_POP2, H_POP3, H_POP4, H_POP5, SKU_NO, USER_ID, Start_Transfer_Time, Finish_Transfer_Time,
                                                     H_PO_ORDER_TYPE, PO_SERIAL_NO, objPO.H_PO16, LOT_NO, OWNER, SubOwner)
          Dim objTmpPO_POSTING As clsPO_POSTING = Nothing
          If dicPO_POSTING.TryGetValue(New_objPO_POSTING.gid, objTmpPO_POSTING) Then
            SendMessageToLog("存在過帳資訊，根據狀態進行數量的更新。", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            '若已存在 且过过账 则不理  反之 更新数量
            If objTmpPO_POSTING.RESULT <> 0 Or objTmpPO_POSTING.RESULT <> 2 Then
              objTmpPO_POSTING.QTY = QTY
            End If

            If dicAdd_PO_POSTING.ContainsKey(objTmpPO_POSTING.gid) = False Then
              dicAdd_PO_POSTING.Add(objTmpPO_POSTING.gid, objTmpPO_POSTING)
            End If
          Else
            SendMessageToLog("過帳資訊不存在，進行新增。", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            If dicAdd_PO_POSTING.ContainsKey(New_objPO_POSTING.gid) = False Then
              dicAdd_PO_POSTING.Add(New_objPO_POSTING.gid, New_objPO_POSTING)
            Else
              Result_Message = "过账资料存在相同Key"
              SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
              Return False
            End If
          End If
        Next
      End If



      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef dicAdd_PO_POSTING As Dictionary(Of String, clsPO_POSTING),
                           ByRef dicDelete_PO_POSTING As Dictionary(Of String, clsPO_POSTING),
                            ByRef lstSql As List(Of String), ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      For Each objPO_POSTING In dicDelete_PO_POSTING.Values
        If objPO_POSTING.O_Add_Delete_SQLString(lstSql, lstQueueSql) = False Then
          Result_Message = "Get Delete PO_POSTING SQL Failed"
          Return False
        End If
      Next
      For Each objPO_POSTING As clsPO_POSTING In dicAdd_PO_POSTING.Values
        If objPO_POSTING.O_Add_Insert_SQLString(lstSql, lstQueueSql) = False Then
          Result_Message = "Get Insert PO_POSTING SQL Failed"
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

  '取得要新增的SQL語句
  Private Function Get_SQL_PO_POSTING_Check(ByRef Result_Message As String,
                                            ByRef dicDelete_PO_POSTING As Dictionary(Of String, clsPO_POSTING),
                                            ByRef dicInsert_PO_POSTING_HIST As Dictionary(Of String, clsPO_POSTING_HIST),
                                            ByRef lstSql As List(Of String), ByRef lstQueueSql As List(Of String)) As Boolean
    Try

      For Each objPO_POSTING As clsPO_POSTING In dicDelete_PO_POSTING.Values
        If objPO_POSTING.O_Add_Delete_SQLString(lstSql, lstQueueSql) = False Then
          Result_Message = "Get Delete PO_POSTING SQL Failed"
          Return False
        End If
      Next
      For Each objInsert_PO_POSTING_HIST As clsPO_POSTING_HIST In dicInsert_PO_POSTING_HIST.Values
        If objInsert_PO_POSTING_HIST.O_Add_Insert_SQLString(lstSql) = False Then
          Result_Message = "Get Insert PO_POSTING_HIST SQL Failed"
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
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        Result_Message = "eHOST 更新资料库失败"
        Return False
      End If
      Common_DBManagement.AddQueued(lstQueueSql)

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Get_Data_Forced_Close(ByRef Result_Message As String, ByVal WO_ID As String,
                           ByRef dicUpdate_PO_POSTING As Dictionary(Of String, clsPO_POSTING)) As Boolean
    Try

      Dim dicPO_POSTING As New Dictionary(Of String, clsPO_POSTING)

      If gMain.objHandling.O_Get_dicPO_POSTING_By_WO_ID(WO_ID, dicPO_POSTING) = False Then
        Result_Message = "O_Get_dicPO_POSTING_By_WO_ID Format Error"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

      If dicPO_POSTING.Any Then
        For Each objPO_POSTING In dicPO_POSTING.Values
          If objPO_POSTING.RESULT <> 0 Then
            objPO_POSTING.RESULT = 0
            objPO_POSTING.RESULT_MESSAGE = "强制过账, 结单前内容: " & objPO_POSTING.RESULT_MESSAGE
            '等于0的不需要理会
            'Else
            '  objPO_POSTING.RESULT_MESSAGE = "强制过账, 结单前内容: " & objPO_POSTING.RESULT_MESSAGE
          End If
          dicUpdate_PO_POSTING.Add(objPO_POSTING.gid, objPO_POSTING)
        Next
      Else
        Result_Message = "WO_ID is not exist PO_POSTING, WO_ID = " & WO_ID
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If



      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL_PO_POSTING_Forced_Close(ByRef Result_Message As String,
                           ByRef dicUpdate_PO_POSTING As Dictionary(Of String, clsPO_POSTING),
                            ByRef lstSql As List(Of String), ByRef lstQueueSql As List(Of String)) As Boolean
    Try

      For Each objPO_POSTING As clsPO_POSTING In dicUpdate_PO_POSTING.Values
        If objPO_POSTING.O_Add_Update_SQLString(lstSql, lstQueueSql) = False Then
          Result_Message = "Get Update PO_POSTING SQL Failed"
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

End Module
