'20180629
'V1.0.0
'Jerry

'结单

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T11F1S1_POClose_old
  Public Function O_T11F1S1_POClose(ByVal Receive_Msg As MSG_T11F1S1_POClose,
                                          ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料處理
      If Process_Data(Receive_Msg, ret_strResultMsg) = False Then
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
  Private Function Process_Data(ByVal Receive_Msg As MSG_T11F1S1_POClose,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim USER_ID = Receive_Msg.Header.ClientInfo.UserID
      Dim UUID = Receive_Msg.Header.UUID
      Dim Forced_Close = Receive_Msg.Body.Forced_Close '强制结单
      Dim WO_ID = "" ' Receive_Msg.Body.WO_ID '工单编号
      If Receive_Msg.Body.POList Is Nothing Or Receive_Msg.Body.POList.POInfo.Count = 0 Then
        ret_strResultMsg = "WMS 给的结单资讯有缺(POList)，无法结单。"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim POSTING_FLAG = False '确定能否过账
      Dim lstH_PO16 As New List(Of String) '需要重提校验的运单
      Dim bln_DeleteOldPosting = True
      Dim bln_First_Init = True

      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        '执行结单前的初始化

        If Module_PO_Posting.O_PO_POSTING_INIT(ret_strResultMsg, PO_ID, WO_ID, UUID, objPOInfo.PO_DTLList.PO_DTLInfo, H_PO_ORDER_TYPE, lstH_PO16, USER_ID, bln_DeleteOldPosting, bln_First_Init, Forced_Close) = False Then
          Return False
        End If

        If Forced_Close = 2 Then
          '强制结单   '不上报
          If Module_PO_Posting.O_PO_POSTING_Forced_Close(ret_strResultMsg, WO_ID) = False Then
            Return False
          End If
        Else
          Select Case H_PO_ORDER_TYPE
            Case enuOrderType.Inbound_Data '
              SendMessageToLog("单据类型为 ERP入庫", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

            Case enuOrderType.material_out
              SendMessageToLog("单据类型为 ERP出庫", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

            Case enuOrderType.m_material_in
              SendMessageToLog("单据类型为 WMS手動入庫單，无需向上位系统过帐", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

            Case enuOrderType.m_material_out
              SendMessageToLog("单据类型为 WMS手動出庫單，无需向上位系统过帐", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

          End Select
        End If
      Next

      SendMessageToLog("检查过帐是否全数通过", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

      '環鴻為多次過帳失敗 無所謂 也是算成功
      '目前無須回報過帳結果 (連第一次都不回)
      'Return True

      '都更新完后 做最后检查
      Dim return_MSG As New List(Of String) '显示过账结果，成功失败都要回复
      If Module_PO_Posting.O_PO_POSTING_Check(ret_strResultMsg, WO_ID, POSTING_FLAG, return_MSG) = False Then
        Return False
      End If

      'USI過帳客製
      POSTING_FLAG = True
      If POSTING_FLAG Then
        ret_strResultMsg = "成功!!;"
        Return True
        For Each msg In return_MSG
          ret_strResultMsg += msg & ";"
        Next
      End If

      If ret_strResultMsg <> "" Then
        For Each msg In return_MSG
          ret_strResultMsg += msg & ";"
        Next
        Return False
      ElseIf ret_strResultMsg = "" And POSTING_FLAG = False Then
        ret_strResultMsg = "失败!! 尚未全数过账成功 "
        For Each msg In return_MSG
          ret_strResultMsg += msg & ";"
        Next
        Return False
      End If

      Return False
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Module
