Imports eCA_HostObject
Imports eCA_TransactionMessage
Module Module_T5F1S13_TransationPOExecution
  Public Function O_T5F1S13_TransationPOExecution(ByVal Receive_Msg As MSG_T5F1S13_TransationPOExecution,
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
      If Process_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      'If Get_SQL(ret_strResultMsg, dicAddPO_Posting, dicUpdatePO_Posting, dicDeletePO_Posting, lstSql, lstQueueSql) = False Then
      '  Return False
      'End If
      'If Execute_DataUpdate(ret_strResultMsg, lstSql, lstQueueSql) = False Then
      '  Return False
      'End If

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T5F1S13_TransationPOExecution,
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
  Private Function Process_Data(ByVal Receive_Msg As MSG_T5F1S13_TransationPOExecution,
                                ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim Hist_UUID As String = GetNewTime_ByDataTimeFormat(DBFullTimeUUIDFormat)
      Dim UUID As String = Receive_Msg.Header.UUID
      Dim Now_Time As String = GetNewTime_DBFormat()
      '取出所有PO_ID
      Dim tmp_dicPO_ID As New Dictionary(Of String, String)
      Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
      For Each POInfo In Receive_Msg.Body.POList.POInfo
        Dim PO_ID As String = POInfo.PO_ID
        If tmp_dicPO_ID.ContainsKey(PO_ID) = False Then
          tmp_dicPO_ID.Add(PO_ID, PO_ID)
        End If
      Next

      Dim tmp_dicPO As New Dictionary(Of String, clsPO)
      If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(tmp_dicPO_ID, tmp_dicPO) = True Then
        For Each objPO As clsPO In tmp_dicPO.Values
          Dim PO_TYPE1 = objPO.PO_Type1
          Dim PO_TYPE2 = objPO.PO_Type2
          '上報貨主調撥單
          If objPO.H_PO_ORDER_TYPE = enuOrderType.transaction_account Then
            '貨主調撥單
            SendTransactionOwnerData(enuRtnCode.Apply, objPO.PO_KEY1, objPO.PO_KEY2)
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
                           ByRef ret_dic_AddProductionInfo As Dictionary(Of String, clsProduce_Info),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得 SQL

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行SQL語句，並進行記憶體資料更新
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
