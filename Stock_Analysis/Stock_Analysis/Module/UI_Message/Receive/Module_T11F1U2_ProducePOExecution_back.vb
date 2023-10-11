'20180629
'V1.0.0
'Jerry

'狀態:Checked

'1.修改WMS_T_PO
'2.修改WMS_CT_Produce_Info

Imports eCA_HOSTObject
Imports eCA_TransactionMessage

Module Module_T11F1U2_ProducePOExecution_back
  Public Function O_Process_Message(ByVal Receive_Msg As MSG_T11F1U2_ProducePOExecution,
                                    ByRef ret_strResultMsg As String) As Boolean
    Try
      '要新增的Carrier資料
      Dim dicUpdate_PO As New Dictionary(Of String, clsPO)
      Dim dicAdd_CProduce_Info As New Dictionary(Of String, clsProduce_Info)
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)
      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '取得要新增的Carrier和Carrier_Status的資料
      If Get_Data(Receive_Msg, ret_strResultMsg, dicUpdate_PO, dicAdd_CProduce_Info) = False Then
        Return False
      End If
      '取得要更新到DB的SQL
      If Get_SQL(ret_strResultMsg, dicUpdate_PO, dicAdd_CProduce_Info, lstSql) = False Then
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(ret_strResultMsg, dicUpdate_PO, dicAdd_CProduce_Info, lstSql) = False Then
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_T11F1U2_ProducePOExecution,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim PO_Type1 As String = objPOInfo.PO_TYPE1
        Dim PO_Type2 As String = objPOInfo.PO_TYPE2
        Dim PO_Type3 As String = objPOInfo.PO_TYPE3
        '檢查PO_ID是否為空
        If PO_ID = "" Then
          ret_strResultMsg = "PO_ID is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查PO_Type1是否正確
        If PO_Type1 = "" Then
          ret_strResultMsg = "PO_Type1 is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If PO_Type1 <> enuPOType_1.Produce_in Then
          ret_strResultMsg = "PO_Type1 is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
          ret_strResultMsg = "PO_TYPE1 is not " & enuPOType_1.Produce_in & " , PO_ID =" & PO_ID & ", PO_Type1 =" & PO_Type1
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
  '取得CarrierInstall要修改的資料
  Private Function Get_Data(ByVal Receive_Msg As MSG_T11F1U2_ProducePOExecution,
                            ByRef ret_strResultMsg As String,
                            ByRef dicUpdate_PO As Dictionary(Of String, clsPO),
                            ByRef dicAdd_CProduce_Info As Dictionary(Of String, clsProduce_Info)) As Boolean
    Try
      Dim UserID As String = Receive_Msg.Header.ClientInfo.UserID
      Dim ClientID As String = Receive_Msg.Header.ClientInfo.ClientID
      Dim dicPO_ID As New Dictionary(Of String, String)
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim PO_Type1 As String = objPOInfo.PO_TYPE1
        Dim PO_Type2 As String = objPOInfo.PO_TYPE2
        Dim PO_Type3 As String = objPOInfo.PO_TYPE3
        If dicPO_ID.ContainsKey(PO_ID) = False Then
          dicPO_ID.Add(PO_ID, PO_ID)
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
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef dicUpdate_PO As Dictionary(Of String, clsPO),
                           ByRef dicAdd_CProduce_Info As Dictionary(Of String, clsProduce_Info),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得PO的Update SQL
      For Each obj As clsPO In dicUpdate_PO.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Update WMS_T_PO SQL Failed"
          Return False
        End If
      Next
      '取得新增LocationUseMap的Insert SQL
      For Each obj As clsProduce_Info In dicAdd_CProduce_Info.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert WMS_CT_Produce_Info SQL Failed"
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
  '執行新增的Carrier和Carrier_Status的SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                      ByRef dicUpdate_PO As Dictionary(Of String, clsPO),
                                      ByRef dicAdd_CProduce_Info As Dictionary(Of String, clsProduce_Info),
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
