'20210823
'V1.0.0
'Vito
'餘料儲位點燈
'狀態:Open

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T5F2U62_AutoInbound


  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~接收回傳的結果~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  Public Function O_CheckMessageResult(ByVal Receive_Msg As MSG_T5F2U62_AutoInbound,
                                       ByRef ret_strResultMsg As String,
                                       ByVal strRejectReason As String,
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
      If Get_Data(Receive_Msg, ret_strResultMsg, Host_Command) = False Then
        Return False
      End If
      '取得要更新到DB的SQL
      If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(ret_strResultMsg, dic_AddProductionInfo, lstSql) = False Then
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_T5F2U62_AutoInbound,
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
  Private Function Get_Data(ByVal Receive_Msg As MSG_T5F2U62_AutoInbound,
                            ByRef ret_strResultMsg As String,
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try
      Dim Hist_UUID As String = GetNewTime_ByDataTimeFormat(DBFullTimeUUIDFormat)
      Dim UUID As String = Receive_Msg.Header.UUID
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim dicLocation_No As New Dictionary(Of String, String)
      Dim dicReturnSupplierSetting As New Dictionary(Of String, clsRETURNSUPPLIERSETTING)

      gMain.objHandling.O_GetDB_dicReturnSupplierSettingByAll(dicReturnSupplierSetting)

      If dicReturnSupplierSetting.Any = False Then

      End If
      For Each PO_DTL_INFO In Receive_Msg.Body.PODetailList.PODetailInfo
        For Each obj In dicReturnSupplierSetting.Values
          If PO_DTL_INFO.SKU_NO = obj.SUPPLIER_NO Then
            If dicLocation_No.ContainsKey(obj.LOCATION_NO) = False Then
              dicLocation_No.Add(obj.LOCATION_NO, obj.LOCATION_NO)
              Exit For
            End If
          End If
        Next
      Next


      If Send_T11F1U14_SwitchOnLocationLight_to_WMS(ret_strResultMsg, Host_Command, dicLocation_No) = False Then
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
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得 SQL
      For Each obj As clsFromHostCommand In Host_Command.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert Host_Command Info SQL Failed"
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

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function



End Module
