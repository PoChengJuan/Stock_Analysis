'20200608
'V1.0.0
'Vito
'接收到ERP的領/退料單據

Imports eCA_TransactionMessage
Imports eCA_HostObject

Module Module_SendSKUChangeData
  Public Function O_SKUManagement_SendSKUChangeData(ByRef objSKU As MSG_SendSKUChangeData,
                                     ByRef ret_strResultMsg As String) As Boolean

    Try
      '要變更的資料
      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
      Dim dicUpdate_SKU As New Dictionary(Of String, clsSKU)
      Dim SKU_NO = ""
      Dim Edition = ""
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '檢查資料
      If Check_Data(objSKU, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(SKU_NO, Edition, objSKU, ret_strResultMsg, Host_Command, dicUpdate_SKU) = False Then
        SendSKUChangeData(enuRtnCode.Fail, SKU_NO, Edition, "")
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
        SendSKUChangeData(enuRtnCode.Fail, SKU_NO, Edition, "")
        Return False
      End If
      '執行SQL與更新物件
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
        SendSKUChangeData(enuRtnCode.Fail, SKU_NO, Edition, "")
        Return False
      End If
      SendSKUChangeData(enuRtnCode.Sucess, SKU_NO, Edition, "")
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Check_Data(ByVal objSKU As MSG_SendSKUChangeData,
                                                          ByRef ret_strResultMsg As String) As Boolean
    Try
      If objSKU.EventID = "" Then
        ret_strResultMsg = "EventID is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      For Each objSKUChangeDataInfo In objSKU.SKUChangeDataList.SKUChangeDataInfo
        '檢查SKU是否為空
        If objSKUChangeDataInfo.SKU = "" Then
          ret_strResultMsg = "SKU is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查ChangeEdition是否為空
        If objSKUChangeDataInfo.ChangeEdition = "" Then
          ret_strResultMsg = "ChangeEdition is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查ChangeDateTime是否為空
        If objSKUChangeDataInfo.ChangeDateTime = "" Then
          ret_strResultMsg = "ChangeDateTime is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查NewSKU是否為空 (新品名)
        If objSKUChangeDataInfo.NewSKU = "" Then
          ret_strResultMsg = "NewSKU is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查NewSpecification是否為空
        If objSKUChangeDataInfo.NewSpecification = "" Then
          ret_strResultMsg = "NewSpecification is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查ConfirmCode是否為空
        If objSKUChangeDataInfo.ConfirmCode = "" Then
          ret_strResultMsg = "ConfirmCode is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        'For Each objSKUChangeDetailData In objSKUChangeDataInfo.SKUChangeDetailDataList.SKUChangeDetailDataInfo
        '檢查FieldID是否為空
        If objSKUChangeDataInfo.FieldID = "" Then
          ret_strResultMsg = "FieldID is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        'Next
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '新增資料或得到要更新的資料
  Private Function Get_UpdateData(ByRef ret_SKU_NO As String, ByRef ret_Edition As String, ByVal objSKU As MSG_SendSKUChangeData,
                                                                  ByRef ret_strResultMsg As String,
                                                                  ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                                                                  ByRef ret_dicUpdate_SKU As Dictionary(Of String, clsSKU)) As Boolean
    Try
      Dim dicAdd_SKU As New Dictionary(Of String, clsSKU)
      'Dim dicUpdate_SKU As New Dictionary(Of String, clsSKU)
      '取得所有的SKU
      Dim tmp_dicSKU As New Dictionary(Of String, clsSKU)
      Dim User_ID As String = objSKU.WebService_ID
      Dim Event_ID As String = objSKU.EventID
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      'Dim Companyid = objSendWorkData.Companyid
      Dim Create_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()
      Dim PickingData_TYPE = ""
      'Dim SKU_NO = ""
      'Dim Edition = ""
      'For Each objSKUDataInfo In objSKU.SKUDataList.SKUDataInfo
      '  Dim SKU_NO As String = objSKUDataInfo.SKU
      '  If tmp_dicSKU_NO.ContainsKey(SKU_NO) = False Then
      '    tmp_dicSKU_NO.Add(SKU_NO, SKU_NO)
      '  End If
      'Next
      '使用dicPO取得資料庫裡的PO資料
      If gMain.objHandling.O_GetDB_dicSKUByAll(tmp_dicSKU) = False Then
        ret_strResultMsg = "WMS get SKU data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If

      For Each objSKUChangeDataInfo In objSKU.SKUChangeDataList.SKUChangeDataInfo
        Dim SKU_NO = objSKUChangeDataInfo.SKU       '品號
        ret_SKU_NO = SKU_NO
        ret_Edition = objSKUChangeDataInfo.ChangeEdition
        Dim SKU_Key = clsSKU.Get_Combination_Key(SKU_NO)
        If tmp_dicSKU.ContainsKey(SKU_Key) = True Then
          'SKU 已存在
          Dim objNewSKU = tmp_dicSKU.Item(SKU_NO).Clone()
          With objNewSKU
            .UPDATE_TIME = objSKUChangeDataInfo.ChangeDateTime
            .SKU_ALIS1 = objSKUChangeDataInfo.NewSKU
            '.SKU_L = objSKUChangeDataInfo.NewSpecification
            '.SKU_W = objSKUChangeDataInfo.NewSpecification
            '.SKU_H = objSKUChangeDataInfo.NewSpecification
            'For Each objSKUChangeDetail In objSKUChangeDataInfo.SKUChangeDetailDataList.SKUChangeDetailDataInfo
            Dim FieldID = objSKUChangeDataInfo.FieldID
            Dim NewStringFieldValue = objSKUChangeDataInfo.NewStringFieldValue

            Select Case FieldID
                Case "MB002"  'SKUName
                .SKU_ALIS1 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_ALIS1,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB003" 'Specification
                .SKU_TYPE3 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_TYPE3,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB004"  'InventoryUnit
                .SKU_UNIT = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_UNIT,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB005"     'SKUType1
                .SKU_COMMON1 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON1,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB006"   'SKUType2
                .SKU_COMMON2 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON2,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB007" 'SKUType3
                .SKU_COMMON3 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON3,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB008" 'SKUType4
                .SKU_COMMON4 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON4,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB009" 'ProductDescription
                .SKU_DESC = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_DESC,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB010"  'StandardSKU
                .SKU_COMMON5 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON5,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB011" 'StandardCodeNumber
                .SKU_COMMON6 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON6,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB017"  'MainWarehouseId
                .SKU_COMMON7 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON7,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB019"  'InventoryManagement
                .SKU_COMMON8 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON8,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB022"  'LotManagement
                .SKU_COMMON9 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON9,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB025" 'SKUAttribute
                .SKU_COMMON10 = NewStringFieldValue
                SendMessageToLog("更改欄位: SKU_COMMON10,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB028" 'Note
                .COMMENTS = NewStringFieldValue
                SendMessageToLog("更改欄位: COMMENTS,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB030"  'EffectiveDateTime
                .EFFECTIVE_DATE = NewStringFieldValue
                SendMessageToLog("更改欄位: EFFECTIVE_DATE,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB031"  'FailureDateTime
                .FAILURE_DATE = NewStringFieldValue
                SendMessageToLog("更改欄位: FAILURE_DATE,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB043" 'InspectionMethod
                .QC_METHOD = NewStringFieldValue
                SendMessageToLog("更改欄位: QC_METHOD,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB055"   'PurchaseUnit
                .INBOUND_UNIT = NewStringFieldValue
                SendMessageToLog("更改欄位: INBOUND_UNIT,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              Case "MB056"    'SalesUnit
                .OUTBOUND_UNIT = NewStringFieldValue
                SendMessageToLog("更改欄位: OUTBOUND_UNIT,NewValue:" & NewStringFieldValue, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            End Select
            'Next
            .UPDATE_TIME = Now_Time
          End With
          ret_dicUpdate_SKU.Add(objNewSKU.gid, objNewSKU)
        Else
          ret_strResultMsg = "找不到對應SKU, SKU_NO:" & SKU_NO.ToString
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If


        If ret_dicUpdate_SKU.Any Then
          If Module_Send_WMSMessage.Send_T2F3U1_SKUManagement_to_WMS(ret_strResultMsg, ret_dicUpdate_SKU, Host_Command, enuAction.Modify.ToString) = False Then
            Return False
          End If
        End If


      Next


      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  'SQL
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
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
End Module
