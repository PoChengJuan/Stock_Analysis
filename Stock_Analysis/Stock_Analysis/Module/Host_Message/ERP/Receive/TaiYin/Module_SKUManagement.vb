'20200608
'V1.0.0
'Vito
'接收到ERP的領/退料單據

Imports eCA_TransactionMessage
Imports eCA_HostObject

Module Module_SKUManagement
  Public Function O_SKUManagement_SendSKUData(ByRef objSKU As MSG_SendSKUData,
                                     ByRef ret_strResultMsg As String) As Boolean

    Try
      '要變更的資料
      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
      Dim dicAdd_SKU As New Dictionary(Of String, clsSKU)
      Dim dicUpdate_SKU As New Dictionary(Of String, clsSKU)
      Dim SKU_NO = ""
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '檢查資料
      If Check_Data(objSKU, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(SKU_NO, objSKU, ret_strResultMsg, Host_Command, dicAdd_SKU, dicUpdate_SKU) = False Then
        SendSKUData(enuRtnCode.Fail, SKU_NO, "")
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
        SendSKUData(enuRtnCode.Fail, SKU_NO, "")
        Return False
      End If
      '執行SQL與更新物件
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
        SendSKUData(enuRtnCode.Fail, SKU_NO, "")
        Return False
      End If
      SendSKUData(enuRtnCode.Sucess, SKU_NO, "")
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Check_Data(ByVal objSKU As MSG_SendSKUData,
                                                          ByRef ret_strResultMsg As String) As Boolean
    Try
      If objSKU.EventID = "" Then
        ret_strResultMsg = "EventID is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      For Each objSKUDataInfo In objSKU.SKUDataList.SKUDataInfo
        '檢查SKU是否為空
        If objSKUDataInfo.SKU = "" Then
          ret_strResultMsg = "SKU is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查InventoryUnit是否為空
        If objSKUDataInfo.InventoryUnit = "" Then
          ret_strResultMsg = "InventoryUnit is empty"
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
  '新增資料或得到要更新的資料
  Private Function Get_UpdateData(ByRef ret_SKU_NO As String, ByVal objSKU As MSG_SendSKUData,
                                                                  ByRef ret_strResultMsg As String,
                                                                  ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
                                                                  ByRef ret_dicAdd_SKU As Dictionary(Of String, clsSKU),
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
      'For Each objSKUDataInfo In objSKU.SKUDataList.SKUDataInfo
      '  Dim SKU_NO As String = objSKUDataInfo.SKU
      '  If tmp_dicSKU_NO.ContainsKey(SKU_NO) = False Then
      '    tmp_dicSKU_NO.Add(SKU_NO, SKU_NO)
      '  End If
      'Next
      '使用dicPO取得資料庫裡的PO資料
      If gMain.objHandling.O_GetDB_dicSKUByAll(tmp_dicSKU) = False Then
        'ret_strResultMsg = "WMS get SKU data From DB Failed"
        'SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False
      End If

      For Each objSKUDataInfo In objSKU.SKUDataList.SKUDataInfo

        Dim SKU_NO = objSKUDataInfo.SKU   '品號
        ret_SKU_NO = objSKUDataInfo.SKU
        Dim SKU_ID1 = objSKUDataInfo.SKU      '品號
        Dim SKU_ID2 = ""     '品號
        Dim SKU_ID3 = ""
        Dim SKU_ALIS1 = objSKUDataInfo.SKUName.Trim  '品名
        Dim SKU_ALIS2 = objSKUDataInfo.Specification.Trim
        Dim SKU_DESC = objSKUDataInfo.ProductDescription.Trim  '商品描述
        SKU_ALIS1 = SKU_ALIS1.Replace("'", "''")
        SKU_ALIS2 = SKU_ALIS2.Replace("'", "''")
        SKU_DESC = SKU_DESC.Replace("'", "''")
        Dim SKU_CATALOG = enuSKU_CATALOG.NULL
        Dim SKU_TYPE1 = objSKUDataInfo.SKUType1.Trim '品號分類一
        Dim SKU_TYPE2 = objSKUDataInfo.SKUType2.Trim '品號分類二 
        Dim SKU_TYPE3 = objSKUDataInfo.SKUType3.Trim '品號分類三
        Dim SKU_COMMON1 = objSKUDataInfo.SKUType4.Trim
        Dim SKU_COMMON2 = objSKUDataInfo.ASRSPart.Trim
        Dim SKU_COMMON3 =""
        Dim SKU_COMMON4 = ""
        Dim SKU_COMMON5 = ""
        Dim SKU_COMMON6 = ""
        Dim SKU_COMMON7 = ""
        Dim SKU_COMMON8 = ""
        Dim SKU_COMMON9 = ""
        Dim SKU_COMMON10 = ""
        Dim SKU_L = 0
        Dim SKU_W = 0
        Dim SKU_H = 0
        Dim SKU_WEIGHT = 0
        Dim SKU_VALUE = 0
        Dim SKU_UNIT = objSKUDataInfo.InventoryUnit.Trim '庫存單位
        Dim INBOUND_UNIT = "" '採購單位
        Dim OUTBOUND_UNIT = "" '銷售單位
        Dim HIGH_WATER = 0
        Dim LOW_WATER = 0
        Dim AVAILABLE_DAYS = 0
        Dim SAVE_DAYS = 0
        'Dim CREATE_TIME = ""
        Dim UPDATE_TIME = ""
        Dim WEIGHT_DIFFERENCE = 0
        Dim ENABLE = 1
        Dim EFFECTIVE_DATE = objSKUDataInfo.EffectiveDateTime.Trim '生效日期
        Dim FAILURE_DATE = objSKUDataInfo.FailureDateTime.Trim     '失效日期
        Dim QC_METHOD = "0" 'Int(objSKUDataInfo.InspectionMethod)       '檢驗方式
        Dim RECEIPT_DAYS = ""
        Dim DISCHARGE_DAYS = ""
        Dim RETURN_DAYS = ""
        Dim ASSIGN_AREA_NO = ""
        Dim COMMENTS = objSKUDataInfo.Comment.Trim  '備註
        COMMENTS = COMMENTS.Replace("'", "''")
        Dim SKU_Key = clsSKU.Get_Combination_Key(SKU_NO)
        If tmp_dicSKU.ContainsKey(SKU_Key) = True Then
          'SKU 已存在
          Dim objNewSKU = tmp_dicSKU.Item(SKU_NO).Clone()
          With objNewSKU
            .SKU_ID1 = SKU_ID1
            .SKU_ID2 = SKU_ID2
            .SKU_ALIS1 = SKU_ALIS1
            .SKU_TYPE1 = SKU_TYPE1
            .SKU_DESC = SKU_DESC
            .SKU_COMMON1 = SKU_COMMON1
            .SKU_COMMON2 = SKU_COMMON2
            .SKU_COMMON3 = SKU_COMMON3
            .SKU_COMMON4 = SKU_COMMON4
            .SKU_COMMON5 = SKU_COMMON5
            .SKU_COMMON6 = SKU_COMMON6
            .SKU_COMMON7 = SKU_COMMON7
            .SKU_COMMON8 = SKU_COMMON8
            .SKU_COMMON9 = SKU_COMMON9
            .SKU_COMMON10 = SKU_COMMON10
            .SKU_UNIT = SKU_UNIT
            .INBOUND_UNIT = INBOUND_UNIT
            .OUTBOUND_UNIT = OUTBOUND_UNIT
            .EFFECTIVE_DATE = EFFECTIVE_DATE
            .FAILURE_DATE = FAILURE_DATE
            .QC_METHOD = QC_METHOD
            .COMMENTS = COMMENTS
            ret_dicUpdate_SKU.Add(objNewSKU.gid, objNewSKU)
          End With
        Else
          Dim objNewSKU = New clsSKU(SKU_NO, SKU_ID1, SKU_ID2, SKU_ID3, SKU_ALIS1, SKU_ALIS2, SKU_DESC, SKU_CATALOG, SKU_TYPE1, SKU_TYPE2, SKU_TYPE3,
                                     SKU_COMMON1, SKU_COMMON2, SKU_COMMON3, SKU_COMMON4, SKU_COMMON5, SKU_COMMON6, SKU_COMMON7, SKU_COMMON8, SKU_COMMON9, SKU_COMMON10,
                                     SKU_L, SKU_W, SKU_H, SKU_WEIGHT, SKU_VALUE, SKU_UNIT, INBOUND_UNIT, OUTBOUND_UNIT, HIGH_WATER, LOW_WATER, AVAILABLE_DAYS, SAVE_DAYS,
                                     Create_Time, UPDATE_TIME, WEIGHT_DIFFERENCE, ENABLE, EFFECTIVE_DATE, FAILURE_DATE, QC_METHOD, COMMENTS)
          If ret_dicAdd_SKU.ContainsKey(objNewSKU.gid) = False Then
            ret_dicAdd_SKU.Add(objNewSKU.gid, objNewSKU)
          End If

        End If




      Next
      If ret_dicAdd_SKU.Any Then
        If Module_Send_WMSMessage.Send_T2F3U1_SKUManagement_to_WMS(ret_strResultMsg, ret_dicAdd_SKU, Host_Command, enuAction.Create.ToString) = False Then
          Return False
        End If
      ElseIf ret_dicUpdate_SKU.Any Then
        If Module_Send_WMSMessage.Send_T2F3U1_SKUManagement_to_WMS(ret_strResultMsg, ret_dicUpdate_SKU, Host_Command, enuAction.Modify.ToString) = False Then
          Return False
        End If
      End If

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
