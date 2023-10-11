'20200608
'V1.0.0
'Vito
'接收到ERP的領/退料單據

Imports eCA_TransactionMessage
Imports eCA_HostObject

Module Module_SendWarehouseData
  Public Function O_SendWarehouseData(ByRef objSKU As MSG_SendWarehouseData,
                                     ByRef ret_strResultMsg As String) As Boolean

    Try
      '要變更的資料
      Dim dicAdd_SL As New Dictionary(Of String, clsSL)
      Dim dicUpdate_SL As New Dictionary(Of String, clsSL)
      Dim Warehouse = ""
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '檢查資料
      If Check_Data(objSKU, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(Warehouse, objSKU, ret_strResultMsg, dicAdd_SL, dicUpdate_SL) = False Then
        SendWarehouseData(enuRtnCode.Fail, Warehouse, "")
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, dicAdd_SL, dicUpdate_SL, lstSql) = False Then
        SendWarehouseData(enuRtnCode.Fail, Warehouse, "")
        Return False
      End If
      '執行SQL與更新物件
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
        SendWarehouseData(enuRtnCode.Fail, Warehouse, "")
        Return False
      End If
      SendWarehouseData(enuRtnCode.Sucess, Warehouse, "")
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Check_Data(ByVal objWarehouse As MSG_SendWarehouseData,
                                                          ByRef ret_strResultMsg As String) As Boolean
    Try
      If objWarehouse.EventID = "" Then
        ret_strResultMsg = "EventID is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      For Each objWarehouseDataInfo In objWarehouse.WarehouseDataList.WarehouseDataInfo
        '檢查SKU是否為空
        If objWarehouseDataInfo.Warehouse = "" Then
          ret_strResultMsg = "Warehouse is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查InventoryUnit是否為空
        If objWarehouseDataInfo.Owner = "" Then
          ret_strResultMsg = "Owner is empty"
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
  Private Function Get_UpdateData(ByRef ret_Warehouse As String, ByVal objWarehourse As MSG_SendWarehouseData,
                                                                  ByRef ret_strResultMsg As String,
                                                                  ByRef ret_dicAdd_SL As Dictionary(Of String, clsSL),
                                                                  ByRef ret_dicUpdate_SL As Dictionary(Of String, clsSL)) As Boolean
    Try
      Dim dicAdd_SKU As New Dictionary(Of String, clsSKU)
      'Dim dicUpdate_SKU As New Dictionary(Of String, clsSKU)
      '取得所有的SKU
      Dim tmp_dicSL As New Dictionary(Of String, clsSL)
      Dim User_ID As String = objWarehourse.WebService_ID
      Dim Event_ID As String = objWarehourse.EventID
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
      If gMain.objHandling.O_GetDB_dicSLByAll(tmp_dicSL) = False Then
        'ret_strResultMsg = "WMS get SKU data From DB Failed"
        'SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False
      End If

      For Each objWarehouseDataInfo In objWarehourse.WarehouseDataList.WarehouseDataInfo

        Dim Owner = objWarehouseDataInfo.Owner
        Dim Warehouse = objWarehouseDataInfo.Warehouse
        Dim WarehouseName = objWarehouseDataInfo.WarehouseName
        Dim Specification = objWarehouseDataInfo.Specification

        Dim objNewSL = New clsSL(Owner, Warehouse, Warehouse, WarehouseName, Specification, enuBND.None, enuQCStatus.NULL, 0, Now_Time, "")

        Dim tmp_objSL As clsSL = Nothing
        If tmp_dicSL.TryGetValue(objNewSL.gid, tmp_objSL) = True Then
          Dim objUpdateSL = tmp_objSL.Clone
          objUpdateSL.SL_ALIS = WarehouseName
          objUpdateSL.SL_DESC = Specification
          objUpdateSL.UPDATE_TIME = Now_Time
          If ret_dicUpdate_SL.ContainsKey(objUpdateSL.gid) = False Then
            ret_dicUpdate_SL.Add(objUpdateSL.gid, objUpdateSL)
          End If
        Else
          If ret_dicAdd_SL.ContainsKey(objNewSL.gid) = False Then
            ret_dicAdd_SL.Add(objNewSL.gid, objNewSL)
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
                           ByRef ret_dicAdd_SL As Dictionary(Of String, clsSL),
                           ByRef ret_dicUpdate_SL As Dictionary(Of String, clsSL),
                          ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得WMS_M_SL的SQL
      For Each obj In ret_dicAdd_SL.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert WMS_M_SL SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj In ret_dicUpdate_SL.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Update WMS_M_SL SQL Failed"
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
