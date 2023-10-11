'20200608
'V1.0.0
'Vito
'接收到ERP的領/退料單據

Imports eCA_TransactionMessage
Imports eCA_HostObject

Module Module_SendSKUUnitConversionData
  Public Function O_POManagement_SendSKUUnitConversionData(ByRef objSKU As MSG_SendSKUUnitConversionData,
                                     ByRef ret_strResultMsg As String) As Boolean

    Try
      '要變更的資料
      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
      'Dim dicAdd_SKU As New Dictionary(Of String, clsSKU)
      Dim dicUpdate_SKU As New Dictionary(Of String, clsSKU)

      Dim dicAdd_Packe_Unit As New Dictionary(Of String, clsMPackeUnit)
      Dim dicUpdate_Packe_Unit As New Dictionary(Of String, clsMPackeUnit)
      Dim dicAdd_SKU_Packe_Structure As New Dictionary(Of String, clsMSKUPackeStructure)
      Dim dicUpdate_SKU_Packe_Structure As New Dictionary(Of String, clsMSKUPackeStructure)
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)

      '檢查資料
      If Check_Data(objSKU, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(objSKU, ret_strResultMsg, dicAdd_Packe_Unit, dicUpdate_Packe_Unit, dicAdd_SKU_Packe_Structure, dicUpdate_SKU_Packe_Structure, Host_Command, dicUpdate_SKU) = False Then
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, Host_Command, lstSql) = False Then
        Return False
      End If
      '執行SQL與更新物件
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Check_Data(ByVal objSKU As MSG_SendSKUUnitConversionData,
                                                          ByRef ret_strResultMsg As String) As Boolean
    Try
      If objSKU.EventID = "" Then
        ret_strResultMsg = "EventID is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      For Each objSKUDataInfo In objSKU.SKUUnitConversionDataList.SKUUnitConversionDataInfo
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
        '檢查ConversionUnit是否為空
        If objSKUDataInfo.ConversionUnit = "" Then
          ret_strResultMsg = "ConversionUnit is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查Molecule是否為空
        If objSKUDataInfo.Molecule = "" Then
          ret_strResultMsg = "Molecule is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        Else
          If IsNumeric(objSKUDataInfo.Molecule) = False Then
            ret_strResultMsg = "Molecule 不為數字"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        End If
        '檢查Denominator是否為空
        If objSKUDataInfo.Denominator = "" Then
          ret_strResultMsg = "Denominator is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        Else
          If IsNumeric(objSKUDataInfo.Denominator) = False Then
            ret_strResultMsg = "Denominator 不為數字"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        End If

      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '新增資料或得到要更新的資料
  Private Function Get_UpdateData(ByVal objSKU As MSG_SendSKUUnitConversionData,
                                                                  ByRef ret_strResultMsg As String,
                                                                  ByRef ret_dicAdd_Packe_Unit As Dictionary(Of String, clsMPackeUnit),
                                                                  ByRef ret_dicUpdate_Packe_Unit As Dictionary(Of String, clsMPackeUnit),
                                                                  ByRef ret_dicAdd_SKU_Packe_Structure As Dictionary(Of String, clsMSKUPackeStructure),
                                                                  ByRef ret_dicUpdate_SKU_Packe_Structure As Dictionary(Of String, clsMSKUPackeStructure),
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
      Dim Create_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()

      '使用dicPO取得資料庫裡的PO資料
      If gMain.objHandling.O_GetDB_dicSKUByAll(tmp_dicSKU) = False Then
        ret_strResultMsg = "WMS get SKU data From DB Failed"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If

      For Each objSKUDataInfo In objSKU.SKUUnitConversionDataList.SKUUnitConversionDataInfo



        Dim PACKE_UNIT = objSKUDataInfo.ConversionUnit
        Dim PACKE_UNIT_NAME = objSKUDataInfo.ConversionUnit
        Dim PACKE_UNIT_COMMON1 = ""
        Dim PACKE_UNIT_COMMON2 = ""
        Dim PACKE_UNIT_COMMON3 = ""
        Dim PACKE_UNIT_COMMON4 = ""
        Dim PACKE_UNIT_COMMON5 = ""
        Dim COMMENTS = ""
        'Dim CREATE_TIME = ""
        Dim UPDATE_TIME = ""

        Dim objPacke_Unit = New clsMPackeUnit(PACKE_UNIT, PACKE_UNIT_NAME, PACKE_UNIT_COMMON1, PACKE_UNIT_COMMON2, PACKE_UNIT_COMMON3, PACKE_UNIT_COMMON4, PACKE_UNIT_COMMON5, COMMENTS, Create_Time, UPDATE_TIME)

        If ret_dicAdd_Packe_Unit.ContainsKey(objPacke_Unit.gid) = False Then
          ret_dicAdd_Packe_Unit.Add(objPacke_Unit.gid, objPacke_Unit)
        End If

        Dim SKU_NO = objSKUDataInfo.SKU
        Dim PACKE_LV = 1
        'Dim PACKE_UNIT = ""
        Dim SUB_PACKE_UNIT = objSKUDataInfo.InventoryUnit
        Dim PACKE_WEIGHT = 0
        Dim PACKE_VOLUME = 0
        Dim PACKE_BCR = ""
        Dim OUT_MAX_UNIT = 0
        Dim IN_MAX_UNIT = 0
        Dim QTY = objSKUDataInfo.Molecule
        'Dim COMMENTS = ""
        'Dim CREATE_TIME = ""
        'Dim UPDATE_TIME = ""

        Dim objSKU_Packe_Structure = New clsMSKUPackeStructure(SKU_NO, PACKE_LV, PACKE_UNIT, SUB_PACKE_UNIT, PACKE_WEIGHT, PACKE_VOLUME, PACKE_BCR, OUT_MAX_UNIT,
                                                               IN_MAX_UNIT, QTY, COMMENTS, Create_Time, UPDATE_TIME)
        If ret_dicAdd_SKU_Packe_Structure.ContainsKey(objSKU_Packe_Structure.gid) = False Then
          ret_dicAdd_SKU_Packe_Structure.Add(objSKU_Packe_Structure.gid, objSKU_Packe_Structure)
        End If


        If ret_dicAdd_Packe_Unit.Any Then
          If Module_Send_WMSMessage.Send_T2F3U11_PackeUnitManagement_to_WMS(ret_strResultMsg, ret_dicAdd_Packe_Unit, Host_Command, enuAction.Create.ToString) = False Then
            Return False
          End If
        End If

        If ret_dicAdd_SKU_Packe_Structure.Any Then
          If Module_Send_WMSMessage.Send_T2F3U12_SKUPackeStructureManagement_to_WMS(ret_strResultMsg, ret_dicAdd_SKU_Packe_Structure, Host_Command, enuAction.Create.ToString) = False Then
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
