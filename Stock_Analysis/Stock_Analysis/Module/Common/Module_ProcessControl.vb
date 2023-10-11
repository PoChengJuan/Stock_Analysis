'20180619
'程式中用來挑選Process和控制Process運作

Module Module_ProcessControl
  '確認是否有執行中的ProcessFlow
  'Public Function O_Check_ProcessFlow(ByVal WO_ID As String, ByVal SKU_No As String) As Boolean
  '  Try
  '    If gMain.objWMS.O_Check_ProcessFlow(WO_ID, SKU_No) = True Then
  '      Return True
  '    Else
  '      Return False
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''取得原有的ProcessFlow的資料
  'Public Function O_Get_ProcessFlow(ByVal WO_ID As String, ByVal SKU_No As String) As eCA_WMSObject.clsProcessFlow
  '  Try
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = True Then
  '      Return objProcessFlow
  '    End If
  '    Return objProcessFlow
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return Nothing
  '  End Try
  'End Function
  ''確認是否有執行中的ProcessFlowDTL
  'Public Function O_Check_ProcessFlowDTL(ByVal WO_ID As String, ByVal SKU_No As String, ByVal Carrier_ID As String) As Boolean
  '  Try
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = O_Get_ProcessFlow(WO_ID, SKU_No)
  '    If objProcessFlow IsNot Nothing Then
  '      If objProcessFlow.O_Check_ProcessFlowDTL(objProcessFlow.get_Process_Flow(), Carrier_ID, SKU_No, WO_ID) = True Then
  '        Return True
  '      Else
  '        Return False
  '      End If
  '    Else
  '      Return False
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''取得原有的ProcessFlow的資料
  'Public Function O_Get_ProcessFlowDTL(ByVal WO_ID As String, ByVal SKU_No As String, ByVal Carrier_ID As String) As eCA_WMSObject.clsProcessFlowDTL
  '  Try
  '    Dim objProcessFlowDTL As eCA_WMSObject.clsProcessFlowDTL = Nothing
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = O_Get_ProcessFlow(WO_ID, SKU_No)
  '    If objProcessFlow IsNot Nothing Then
  '      If objProcessFlow.O_Get_ProcessFlowDTL(objProcessFlow.get_Process_Flow(), Carrier_ID, SKU_No, WO_ID, objProcessFlowDTL) = True Then
  '        Return objProcessFlowDTL
  '      End If
  '    End If
  '    Return objProcessFlowDTL
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return Nothing
  '  End Try
  'End Function

  ''-取得新的ProcessFlow
  ''-修改日期：2018/07/11 修改人：Mark
  'Public Function O_Get_New_ProcessFlow(ByVal WO_ID As String, ByVal SKU_No As String, ByVal Area_No As String, ByRef objProcessFlow As eCA_WMSObject.clsProcessFlow) As Boolean
  '  Try
  '    '建立ProcessFlow的資訊

  '    Dim objWO As eCA_WMSObject.clsWO = Nothing
  '    If gMain.objWMS.O_Get_WO(WO_ID, objWO) = True Then
  '      Dim WO_Type As eCA_WMSObject.enuWOType = objWO.get_WO_Type()
  '      Dim Factory_No As String = objWO.get_Factory_No()
  '      Dim Customer_No As String = objWO.get_Customer_No()
  '      '取得Customer_Type
  '      Dim Customer_Type = eCA_WMSObject.enuCustomerType.Null
  '      Dim objCustomer As eCA_WMSObject.clsCustomer = Nothing
  '      If gMain.objWMS.O_Get_Customer(Customer_No, objCustomer) = True Then
  '        Customer_Type = objCustomer.get_Customer_Type()
  '      End If
  '      '取得SKU_Catalog
  '      Dim SKU_Catalog = eCA_WMSObject.enuSKU_CATALOG.NULL
  '      Dim objSKU As eCA_WMSObject.clsSKU = Nothing
  '      If gMain.objWMS.O_Get_SKU(SKU_No, objSKU) = True Then
  '        SKU_Catalog = objSKU.get_SKU_CATALOG()
  '      End If
  '      '取得ProcessNo
  '      Dim lstAreaProcess As List(Of eCA_WMSObject.clsAreaProcess) = gMain.objWMS.GetAreaProcessByFactoryNo_AreaNo_SKUCatalog_CustomerType_WOType(Factory_No, Area_No, SKU_Catalog, Customer_No, WO_Type)
  '      If lstAreaProcess.Count > 0 Then
  '        Dim objAreaProcess As eCA_WMSObject.clsAreaProcess = lstAreaProcess(0)
  '        '加入新的ProcessFlow
  '        Dim PorcessNo As String = objAreaProcess.get_Process_No()
  '        Dim NewProcessFlowNo As String = O_Get_NewProcessFlowUUID(Factory_No, Area_No, SKU_No)
  '        Dim ProcessStatus = eCA_WMSObject.enuProcessStatus.Queued
  '        Dim Create_Time As String = GetNewTime_ByDataTimeFormat(GetNewTime_DBFormat)
  '        Dim Update_Time As String = Create_Time
  '        objProcessFlow = New eCA_WMSObject.clsProcessFlow(NewProcessFlowNo, SKU_No, WO_ID, Factory_No, Area_No, PorcessNo, ProcessStatus, Create_Time, Update_Time)
  '      Else
  '        '找不到對應的ProcessNo
  '        SendMessageToLog("can not find Process at WMS_M_Area_Process , Factory_No=" & Factory_No & ", Area_No=" & Area_No & ", SKU_Catalog=" & SKU_Catalog & ", Customer_No=" & Customer_No & ", WO_Type=" & WO_Type, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '    Else
  '      '找不到對應的WO_ID
  '      SendMessageToLog("can not find WO_ID at WMS_T_WO , WO_ID=" & WO_ID, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function

  ''-取得新的ProcessFlow
  ''-修改日期：2018/07/11 修改人：Mark
  'Public Function O_Get_New_ProcessFlow(ByVal WO_ID As String, ByVal WO_Type As eCA_WMSObject.enuWOType,
  '                                      ByVal SKU_No As String, ByVal Factory_No As String, ByVal Area_no As String,
  '                                      ByVal Customer_No As String, ByRef objProcessFlow As eCA_WMSObject.clsProcessFlow) As Boolean
  '  Try
  '    '建立ProcessFlow的資訊
  '    '取得Customer_Type
  '    Dim Customer_Type = eCA_WMSObject.enuCustomerType.Null
  '    Dim objCustomer As eCA_WMSObject.clsCustomer = Nothing
  '    If gMain.objWMS.O_Get_Customer(Customer_No, objCustomer) = True Then
  '      Customer_Type = objCustomer.get_Customer_Type()
  '    End If
  '    '取得SKU_Catalog
  '    Dim SKU_Catalog = eCA_WMSObject.enuSKU_CATALOG.NULL
  '    Dim objSKU As eCA_WMSObject.clsSKU = Nothing
  '    If gMain.objWMS.O_Get_SKU(SKU_No, objSKU) = True Then
  '      SKU_Catalog = objSKU.get_SKU_CATALOG()
  '    End If
  '    '取得ProcessNo
  '    Dim lstAreaProcess As List(Of eCA_WMSObject.clsAreaProcess) = gMain.objWMS.GetAreaProcessByFactoryNo_AreaNo_SKUCatalog_CustomerType_WOType(Factory_No, Area_no, SKU_Catalog, Customer_No, WO_Type)
  '    If lstAreaProcess.Count > 0 Then
  '      Dim objAreaProcess As eCA_WMSObject.clsAreaProcess = lstAreaProcess(0)
  '      '加入新的ProcessFlow
  '      Dim PorcessNo As String = objAreaProcess.get_Process_No()
  '      Dim NewProcessFlowNo As String = O_Get_NewProcessFlowUUID(Factory_No, Area_no, SKU_No)
  '      Dim ProcessStatus = eCA_WMSObject.enuProcessStatus.Queued
  '      Dim Create_Time As String = GetNewTime_ByDataTimeFormat(GetNewTime_DBFormat)
  '      Dim Update_Time As String = Create_Time
  '      objProcessFlow = New eCA_WMSObject.clsProcessFlow(NewProcessFlowNo, SKU_No, WO_ID, Factory_No, Area_no, PorcessNo, ProcessStatus, Create_Time, Update_Time)
  '    Else
  '      '找不到對應的ProcessNo
  '      SendMessageToLog("can not find Process at WMS_M_Area_Process , Factory_No=" & Factory_No & ", Area_No=" & Area_no & ", SKU_Catalog=" & SKU_Catalog & ", Customer_No=" & Customer_No & ", WO_Type=" & WO_Type, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function

  ''-取得新的Process_Flow_DTL
  ''-修改日期：2018/07/11 修改人：Mark
  'Public Function O_Get_New_ProcessFlowDTL(ByRef objPrcoessFlow As eCA_WMSObject.clsProcessFlow, ByVal Carrier_ID As String,
  '                                         ByRef objProcessFlowDTL As eCA_WMSObject.clsProcessFlowDTL) As Boolean
  '  Try
  '    Dim ProcessFlow_No As String = objPrcoessFlow.get_Process_Flow()
  '    Dim SKU_No As String = objPrcoessFlow.get_SKU_NO()
  '    Dim WO_ID As String = objPrcoessFlow.get_WO_ID()
  '    Dim Process_No As String = objPrcoessFlow.get_Process_NO()
  '    '取得Process的資料
  '    Dim objProcess As eCA_WMSObject.clsProcess = Nothing
  '    If gMain.objWMS.O_Get_Process(Process_No, objProcess) = True Then
  '      '取得ProcessDTL的資料
  '      Dim objProcessDTL As eCA_WMSObject.clsProcessDTL = Nothing
  '      Dim Sub_Process_Seq As Long = 1
  '      If objProcess.O_Get_ProcessDTL(Process_No, Sub_Process_Seq, objProcessDTL) = True Then
  '        '取得SubProcess的資料
  '        Dim objSubProcess As eCA_WMSObject.clsSubProcess = Nothing
  '        Dim Sub_Process_No As String = objProcessDTL.get_Sub_Process_NO()
  '        If gMain.objWMS.O_Get_SubProcess(Sub_Process_No, objSubProcess) = True Then
  '          '取得SubProcess_DTL的資料
  '          Dim objSubProcessDTL As eCA_WMSObject.clsSubProcessDTL = Nothing
  '          Dim Step_Seq As Long = 1
  '          If objSubProcess.O_Get_SubProcessDTL(Sub_Process_No, Step_Seq, objSubProcessDTL) = True Then
  '            Dim Step_No As String = objSubProcessDTL.get_Step_NO()
  '            Dim Step_Result As String = ""
  '            Dim SubProcess_Status = eCA_WMSObject.enuSubProcessStatus.Queued
  '            Dim Create_Time As String = GetNewTime_ByDataTimeFormat(GetNewTime_DBFormat)
  '            Dim Update_Time As String = Create_Time
  '            objProcessFlowDTL = New eCA_WMSObject.clsProcessFlowDTL(ProcessFlow_No, SKU_No, WO_ID, Carrier_ID, Sub_Process_No, Step_Seq, Step_No, Step_Result, SubProcess_Status, Create_Time, Update_Time)
  '            Return True
  '          Else
  '            '找不到對應的Sub_Process_No的Step_Seq
  '            SendMessageToLog("can not find Sub_Process_No at WMS_M_Sub_Process_DTL , Sub_Process_No=" & Sub_Process_No & ", Step_Seq=" & Step_Seq, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '            Return False
  '          End If
  '        Else
  '          '找不到對應的Sub_Process_No
  '          SendMessageToLog("can not find Sub_Process_No at WMS_M_Sub_Process , Sub_Process_No=" & Sub_Process_No, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '          Return False
  '        End If
  '      Else
  '        '找不到對應的Process_No的Sub_Process_Seq
  '        SendMessageToLog("can not find Process_No at WMS_M_Process_DTL , Process_No=" & Process_No & ", Sub_Process_Seq=" & Sub_Process_Seq, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '    Else
  '      '找不到對應的Process_No
  '      SendMessageToLog("can not find Process_No at WMS_M_Process , Process_No=" & Process_No, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return False
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function


  ''-Get New Process UUID
  ''-取得新的ProcessFlow的流水號
  ''-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Get_NewProcessFlowUUID(ByVal Factory_No As String, ByVal Area_No As String, ByVal SKU_No As String) As String
  '  Try
  '    Dim ProcessFlow As String = ""
  '    Dim SKU_Catalog As String = ""
  '    Dim Modify_date As String = GetNewTime_ByDataTimeFormat(DBDate_IDFormat)
  '    Dim objProcessFlowUUID As eCA_WMSObject.clsProcessFlowUUID = Nothing
  '    Dim UUID As Long = 0
  '    '檢查SKU是否存在，並取得SKU_Catalog
  '    Dim objSKU As eCA_WMSObject.clsSKU = Nothing
  '    If gMain.objWMS.O_Get_SKU(SKU_No, objSKU) = False Then
  '      SendMessageToLog("Get SKU Data Failed, key not exists key=" & SKU_No, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return ""
  '    End If
  '    SKU_Catalog = objSKU.get_SKU_CATALOG()
  '    '使用Factory_No和Area_No和SKU_No取得ProcessFlow的流水號
  '    If gMain.objWMS.O_Get_ProcessFlowUUID(SKU_Catalog, Factory_No, Area_No, objProcessFlowUUID) = False Then
  '      '不存在該SKU_Catalog、Factory_No、Area_No的資料
  '      '建立一筆資料到ProcessFlowUUID
  '      objProcessFlowUUID = New eCA_WMSObject.clsProcessFlowUUID(SKU_Catalog, Factory_No, Area_No, Modify_date, UUID)
  '      If objProcessFlowUUID.O_Insert_ProcessFlowUUID_To_DB(gMain.objWMS) = True Then
  '      Else
  '        SendMessageToLog("Insert Process Flow UUID Failed ,SKU_Catalog=" & objProcessFlowUUID.get_SKU_Catalog() & " ,Factory_No=" & objProcessFlowUUID.get_Factory_No() & " ,Area_No=" & objProcessFlowUUID.get_Area_No(), eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return ""
  '      End If
  '    End If
  '    ProcessFlow = objProcessFlowUUID.Get_NewProcessFlow()
  '    Return ProcessFlow
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return ""
  '  End Try
  'End Function
  '-Create WMS_T_Process_Flow
  '-於資料庫建立WMS_T_Process_Flow，並寫入記憶體
  '-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Create_ProcessFlow(ByVal ProcessFlow As String, ByVal SKU_No As String, ByVal WO_ID As String,
  '                                  ByVal Process_No As String) As Boolean
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    If gMain.objWMS.O_Check_ProcessFlow(WO_ID, SKU_No) = False Then
  '      Dim Process_Status As eCA_WMSObject.enuProcessStatus = eCA_WMSObject.enuProcessStatus.Queued
  '      Dim Create_Time As String = GetNewTime_ByDataTimeFormat(DBTimeFormat)
  '      Dim Update_Time As String = Create_Time
  '      Dim objProcessFlow As New eCA_WMSObject.clsProcessFlow(ProcessFlow, SKU_No, WO_ID, Process_No, Process_Status, Create_Time, Update_Time)
  '      '在資料庫Insert ProcessFlow的資料
  '      If objProcessFlow.O_Insert_ProcessFlow_To_DB(gMain.objWMS) = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("Insert Process Flow Failed ,ProcessFlow=" & objProcessFlow.get_Process_Flow(), eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If
  '    Else
  '      SendMessageToLog("Process Flow already exists ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return False
  '    End If
  '    Return False
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''-Update WMS_T_Process_Flow
  ''-於資料庫更新WMS_T_Process_Flow，並寫入記憶體
  ''-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Update_ProcessFlowStatus(ByVal SKU_No As String, ByVal WO_ID As String,
  '                                  ByVal Process_Status As eCA_WMSObject.enuProcessStatus) As Boolean
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = True Then
  '      If objProcessFlow.O_Update_ProcessStatus_To_DB(Process_Status) = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("Update Process Status Failed ,ProcessFlow=" & objProcessFlow.get_Process_Flow(), eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If
  '    Else
  '      SendMessageToLog("Process Flow not exist ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return False
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''-Delete WMS_T_Process_Flow
  ''-於資料庫刪除WMS_T_Process_Flow，並刪除記憶體關聯
  ''-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Delete_ProcessFlow(ByVal SKU_No As String, ByVal WO_ID As String) As Boolean
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = True Then
  '      If objProcessFlow.O_Delete_ProcessFlow_To_DB() = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("Delete Process Flow Failed ,ProcessFlow=" & objProcessFlow.get_Process_Flow(), eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If
  '    Else
  '      SendMessageToLog("Process Flow not exist ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return False
  '    End If
  '    Return False
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''-Create WMS_T_Process_Flow_DTL
  ''-於資料庫建立WMS_T_Process_Flow_DTL，並寫入記憶體
  ''-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Create_ProcessFlowDTL(ByVal ProcessFlow As String, ByVal SKU_No As String, ByVal WO_ID As String,
  '                                     ByVal Carrier_ID As String, ByVal Sub_Process_No As String, ByVal Current_Step_Seq As String,
  '                                     ByVal Current_Step_No As String) As Boolean
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = False Then
  '      Return False
  '    End If
  '    '檢查ProcessFlowDTL是否已經存在
  '    If objProcessFlow.O_Check_ProcessFlowDTL(ProcessFlow, Carrier_ID, SKU_No, WO_ID) = False Then
  '      Dim SubProcess_Status As eCA_WMSObject.enuSubProcessStatus = eCA_WMSObject.enuSubProcessStatus.Queued
  '      Dim Step_Result As String = ""
  '      Dim Create_Time As String = GetNewTime_ByDataTimeFormat(DBTimeFormat)
  '      Dim Update_Time As String = Create_Time

  '      Dim objProcessFlowDTL As New eCA_WMSObject.clsProcessFlowDTL(ProcessFlow, SKU_No, WO_ID, Carrier_ID, Sub_Process_No, Current_Step_Seq, Current_Step_No, Step_Result, SubProcess_Status, Create_Time, Update_Time)
  '      '在資料庫Insert ProcessFlowDTL的資料
  '      If objProcessFlowDTL.O_Insert_ProcessFlowDTL_To_DB(gMain.objWMS) = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("Insert Process Flow DTL Failed ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If

  '    Else
  '      SendMessageToLog("Process Flow DTL already exists ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return False
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''-Update WMS_T_Process_Flow_DTL 的 Step_Result和SubProcessStatus
  ''-於資料庫更新WMS_T_Process_Flow_DTL，並寫入記憶體
  ''-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Update_ProcessFlowDTLSubProcessStatus(ByVal ProcessFlow As String, ByVal SKU_No As String, ByVal WO_ID As String, ByVal Carrier_ID As String,
  '                                   ByVal Step_Result As String, ByVal SubProcess_Status As eCA_WMSObject.enuSubProcessStatus) As Boolean
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = False Then
  '      Return False
  '    End If
  '    '檢查ProcessFlowDTL是否已經存在
  '    Dim objProcessFlowDTL As eCA_WMSObject.clsProcessFlowDTL = Nothing
  '    If objProcessFlow.O_Get_ProcessFlowDTL(ProcessFlow, Carrier_ID, SKU_No, WO_ID, objProcessFlowDTL) = True Then
  '      If objProcessFlowDTL.O_Update_SubProcessStatus_To_DB(Step_Result, SubProcess_Status) = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("Insert Process Flow DTL Failed ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If
  '    Else
  '      SendMessageToLog("Process Flow DTL not exist ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return False
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''-Update WMS_T_Process_Flow_DTL 的 Current_Step_Seq、Cuerrent_Step_No、Step_Result和SubProcessStatus
  ''-於資料庫更新WMS_T_Process_Flow_DTL，並寫入記憶體
  ''-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Update_ProcessFlowDTLStep(ByVal ProcessFlow As String, ByVal SKU_No As String, ByVal WO_ID As String, ByVal Carrier_ID As String,
  '                                   ByVal Step_Seq As Long) As Boolean
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = False Then
  '      Return False
  '    End If
  '    '檢查ProcessFlowDTL是否已經存在
  '    Dim objProcessFlowDTL As eCA_WMSObject.clsProcessFlowDTL = Nothing
  '    If objProcessFlow.O_Get_ProcessFlowDTL(ProcessFlow, Carrier_ID, SKU_No, WO_ID, objProcessFlowDTL) = True Then
  '      Dim Step_No As String = O_Get_SubProcessDTL_StepNo(objProcessFlowDTL.get_Sub_Process_NO, Step_Seq)
  '      If Step_No <> "" Then
  '        If objProcessFlowDTL.O_Update_CurrentStep_To_DB(Step_Seq, Step_No) = True Then
  '          Return True
  '        Else
  '          SendMessageToLog("Insert Process Flow DTL Failed ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '          Return False
  '        End If
  '      Else
  '        SendMessageToLog("Can not get Step_No ,Sub_Process_NO=" & objProcessFlowDTL.get_Sub_Process_NO() & " ,Step_Seq=" & Step_Seq, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '    Else
  '      SendMessageToLog("Process Flow DTL not exist ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return False
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''-Update WMS_T_Process_Flow_DTL 的 Sub_Process_No、Current_Step_Seq、Cuerrent_Step_No、Step_Result和SubProcessStatus
  ''-於資料庫更新WMS_T_Process_Flow_DTL，並寫入記憶體
  ''-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Update_ProcessFlowDTLSubProcessNo(ByVal ProcessFlow As String, ByVal SKU_No As String, ByVal WO_ID As String, ByVal Carrier_ID As String,
  '                                   ByVal Sub_Process_No As String) As Boolean
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = False Then
  '      Return False
  '    End If
  '    '檢查ProcessFlowDTL是否已經存在
  '    Dim objProcessFlowDTL As eCA_WMSObject.clsProcessFlowDTL = Nothing
  '    If objProcessFlow.O_Get_ProcessFlowDTL(ProcessFlow, Carrier_ID, SKU_No, WO_ID, objProcessFlowDTL) = True Then
  '      Dim Step_Seq As Long = 1
  '      Dim Step_No As String = O_Get_SubProcessDTL_StepNo(Sub_Process_No, Step_Seq)
  '      If Step_No <> "" Then
  '        If objProcessFlowDTL.O_Update_SubProcess_To_DB(Sub_Process_No, Step_Seq, Step_No) = True Then
  '          Return True
  '        Else
  '          SendMessageToLog("Update Process Flow DTL Failed ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '          Return False
  '        End If
  '      Else
  '        SendMessageToLog("Can not get Step_No ,Sub_Process_NO=" & objProcessFlowDTL.get_Sub_Process_NO() & " ,Step_Seq=" & Step_Seq, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '        Return False
  '      End If
  '    Else
  '      SendMessageToLog("Process Flow DTL not exist ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return False
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''-Delete WMS_T_Process_Flow
  ''-於資料庫刪除WMS_T_Process_Flow，並刪除記憶體關聯
  ''-修改日期：2018/06/20 修改人：Mark
  'Public Function O_Delete_ProcessFlowDTL(ByVal ProcessFlow As String, ByVal SKU_No As String, ByVal WO_ID As String, ByVal Carrier_ID As String) As Boolean
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = False Then
  '      Return False
  '    End If
  '    '檢查ProcessFlowDTL是否已經存在
  '    Dim objProcessFlowDTL As eCA_WMSObject.clsProcessFlowDTL = Nothing
  '    If objProcessFlow.O_Get_ProcessFlowDTL(ProcessFlow, Carrier_ID, SKU_No, WO_ID, objProcessFlowDTL) = True Then
  '      Dim Update_Time As String = GetNewTime_ByDataTimeFormat(DBTimeFormat)
  '      If objProcessFlowDTL.O_Delete_ProcessFlowDTL_To_DB() = True Then
  '        Return True
  '      Else
  '        SendMessageToLog("Delete Process Flow DTL Failed ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '      End If
  '    Else
  '      SendMessageToLog("Process Flow DTL not exist ,ProcessFlow=" & ProcessFlow & " ,WO_ID=" & WO_ID & " ,SKU_No=" & SKU_No & " ,Carrier_ID=" & Carrier_ID, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Return False
  '    End If
  '    Return False
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  ''-Get ProcessFlowDTL Next_Step_Seq
  ''-取得ProcessFlowDTL的下一個Step_Seq
  ''-修改日期：2018/06/20 修改人：Mark
  ''回傳-1表示程式錯誤
  ''回傳0表示以經到了最後一個Step
  'Public Function O_Get_ProcessFlowDTL_NextStepSeq(ByVal ProcessFlow As String, ByVal WO_ID As String, ByVal SKU_No As String, ByVal Carrier_ID As String) As Long
  '  Try
  '    '先檢查ProcessFlow是否已經存在
  '    Dim objProcessFlow As eCA_WMSObject.clsProcessFlow = Nothing
  '    If gMain.objWMS.O_Get_ProcessFlow(WO_ID, SKU_No, objProcessFlow) = False Then
  '      Return -1
  '    End If
  '    '檢查ProcessFlowDTL是否已經存在
  '    Dim objProcessFlowDTL As eCA_WMSObject.clsProcessFlowDTL = Nothing
  '    If objProcessFlow.O_Get_ProcessFlowDTL(ProcessFlow, Carrier_ID, SKU_No, WO_ID, objProcessFlowDTL) = False Then
  '      Return -1
  '    End If
  '    Dim objSubProcess As eCA_WMSObject.clsSubProcess = Nothing
  '    Dim objSubProcessDTL As eCA_WMSObject.clsSubProcessDTL = Nothing
  '    Dim objSubProcessDTLCTRL As eCA_WMSObject.clsSubProcessDTLCTRL = Nothing
  '    Dim SubProcess_No As String = objProcessFlowDTL.get_Sub_Process_NO()
  '    Dim Currnet_Step_Seq As Long = objProcessFlowDTL.get_Current_Step_Seq()
  '    Dim Step_Result As String = objProcessFlowDTL.get_Step_Result()
  '    Dim Next_Step_Seq As Long = 0
  '    Dim blnGetSubPrpcessDTLCTRL As Boolean = False
  '    '檢查SubProcess是否已經存在
  '    If gMain.objWMS.O_Get_SubProcess(SubProcess_No, objSubProcess) = False Then
  '      Return -1
  '    End If
  '    '檢查SubProcessDTL是否已經存在
  '    If objSubProcess.O_Get_SubProcessDTL(SubProcess_No, Currnet_Step_Seq, objSubProcessDTL) = False Then
  '      Return -1
  '    End If
  '    '檢查SubProcess是否需要StepResult
  '    Dim objStep As eCA_WMSObject.clsStep = Nothing
  '    If gMain.objWMS.O_Get_Step(objSubProcessDTL.get_Step_NO(), objStep) = True Then
  '      '如果Step Flag是True，則檢查SubProcessDTLCTRL，否則Step+1
  '      If objStep.get_Step_Flag() = True Then
  '        '檢查SubProcessDTLCTRL是否已經存在
  '        If objSubProcessDTL.O_Get_SubProcessDTLCTRL(SubProcess_No, Currnet_Step_Seq, Step_Result, objSubProcessDTLCTRL) = True Then
  '          Next_Step_Seq = objSubProcessDTLCTRL.get_Next_Step_Seq()
  '          blnGetSubPrpcessDTLCTRL = True
  '        Else
  '          Next_Step_Seq = Currnet_Step_Seq + 1
  '          SendMessageToLog("Sub Process DTL CTRL not exists, SubProcess_No=" & SubProcess_No & ", Currnet_Step_Seq=" & Currnet_Step_Seq & ", Step_Result=" & Step_Result, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
  '        End If
  '      Else
  '        Next_Step_Seq = Currnet_Step_Seq + 1
  '      End If
  '    End If
  '    '取得新的Next_Step_Seq後，再回來檢查新的Next_Step是否有在設定內
  '    If objSubProcess.O_Get_SubProcessDTL(SubProcess_No, Next_Step_Seq, objSubProcessDTL) = False Then
  '      '新的Next_Step不在設定內
  '      If blnGetSubPrpcessDTLCTRL = True Then
  '        '如果是從SubProcessDTLCTRL中取得的，但卻不在設定內則回傳錯誤
  '        SendMessageToLog("Next Step not at SubProcessDTL, SubProcess_No=" & SubProcess_No & ", Next_Step_Seq=" & Next_Step_Seq, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
  '        Return -1
  '      Else
  '        '由Current_Step_Seq加一後取得的，不在設定內，表示可能是最後一步了
  '        Return 0
  '      End If
  '    Else
  '      Return Next_Step_Seq
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return -1
  '  End Try
  'End Function
  ''-Get StepNo From SubProcessDTL
  ''-從SubProcessDTL取得StepNo
  ''-修改日期：2018/06/20 修改人：Mark
  'Private Function O_Get_SubProcessDTL_StepNo(ByVal Sub_Process_No As String, ByVal step_Seq As Long) As String
  '  Try
  '    Dim objSubProcess As eCA_WMSObject.clsSubProcess = Nothing
  '    If gMain.objWMS.O_Get_SubProcess(Sub_Process_No, objSubProcess) = False Then
  '      SendMessageToLog("SubProcess not exists Sub_Process_No=" & Sub_Process_No, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return ""
  '    End If
  '    Dim objSubProcessDTL As eCA_WMSObject.clsSubProcessDTL = Nothing
  '    If objSubProcess.O_Get_SubProcessDTL(Sub_Process_No, step_Seq, objSubProcessDTL) = False Then
  '      SendMessageToLog("SubProcessDTL not exists Sub_Process_No=" & Sub_Process_No & ", Step_Seq=" & step_Seq, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      Return ""
  '    End If
  '    Dim Step_No As String = objSubProcessDTL.get_Step_NO()
  '    Return Step_No
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return ""
  '  End Try
  'End Function





End Module
