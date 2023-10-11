''20180718
''V1.0.0
''Jerry
''Cosmo_WMS 回覆的過帳結果


'Module Module_Send_Outbound_Info

'  '過帳結果的回覆
'  Public Function O_Send_Outbound_Info(ByRef Receive_Msg As List(Of eCA_TransactionMessage.MSG_Outbound_Info), ByRef Result_Message As String) As Boolean
'    Try
'      '要新增的資料
'      Dim dicAdd_PO_IN_1_Header As New Dictionary(Of String, eCA_WMSObject.clsHost_PO_IN_1_Header)
'      Dim dicAdd_PO_IN_1_DTL As New Dictionary(Of String, eCA_WMSObject.clsHost_PO_IN_1_DTL)
'      Dim dicAdd_PO_DTL As New Dictionary(Of String, eCA_WMSObject.clsPO_DTL)
'      Dim dicAdd_PO As New Dictionary(Of String, eCA_WMSObject.clsPO)

'      '儲存要更新的SQL，進行一次性更新
'      Dim lstSql As New List(Of String)

'      '先進行資料邏輯檢查
'      If Check_Data(Receive_Msg, Result_Message) = False Then
'        Return False
'      End If
'      '取得要變更的資料
'      If Get_Data(Receive_Msg, Result_Message, dicAdd_PO_IN_1_Header, dicAdd_PO_IN_1_DTL, dicAdd_PO_DTL, dicAdd_PO) = False Then
'        Return False
'      End If
'      '取得要更新到DB的SQL
'      If Get_SQL(Result_Message, lstSql, dicAdd_PO_IN_1_Header, dicAdd_PO_IN_1_DTL, dicAdd_PO_DTL, dicAdd_PO) = False Then
'        Return False
'      End If
'      '執行資料更新
'      If Execute_DataUpdate(Result_Message, lstSql, dicAdd_PO_IN_1_Header, dicAdd_PO_IN_1_DTL, dicAdd_PO_DTL, dicAdd_PO) = False Then
'        Return False
'      End If
'      Return True


'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  '確認資料
'  Private Function Check_Data(ByRef Receive_Msg As List(Of eCA_TransactionMessage.MSG_Outbound_Info), ByRef Result_Message As String) As Boolean
'    Try
'      For Each Info In Receive_Msg
'        If Info.id = "" Then
'          Result_Message = "ID is null"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If Info.ll_no = "" Then
'          Result_Message = "ll_no is null"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If Info.factory_code = "" Then
'          Result_Message = "factory_code is null"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If Info.sap_factory_code = "" Then
'          Result_Message = "sap_factory_code is null"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If Info.material_code = "" Then
'          Result_Message = "material_code is null"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If Info.amount = "" Then
'          Result_Message = "amount is null"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If Info.send_spot = "" Then
'          Result_Message = "send_spot is null"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If Info.wkpos_code = "" Then
'          Result_Message = "wkpos_code is null"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        '檢查單據是否執行過
'        Dim NewojbPO_IN_Header_2 As eCA_WMSObject.clsHost_PO_IN_1_Header = Nothing
'        'Header的部份
'        If gMain.objWMS.O_Get_Host_PO_IN_1_Header(Info.ll_no, NewojbPO_IN_Header_2) Then
'          '如果存在
'          If NewojbPO_IN_Header_2.get_STEP_NO <> eCA_WMSObject.enuSTEP_NO.Queue Then
'            Result_Message = "單據ll_no : " & Info.ll_no & " 已執行，無法更新。"
'            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'            Return False
'          End If
'        End If
'      Next


'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False

'    End Try
'  End Function

'  '取得資訊 add、update
'  Private Function Get_Data(ByRef Receive_Msg As List(Of eCA_TransactionMessage.MSG_Outbound_Info), ByRef Result_Message As String,
'                            ByRef dicAdd_PO_IN_1_Header As Dictionary(Of String, eCA_WMSObject.clsHost_PO_IN_1_Header),
'                            ByRef dicAdd_PO_IN_1_DTL As Dictionary(Of String, eCA_WMSObject.clsHost_PO_IN_1_DTL),
'                            ByRef dicAdd_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                            ByRef dicAdd_PO As Dictionary(Of String, eCA_WMSObject.clsPO)) As Boolean
'    Try
'      For Each Info In Receive_Msg
'        '建立新的obj
'        Dim NewojbPO_IN_Header_2 As eCA_WMSObject.clsHost_PO_IN_1_Header = Nothing
'        Dim NewobjPO As eCA_WMSObject.clsPO = Nothing
'        Dim NewobjPO_IN_DTL_2 As eCA_WMSObject.clsHost_PO_IN_1_DTL = Nothing
'        Dim NewobjPO_DTL As eCA_WMSObject.clsPO_DTL = Nothing
'        NewojbPO_IN_Header_2 = New eCA_WMSObject.clsHost_PO_IN_1_Header(Info.ll_no, "", "", Info.supplier, "", "", "", "", "", "", "", "", "", "", "", "", "", Info.id, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ModuleHelpFunc.GetNewTime_DBFormat, "", eCA_WMSObject.enuSTEP_NO.Queue, 0)
'        NewobjPO = New eCA_WMSObject.clsPO(Info.ll_no, eCA_WMSObject.enuPOType.None, "50", ModuleHelpFunc.GetNewTime_DBFormat, "", "", "", "", "", "", Info.sap_factory_code, Info.send_spot, "", eCA_WMSObject.enuPOStatus.Queued, eCA_WMSObject.enuWOType.Receipt, "")
'        NewobjPO_IN_DTL_2 = New eCA_WMSObject.clsHost_PO_IN_1_DTL(Info.ll_no, 1, Info.amount, Info.sap_factory_code, Info.send_spot, Info.material_code, "", "", Info.factory_code, Info.id, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ModuleHelpFunc.GetNewTime_DBFormat)
'        NewobjPO_DTL = New eCA_WMSObject.clsPO_DTL(Info.ll_no, "1", Info.material_code, "", Info.amount, 0, 0, "", "", "", "", "", "", "", "", "", "", "", "")


'        '根據存在與否進行update或add
'        Dim ojbPO_IN_Header_2 As eCA_WMSObject.clsHost_PO_IN_1_Header = Nothing
'        Dim objPO As eCA_WMSObject.clsPO = Nothing
'        Dim objPO_IN_DTL_2 As eCA_WMSObject.clsHost_PO_IN_1_DTL = Nothing
'        Dim objPO_DTL As eCA_WMSObject.clsPO_DTL = Nothing

'        'Header的部份
'        If gMain.objWMS.O_Get_Host_PO_IN_1_Header(Info.ll_no, ojbPO_IN_Header_2) Then
'          '如果存在
'          Dim NewojbPO_IN_Header_2_Clone = ojbPO_IN_Header_2.Clone
'          NewojbPO_IN_Header_2_Clone.Update_To_Memory(NewojbPO_IN_Header_2) '更新狀態 '內容須微調
'          dicAdd_PO_IN_1_Header.Add(NewojbPO_IN_Header_2_Clone.get_gid, NewojbPO_IN_Header_2_Clone)
'        Else
'          '如果不存在
'          dicAdd_PO_IN_1_Header.Add(NewojbPO_IN_Header_2.get_gid, NewojbPO_IN_Header_2)
'        End If

'        If gMain.objWMS.O_Get_PO(Info.ll_no, objPO) Then
'          '如果存在
'          Dim NewobjPO_Clone = objPO.Clone
'          NewobjPO_Clone.Update_To_Memory(NewobjPO)
'          dicAdd_PO.Add(NewobjPO_Clone.get_gid, NewobjPO_Clone)
'        Else
'          '如果不存在
'          dicAdd_PO.Add(NewobjPO.get_gid, NewobjPO)
'        End If
'        'DTL的部分 (通常是對應的 不會一個有一個沒有)
'        If gMain.objWMS.O_Get_Host_PO_IN_1_DTL(Info.ll_no, "", objPO_IN_DTL_2) Then
'          '存在
'          Dim NewobjPO_IN_DTL_2_Clone = objPO_IN_DTL_2.Clone
'          NewobjPO_IN_DTL_2_Clone.Update_To_Memory(objPO_IN_DTL_2)
'          dicAdd_PO_IN_1_DTL.Add(NewobjPO_IN_DTL_2_Clone.get_gid, NewobjPO_IN_DTL_2_Clone)
'        Else
'          '不存在
'          dicAdd_PO_IN_1_DTL.Add(NewobjPO_IN_DTL_2.get_gid, NewobjPO_IN_DTL_2)
'        End If

'        If gMain.objWMS.O_Get_PO_DTL(Info.ll_no, 1, objPO_DTL) Then
'          '如果存在
'          Dim NewobjPO_DTL_Clone = objPO_DTL.Clone
'          NewobjPO_DTL_Clone.Update_To_Memory(NewobjPO_DTL)
'          dicAdd_PO_DTL.Add(NewobjPO_DTL_Clone.get_gid, NewobjPO_DTL_Clone)
'        Else
'          '如果不存在
'          dicAdd_PO_DTL.Add(NewobjPO_DTL.get_gid, NewobjPO_DTL)
'        End If


'      Next

'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False

'    End Try
'  End Function

'  '取得SQL語句
'  Private Function Get_SQL(ByRef Result_Message As String,
'                           ByRef lstSql As List(Of String),
'                           ByRef dicAdd_PO_IN_1_Header As Dictionary(Of String, eCA_WMSObject.clsHost_PO_IN_1_Header),
'                           ByRef dicAdd_PO_IN_1_DTL As Dictionary(Of String, eCA_WMSObject.clsHost_PO_IN_1_DTL),
'                           ByRef dicAdd_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                           ByRef dicAdd_PO As Dictionary(Of String, eCA_WMSObject.clsPO)) As Boolean
'    Try
'      '取得SQL
'      For Each objPO_IN_1_Header In dicAdd_PO_IN_1_Header
'        '存在的下update '不存在的下Insert
'        If gMain.objWMS.gdicHost_PO_IN_Check_Hist.ContainsKey(objPO_IN_1_Header.Key) Then
'          If objPO_IN_1_Header.Value.O_Add_Update_SQLString(lstSql) = False Then
'            Result_Message = "Get update Host_PO_IN_1_Header SQL Failed"
'            Return False
'          End If
'        Else
'          If objPO_IN_1_Header.Value.O_Add_Insert_SQLString(lstSql) = False Then
'            Result_Message = "Get Insert Host_PO_IN_1_Header SQL Failed"
'            Return False
'          End If
'        End If
'      Next

'      For Each objPO_IN_1_DTL In dicAdd_PO_IN_1_DTL
'        '存在的下update '不存在的下Insert
'        If gMain.objWMS.gdicHost_PO_IN_1_DTL.ContainsKey(objPO_IN_1_DTL.Key) Then
'          If objPO_IN_1_DTL.Value.O_Add_Update_SQLString(lstSql) = False Then
'            Result_Message = "Get update Host_PO_IN_1_DTL SQL Failed"
'            Return False
'          End If
'        Else
'          If objPO_IN_1_DTL.Value.O_Add_Insert_SQLString(lstSql) = False Then
'            Result_Message = "Get Insert Host_PO_IN_1_DTL SQL Failed"
'            Return False
'          End If
'        End If
'      Next

'      For Each objPO_DTL In dicAdd_PO_DTL
'        '存在的下update '不存在的下Insert
'        If gMain.objWMS.gdicPO_DTL.ContainsKey(objPO_DTL.Key) Then
'          If objPO_DTL.Value.O_Add_Update_SQLString(lstSql) = False Then
'            Result_Message = "Get update PO_DTL SQL Failed"
'            Return False
'          End If
'        Else
'          If objPO_DTL.Value.O_Add_Insert_SQLString(lstSql) = False Then
'            Result_Message = "Get Insert PO_DTL SQL Failed"
'            Return False
'          End If
'        End If
'      Next

'      For Each objPO In dicAdd_PO
'        '存在的下update '不存在的下Insert
'        If gMain.objWMS.gdicPO.ContainsKey(objPO.Key) Then
'          If objPO.Value.O_Add_Update_SQLString(lstSql) = False Then
'            Result_Message = "Get update PO SQL Failed"
'            Return False
'          End If
'        Else
'          If objPO.Value.O_Add_Insert_SQLString(lstSql) = False Then
'            Result_Message = "Get Insert PO SQL Failed"
'            Return False
'          End If
'        End If
'      Next

'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  '執行新增SQL語句，並進行記憶體資料更新
'  Private Function Execute_DataUpdate(ByRef Result_Message As String,
'                                      ByRef lstSql As List(Of String),
'                                      ByRef dicAdd_PO_IN_1_Header As Dictionary(Of String, eCA_WMSObject.clsHost_PO_IN_1_Header),
'                                      ByRef dicAdd_PO_IN_1_DTL As Dictionary(Of String, eCA_WMSObject.clsHost_PO_IN_1_DTL),
'                                      ByRef dicAdd_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                                      ByRef dicAdd_PO As Dictionary(Of String, eCA_WMSObject.clsPO)) As Boolean
'    Try
'      '更新所有的SQL
'      If eCA_WMSObject.Common_DBManagement.BatchUpdate_DynamicConnection(lstSql) = False Then
'        '更新DB失敗則回傳False
'        Result_Message = "WMS Update DB Failed"
'        Return False
'      End If
'      '修改記憶體資料 '更新或新增
'      For Each objHost_PD_Header In dicAdd_PO_IN_1_Header.Values
'        Dim objoldHost_PD_Header As eCA_WMSObject.clsHost_PO_IN_1_Header = Nothing
'        If gMain.objWMS.O_Get_Host_PO_IN_1_Header(objHost_PD_Header.get_PO_ID, objoldHost_PD_Header) Then
'          objoldHost_PD_Header.Update_To_Memory(objHost_PD_Header)
'        Else
'          objHost_PD_Header.Add_Relationship(gMain.objWMS)
'        End If
'      Next

'      For Each objHost_PD_DTL In dicAdd_PO_IN_1_DTL.Values
'        Dim objoldHost_PD_DTL As eCA_WMSObject.clsHost_PO_IN_1_DTL = Nothing
'        If gMain.objWMS.O_Get_Host_PO_IN_1_DTL(objHost_PD_DTL.get_PO_ID, objHost_PD_DTL.get_PO_SERIAL_NO, objoldHost_PD_DTL) Then
'          objoldHost_PD_DTL.Update_To_Memory(objHost_PD_DTL)
'        Else
'          objHost_PD_DTL.Add_Relationship(gMain.objWMS)
'        End If
'      Next

'      For Each objPD_DTL In dicAdd_PO_DTL.Values
'        Dim objoldPD_DTL As eCA_WMSObject.clsPO_DTL = Nothing
'        If gMain.objWMS.O_Get_PO_DTL(objPD_DTL.get_PO_ID, objPD_DTL.get_PO_SERIAL_NO, objoldPD_DTL) Then
'          objoldPD_DTL.Update_To_Memory(objPD_DTL)
'        Else
'          objPD_DTL.Add_Relationship(gMain.objWMS)
'        End If
'      Next

'      For Each objPD In dicAdd_PO.Values
'        Dim objoldPD As eCA_WMSObject.clsPO = Nothing
'        If gMain.objWMS.O_Get_PO(objPD.get_PO_ID, objoldPD) Then
'          objoldPD.Update_To_Memory(objPD)
'        Else
'          objPD.Add_Relationship(gMain.objWMS)
'        End If
'      Next


'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function


'End Module
