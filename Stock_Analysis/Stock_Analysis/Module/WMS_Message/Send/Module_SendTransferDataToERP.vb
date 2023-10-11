''20180718
''V1.0.0
''Jerry
''接收到ERP的製令單


'Module Module_SendTransferDataToERP
'  Public Function O_SendTransferDataToERP(ByRef objSendTransferDataToERP As eCA_TransactionMessage.MSG_SendTransferDataToERP,
'                                      ByRef Result_Message As String) As Boolean

'    Try
'      ''儲存要更新的SQL，進行一次性更新
'      Dim lstSql As New List(Of String)

'      '要變更的資料
'      Dim dicAdd_PO_DTL As New Dictionary(Of String, eCA_WMSObject.clsPO_DTL)
'      Dim dicAdd_PO As New Dictionary(Of String, eCA_WMSObject.clsPO)
'      Dim dicReAdd_PO_DTL As New Dictionary(Of String, eCA_WMSObject.clsPO_DTL)
'      Dim dicReAdd_PO As New Dictionary(Of String, eCA_WMSObject.clsPO)


'      '若存在且未執行則更新，若不存在則新增
'      '並提單取回資料
'      If _CheckPO_Exist(objSendTransferDataToERP, Result_Message) = False Then
'        Return False '有單且已執行 則無法提單
'      End If

'      '根據單據的有無做出對應的處理 '有則更新 '無則新增
'      If _Get_Data(objSendTransferDataToERP, Result_Message, dicAdd_PO_DTL, dicAdd_PO, dicReAdd_PO_DTL, dicReAdd_PO) = False Then
'        Return False '提單失敗
'      End If

'      '取得SQL
'      If _Get_SQL(Result_Message, dicAdd_PO_DTL, dicAdd_PO, dicReAdd_PO_DTL, dicReAdd_PO, lstSql) = False Then
'        Return False
'      End If

'      '執行SQL與更新物件
'      If _Execute_DataUpdate(Result_Message, dicAdd_PO_DTL, dicAdd_PO, dicReAdd_PO_DTL, dicReAdd_PO, lstSql) = False Then
'        Return False
'      End If


'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  '檢查單據是否存在 若存在則更新 不存在則新增
'  '回傳存不存在
'  Private Function _CheckPO_Exist(ByVal objSendWorkData As eCA_TransactionMessage.MSG_SendTransferDataToERP, ByRef Result_Message As String) As Boolean
'    Try
'      '檢查是否存在
'      For Each WrokDataInfo In objSendWorkData.Data.FormBody.RecordList
'        Dim objPO_ID As eCA_WMSObject.clsPO = Nothing
'        If gMain.objWMS.O_Get_PO(WrokDataInfo.TC005, objPO_ID) Then
'          If objPO_ID.get_PO_STATUS <> eCA_WMSObject.enuPOStatus.Queued Then
'            Result_Message = "單據:" & WrokDataInfo.TC005 & " 已執行，無法更新"
'            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'            Return False
'          End If
'        End If
'      Next

'      Return True
'    Catch ex As Exception
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  '新增資料或得到要更新的資料
'  Private Function _Get_Data(ByVal objSendWorkData As eCA_TransactionMessage.MSG_SendTransferDataToERP, ByRef Result_Message As String,
'                             ByRef dicAdd_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                             ByRef dicAdd_PO As Dictionary(Of String, eCA_WMSObject.clsPO),
'                             ByRef dicUpdate_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                             ByRef dicUpdate_PO As Dictionary(Of String, eCA_WMSObject.clsPO)) As Boolean
'    Try


'      For Each WrokDataInfo In objSendWorkData.Data.FormBody.RecordList

'        'Dim TC003 '加工順序
'        Dim TC004 = WrokDataInfo.TC004     '製令單別
'        Dim TC005 = WrokDataInfo.TC005  '製令單號
'        Dim TC006 = WrokDataInfo.TC006  '移出工序
'        Dim TC007 = WrokDataInfo.TC007  '移出製程
'        Dim TC008 = WrokDataInfo.TC008 '移入工序
'        Dim TC009 = WrokDataInfo.TC009  '移入製程
'        Dim TC010 = WrokDataInfo.TC010 '單位
'        Dim TC014 = WrokDataInfo.TC014  '驗收數量
'        Dim TC016 = WrokDataInfo.TC016 '報廢數量
'        Dim TC020 = WrokDataInfo.TC020  '使用人時
'        Dim TC021 = WrokDataInfo.TC021 '使用機時

'        '單據存在要更新
'        Dim objPO As eCA_WMSObject.clsPO = Nothing
'        Dim objPO_DTL As eCA_WMSObject.clsPO_DTL = Nothing
'        '取全部單據 '製令單沒有項次 所以都填 1
'        If gMain.objWMS.O_Get_PO(TC005, objPO) = False Or gMain.objWMS.O_Get_PO_DTL(TC005, "1", objPO_DTL) = False Then
'          '無單據 要新增
'          '新增PO_DTL
'          Dim New_PO_DTL = New eCA_WMSObject.clsPO_DTL(TC005, "1", 1, TA006, TA063, TA015, 0, 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ModuleHelpFunc.GetNewTime_DBFormat, "", TA007, TA013, TA021, TA030, TA063, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
'          dicAdd_PO_DTL.Add(New_PO_DTL.get_gid, New_PO_DTL)

'          '新增PO
'          Dim New_PO = New eCA_WMSObject.clsPO(TC005, TC004, "", "", 50, ModuleHelpFunc.GetNewTime_DBFormat, "", "", "", "", "", "", "", "", "", eCA_WMSObject.enuPOStatus.Queued, eCA_WMSObject.enuWOType.Discharge, "", ModuleHelpFunc.GetNewTime_DBFormat, "", eCA_WMSObject.enuSTEP_NO.Queue, eCA_WMSObject.enuOrderType.MOCAB01_SENDWORKDATA, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", TA007, TA013, TA021, TA063, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
'          dicAdd_PO.Add(New_PO.get_gid, New_PO)

'        Else
'          '有單據 要更新
'          Dim NewobjPO_DTL = objPO_DTL.Clone
'          NewobjPO_DTL.Update_To_Memory_ByModule_MOCAB01_SENDWORKDATA(TA006, TA063, TA015, TA007, TA011, TA013, TA021, TA030)
'          dicUpdate_PO_DTL.Add(NewobjPO_DTL.get_gid, NewobjPO_DTL)

'          Dim NewobjPO = objPO.Clone
'          NewobjPO.Update_To_Memory_ByMOCAB01_SENDWORKDATA(TC004, TA030, TA007, TA013, TA021, TA063)
'          dicUpdate_PO.Add(NewobjPO.get_gid, NewobjPO)

'        End If
'      Next



'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function


'  'SQL
'  Private Function _Get_SQL(ByRef Result_Message As String,
'                             ByRef dicAdd_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                             ByRef dicAdd_PO As Dictionary(Of String, eCA_WMSObject.clsPO),
'                             ByRef dicReAdd_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                             ByRef dicReAdd_PO As Dictionary(Of String, eCA_WMSObject.clsPO),
'                             ByRef lstSql As List(Of String)) As Boolean
'    Try
'      For Each values In dicReAdd_PO_DTL.Values
'        If values.O_Add_Delete_SQLString(lstSql) = False Then
'          Result_Message = "Get Delete PO_DTL SQL Failed"
'          Return False
'        End If
'        If values.O_Add_Insert_SQLString(lstSql) = False Then
'          Result_Message = "Get Insert PO_DTL SQL Failed"
'          Return False
'        End If
'      Next
'      For Each values In dicReAdd_PO.Values
'        If values.O_Add_Delete_SQLString(lstSql) = False Then
'          Result_Message = "Get Delete PO SQL Failed"
'          Return False
'        End If
'        If values.O_Add_Insert_SQLString(lstSql) = False Then
'          Result_Message = "Get Insert PO SQL Failed"
'          Return False
'        End If
'      Next
'      For Each values In dicAdd_PO_DTL.Values
'        If values.O_Add_Insert_SQLString(lstSql) = False Then
'          Result_Message = "Get Insert PO_DTL SQL Failed"
'          Return False
'        End If
'      Next
'      For Each values In dicAdd_PO.Values
'        If values.O_Add_Insert_SQLString(lstSql) = False Then
'          Result_Message = "Get Insert PO SQL Failed"
'          Return False
'        End If
'      Next

'      Return True
'    Catch ex As Exception
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  '執行刪除和新增的SQL語句，並進行記憶體資料更新
'  Private Function _Execute_DataUpdate(ByRef Result_Message As String,
'                                      ByRef dicAdd_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                                      ByRef dicAdd_PO As Dictionary(Of String, eCA_WMSObject.clsPO),
'                                      ByRef dicReAdd_PO_DTL As Dictionary(Of String, eCA_WMSObject.clsPO_DTL),
'                                      ByRef dicReAdd_PO As Dictionary(Of String, eCA_WMSObject.clsPO),
'                                      ByRef lstSql As List(Of String)) As Boolean
'    Try
'      If dicReAdd_PO_DTL.Count <> 0 Or dicReAdd_PO.Count <> 0 Then
'        '要刪除、更新前通知WMS WMS同意才能刪
'        '通知WMS (暫時無接口)

'      End If



'      '更新所有的SQL
'      If eCA_WMSObject.Common_DBManagement.BatchUpdate_DynamicConnection(lstSql) = False Then
'        '更新DB失敗則回傳False
'        Result_Message = "WMS Update DB Failed"
'        Return False
'      End If
'      '修改記憶體資料

'      '刪除後新增objPO_DTL
'      For Each objNewPO_DTL In dicReAdd_PO_DTL.Values
'        objNewPO_DTL.Remove_Relationship()
'        objNewPO_DTL.Add_Relationship(gMain.objWMS)
'      Next
'      '刪除後新增objPO
'      For Each objNewPO In dicReAdd_PO.Values
'        objNewPO.Remove_Relationship()
'        objNewPO.Add_Relationship(gMain.objWMS)
'      Next

'      '新增objPO_DTL
'      For Each objPO_DTL In dicAdd_PO_DTL.Values
'        '建立關聯
'        objPO_DTL.Add_Relationship(gMain.objWMS)
'      Next
'      '新增objPO
'      For Each objPO In dicAdd_PO.Values
'        '建立關聯
'        objPO.Add_Relationship(gMain.objWMS)
'      Next



'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function


'End Module
