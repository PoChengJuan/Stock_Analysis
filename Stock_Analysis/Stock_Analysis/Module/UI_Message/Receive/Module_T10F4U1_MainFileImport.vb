''20190226
''V1.0.0
''Jerry

''主檔匯入

'Imports eCA_HostObject
'Imports eCA_TransactionMessage
'Imports Microsoft.Office.Interop
'Imports NPOI.HSSF.UserModel
'Imports NPOI.XSSF.UserModel
'Imports NPOI.SS.UserModel
'Imports System.IO

'Module Module_T10F4U1_MainFileImport
'  Public Function O_T10F4U1_MainFileImport(ByVal Receive_Msg As MSG_T10F4U1_MainFileImport,
'                                          ByRef ret_strResultMsg As String,
'                                       ByRef ret_Wait_UUID As String) As Boolean
'    Try
'      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)

'      Dim dicAddPO As New Dictionary(Of String, clsPO)
'      Dim dicAddPO_DTL As New Dictionary(Of String, clsPO_DTL)
'      Dim dicDelete_PO As New Dictionary(Of String, clsPO)
'      Dim dicAdd_PO_Line As New Dictionary(Of String, clsPO_LINE)
'      Dim dicDelete_PO_Line As New Dictionary(Of String, clsPO_LINE)
'      Dim dicDelete_PO_DTL As New Dictionary(Of String, clsPO_DTL)

'      '预览资讯
'      Dim strPreview = ""

'      ''儲存要更新的SQL，進行一次性更新
'      Dim lstSql As New List(Of String)

'      '是否提早return (用于预览资料)
'      Dim bln_Return As Boolean = False

'      '先進行資料邏輯檢查
'      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
'        Return False
'      End If
'      '進行資料處理
'      If Process_Data(Receive_Msg, ret_strResultMsg, ret_Wait_UUID, bln_Return, strPreview, Host_Command, dicAddPO, dicAddPO_DTL, dicAdd_PO_Line, dicDelete_PO, dicDelete_PO_DTL, dicDelete_PO_Line) = False Then
'        Return False
'      End If
'      If bln_Return Then
'        If ret_strResultMsg = "" Then
'          If strPreview.Length <= 2000 Then
'            ret_strResultMsg = strPreview
'          Else
'            ret_strResultMsg = strPreview.Substring(0, 2000)
'          End If
'        End If
'        Return True
'      End If

'      '取得SQL
'      If _Get_SQL(ret_strResultMsg, Host_Command, dicAddPO, dicAddPO_DTL, dicAdd_PO_Line, dicDelete_PO, dicDelete_PO_DTL, dicDelete_PO_Line, lstSql) = False Then
'        Return False
'      End If

'      '執行SQL與更新物件
'      If _Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
'        Return False
'      End If



'      Return True
'    Catch ex As Exception
'      ret_strResultMsg = ex.InnerException.Message
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  '檢查相關資料是否正確
'  Private Function Check_Data(ByVal Receive_Msg As MSG_T10F4U1_MainFileImport,
'                              ByRef ret_strResultMsg As String) As Boolean
'    Try
'      '先進行資料邏輯檢查
'      Dim Excute As String = Receive_Msg.Body.Excute
'      If Excute = "" Then
'        ret_strResultMsg = "Excute is null"
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
'        Return False
'      End If

'      For Each objFileInfo In Receive_Msg.Body.FileList.FileInfo
'        '資料檢查
'        Dim MainFileType As String = objFileInfo.MainFileType
'        Dim FileType As String = objFileInfo.FileType
'        Dim FilePath As String = objFileInfo.FilePath
'        If MainFileType = "" Then
'          ret_strResultMsg = "MainFileType is null"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
'          Return False
'        ElseIf CheckValueInEnum(Of enuMainFileType)(MainFileType) = False Then
'          ret_strResultMsg = "MainFileType = " & MainFileType & "  is not defined"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
'          Return False
'        End If
'        If FileType = "" Then
'          ret_strResultMsg = "FileType is null"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
'          Return False
'        ElseIf CheckValueInEnum(Of enuFileType)(FileType) = False Then
'          ret_strResultMsg = "FileType = " & FileType & " is not defined"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
'          Return False
'        End If
'        If FilePath = "" Then
'          ret_strResultMsg = "FilePath is null"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
'          Return False
'        End If
'      Next
'      Return True
'    Catch ex As Exception
'      ret_strResultMsg = ex.InnerException.Message
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'  '資料處理
'  Private Function Process_Data(ByVal Receive_Msg As MSG_T10F4U1_MainFileImport,
'                                ByRef ret_strResultMsg As String, ByRef ret_Wait_UUID As String,
'                                ByRef bln_Return As Boolean, ByRef ret_strPreview As String,
'                                ByRef Host_Command As Dictionary(Of String, clsFromHostCommand),
'                                ByRef ret_dicAddPO As Dictionary(Of String, clsPO),
'                                ByRef ret_dicAddPO_DTL As Dictionary(Of String, clsPO_DTL),
'                                ByRef ret_dicAdd_PO_Line As Dictionary(Of String, clsPO_LINE),
'                                ByRef ret_dicDeletePO As Dictionary(Of String, clsPO),
'                                ByRef ret_dicDeletePO_DTL As Dictionary(Of String, clsPO_DTL),
'                                ByRef ret_dicDelete_PO_Line As Dictionary(Of String, clsPO_LINE)) As Boolean
'    Try
'      '先進行資料邏輯檢查
'      Dim dicAddSKU_NO As New Dictionary(Of String, clsSKU)
'      Dim dicUpdateSKU_NO As New Dictionary(Of String, clsSKU)
'      Dim dicDeleteSKU_NO As New Dictionary(Of String, clsSKU)
'      '收料編輯資訊
'      Dim MSG_Receipt As New MSG_T5F2U5_BatchCreateReceiptByPO
'      Dim MSG_Receipt_Stocktaking As New MSG_T10F2U1_StocktakingManagement

'      '0预览 1执行
'      Dim Excute As String = Receive_Msg.Body.Excute

'      For Each objFileInfo In Receive_Msg.Body.FileList.FileInfo
'        '資料檢查
'        Dim MainFileType As String = objFileInfo.MainFileType
'        Dim FileType As String = objFileInfo.FileType
'        Dim FilePath As String = objFileInfo.FilePath

'        Select Case MainFileType
'          Case enuMainFileType.SKU
'            Select Case FileType
'              Case enuFileType.EXCEL
'                SendMessageToLog("解汇入的Excel_料品", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
'                'Dim app As New Excel.Application 'app 是操作 Excel 的變數
'                'Dim workbook As Excel.Workbook 'Workbook 代表的是一個 Excel 本體
'                'workbook = app.Workbooks.Open(FilePath) '開啟一張已存在的 Excel 檔案
'                Dim file As New FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite)
'                Dim workbook As IWorkbook
'                If FilePath.Contains("xlsx") = True Then
'                  workbook = New XSSFWorkbook(file)
'                Else
'                  workbook = New HSSFWorkbook(file)
'                End If
'                If I_SKUExeclImport(ret_strResultMsg, FilePath, dicAddSKU_NO, dicUpdateSKU_NO, dicDeleteSKU_NO, ret_strPreview, workbook) = False Then
'                  'workbook.Close()
'                  'app.Quit() '結束操作
'                  file.Close()
'                  Return False
'                End If
'                'workbook.Close()
'                'app.Quit() '結束操作
'                file.Close()
'            End Select
'            '检查是否可以执行
'            SendMessageToLog("检查新增、修改、删除的料号是否可执行", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
'            If I_Check_Add_Update_DeleteSKU(ret_strResultMsg, dicAddSKU_NO, dicUpdateSKU_NO, dicDeleteSKU_NO) = False Then
'              Return False
'            End If
'          Case enuMainFileType.PO_IN_OUT
'            Select Case FileType
'              Case enuFileType.EXCEL
'                SendMessageToLog("解汇入的Excel_入出庫單", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
'                Dim app As New Excel.Application 'app 是操作 Excel 的變數
'                Dim workbook As Excel.Workbook 'Workbook 代表的是一個 Excel 本體
'                workbook = app.Workbooks.Open(FilePath) '開啟一張已存在的 Excel 檔案
'                If I_POExeclImport(Receive_Msg, ret_strResultMsg, FilePath, ret_dicAddPO, ret_dicAddPO_DTL, ret_dicAdd_PO_Line, ret_strPreview, workbook) = False Then
'                  workbook.Close()
'                  app.Quit() '結束操作
'                  Return False
'                End If
'                workbook.Close()
'                app.Quit() '結束操作
'            End Select
'            '检查是否可以执行
'            SendMessageToLog("检查訂單料號是否存在", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
'            If I_Check_Add_PO(ret_strResultMsg, ret_dicAddPO, ret_dicAddPO_DTL, ret_dicAdd_PO_Line, ret_dicDeletePO, ret_dicDeletePO_DTL, ret_dicDelete_PO_Line) = False Then
'              Return False
'            End If
'          Case enuMainFileType.ReceiptInfo
'            Select Case FileType
'              Case enuFileType.EXCEL
'                SendMessageToLog("解汇入的Excel_收料編輯", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
'                Dim app As New Excel.Application 'app 是操作 Excel 的變數
'                Dim workbook As Excel.Workbook 'Workbook 代表的是一個 Excel 本體
'                workbook = app.Workbooks.Open(FilePath) '開啟一張已存在的 Excel 檔案
'                If I_ReceiptExeclImport(Receive_Msg, ret_strResultMsg, FilePath, MSG_Receipt, ret_strPreview, workbook) = False Then
'                  workbook.Close()
'                  app.Quit() '結束操作
'                  Return False
'                End If
'                workbook.Close()
'                app.Quit() '結束操作
'            End Select
'            ''检查是否可以执行
'            'SendMessageToLog("检查訂單料號是否存在", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
'            'If I_Check_Add_PO(ret_strResultMsg, ret_dicAddPO, ret_dicAddPO_DTL, ret_dicAdd_PO_Line, ret_dicDeletePO, ret_dicDeletePO_DTL, ret_dicDelete_PO_Line) = False Then
'            '  Return False
'            'End If
'          Case enuMainFileType.Stocktaking
'            Select Case FileType
'              Case enuFileType.EXCEL
'                SendMessageToLog("解汇入的Excel_盤點單", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
'                Dim app As New Excel.Application 'app 是操作 Excel 的變數
'                Dim workbook As Excel.Workbook 'Workbook 代表的是一個 Excel 本體
'                workbook = app.Workbooks.Open(FilePath) '開啟一張已存在的 Excel 檔案
'                If I_StocktakingExeclImport(Receive_Msg, ret_strResultMsg, FilePath, MSG_Receipt_Stocktaking, ret_strPreview, workbook) = False Then
'                  workbook.Close()
'                  app.Quit() '結束操作
'                  Return False
'                End If
'                workbook.Close()
'                app.Quit() '結束操作
'            End Select
'        End Select
'      Next

'      If Excute = "1" Then
'        '发送事件给WMS
'        If dicAddSKU_NO.Any Then
'          If Module_Send_WMSMessage.Send_T2F3U1_SKUManagement_to_WMS(ret_strResultMsg, dicAddSKU_NO, Host_Command, "Create") = False Then
'            Return False
'          End If
'        End If
'        If dicUpdateSKU_NO.Any Then
'          If Module_Send_WMSMessage.Send_T2F3U1_SKUManagement_to_WMS(ret_strResultMsg, dicUpdateSKU_NO, Host_Command, "Modify") = False Then
'            Return False
'          End If
'        End If
'        If dicDeleteSKU_NO.Any Then
'          If Module_Send_WMSMessage.Send_T2F3U1_SKUManagement_to_WMS(ret_strResultMsg, dicDeleteSKU_NO, Host_Command, "Delete") = False Then
'            Return False
'          End If
'        End If
'        If MSG_Receipt.Body.ReceiptCarrierList.ReceiptCarrierInfo.Any Then
'          '將物件轉成xml
'          Dim strXML = ""
'          If PrepareMessage_MSG(Of MSG_T5F2U5_BatchCreateReceiptByPO)(strXML, MSG_Receipt, ret_strResultMsg) = False Then
'            If ret_strResultMsg = "" Then
'              ret_strResultMsg = "轉XML錯誤(MSG_T5F2U5_BatchCreateReceiptByPO)"
'            End If
'            Return False
'          End If
'          O_Send_MessageToWMS(strXML, MSG_Receipt.Header, Host_Command)
'        End If
'        If MSG_Receipt_Stocktaking.Body IsNot Nothing AndAlso MSG_Receipt_Stocktaking.Body.StocktakingInfo.StocktakingDTLList.StocktakingDTLInfo.Any Then
'          '將物件轉成xml
'          Dim strXML = ""
'          If PrepareMessage_MSG(Of MSG_T10F2U1_StocktakingManagement)(strXML, MSG_Receipt_Stocktaking, ret_strResultMsg) = False Then
'            If ret_strResultMsg = "" Then
'              ret_strResultMsg = "轉XML錯誤(MSG_T10F2U1_StocktakingManagement)"
'            End If
'            Return False
'          End If
'          O_Send_MessageToWMS(strXML, MSG_Receipt_Stocktaking.Header, Host_Command)
'        End If
'      Else
'          bln_Return = True
'      End If


'      Return True
'    Catch ex As Exception
'      ret_strResultMsg = ex.InnerException.Message
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'  '

'  '單據新增檢查
'  Private Function I_Check_Add_PO(ByRef ret_strResultMsg As String,
'                                  ByRef dicAddPO As Dictionary(Of String, clsPO),
'                                  ByRef dicAddPO_DTL As Dictionary(Of String, clsPO_DTL),
'                                  ByRef dicAdd_PO_Line As Dictionary(Of String, clsPO_LINE),
'                                  ByRef dicDeletePO As Dictionary(Of String, clsPO),
'                                  ByRef dicDeletePO_DTL As Dictionary(Of String, clsPO_DTL),
'                                  ByRef dicDelete_PO_Line As Dictionary(Of String, clsPO_LINE)) As Boolean
'    Try
'      '把料品主挡都捞出来
'      Dim dicSKU As New Dictionary(Of String, clsSKU)
'      'If gMain.objHandling.O_GetDB_dicSKUByAll(gdicSKU) = False Then
'      '  ret_strResultMsg = "无法取得料品主擋"
'      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'      '  Return False
'      'End If

'      '檢查項次料品是否存在
'      Dim dicSKU_NO As New Dictionary(Of String, String)
'      For Each objPO_DTL In dicAddPO_DTL.Values
'        If dicSKU_NO.ContainsKey(objPO_DTL.SKU_NO) = False Then
'          dicSKU_NO.Add(objPO_DTL.SKU_NO, objPO_DTL.SKU_NO)
'        End If
'      Next
'      If gMain.objHandling.O_GetDB_lstSKUBydicSKUNo(dicSKU_NO, dicSKU) = False Then
'        ret_strResultMsg = "无法取得料品主擋"
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If


'      Dim gdicPO_DTL As New Dictionary(Of String, clsPO_DTL)
'      If gMain.objHandling.O_GetDB_dicPODTLByALL(gdicPO_DTL) = False Then
'        ret_strResultMsg = "无法取得單據細項"
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If

'      Dim gdicPO_LINE As New Dictionary(Of String, clsPO_LINE)
'      If gMain.objHandling.O_GetDB_dicPOLineByAll(gdicPO_LINE) = False Then
'        ret_strResultMsg = "无法取得單據細項."
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If

'      '撈出所有的單
'      Dim gdicPO As New Dictionary(Of String, clsPO)
'      If gMain.objHandling.O_GetDB_dicPOByALL(gdicPO) = False Then
'        ret_strResultMsg = "无法取得訂單資訊"
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If

'      '檢查單據是否已執行
'      For Each objPO In dicAddPO.Values
'        Dim tmpPO As clsPO = Nothing
'        If gdicPO.TryGetValue(objPO.gid, tmpPO) Then
'          If tmpPO.PO_Status <> enuPOStatus.Queued Then
'            ret_strResultMsg = "單號：" & objPO.PO_ID & " 已執行無法建單。"
'            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'            Return False
'          End If
'          '檢查是不是已有單據 已有則刪除
'          If tmpPO.H_PO_ORDER_TYPE <> enuOrderType.m_general_in AndAlso tmpPO.H_PO_ORDER_TYPE <> enuOrderType.m_grneral_out Then
'            ret_strResultMsg = "單號：" & objPO.PO_ID & " 不屬於WMS內部單據無法更新。"
'            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'            Return False
'          End If


'          If dicDeletePO.ContainsKey(tmpPO.gid) = False Then
'            dicDeletePO.Add(tmpPO.gid, tmpPO)
'          End If
'          Dim dictmpDeletePO_DTL = gdicPO_DTL.Where(Function(_obj)
'                                                      If _obj.Value.PO_ID = objPO.PO_ID Then Return True
'                                                      Return False
'                                                    End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
'          For Each objPODTL In dictmpDeletePO_DTL.Values
'            If dicDeletePO_DTL.ContainsKey(objPODTL.gid) = False Then
'              dicDeletePO_DTL.Add(objPODTL.gid, objPODTL)
'            End If
'          Next
'          Dim dictmpDeletePO_LINE = gdicPO_LINE.Where(Function(_obj)
'                                                        If _obj.Value.PO_ID = objPO.PO_ID Then Return True
'                                                        Return False
'                                                      End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
'          For Each objPOLINE In dictmpDeletePO_LINE.Values
'            If dicDelete_PO_Line.ContainsKey(objPOLINE.gid) = False Then
'              dicDelete_PO_Line.Add(objPOLINE.gid, objPOLINE)
'            End If
'          Next
'        End If
'      Next

'      '檢查項次料品是否存在
'      For Each objPO_DTL In dicAddPO_DTL.Values
'        Dim SKU_KEY = clsSKU.Get_Combination_Key(objPO_DTL.SKU_NO)
'        If dicSKU.ContainsKey(SKU_KEY) = False Then
'          ret_strResultMsg = "單號：" & objPO_DTL.PO_ID & " 項次：" & objPO_DTL.PO_SERIAL_NO & " 料号：" & objPO_DTL.SKU_NO & " 不存在。"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next


'      Return True
'    Catch ex As Exception
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'  '料號更新檢查
'  Private Function I_Check_Add_Update_DeleteSKU(ByRef ret_strResultMsg As String,
'                                                ByRef dicAddSKU_NO As Dictionary(Of String, clsSKU),
'                                                ByRef dicUpdateSKU_NO As Dictionary(Of String, clsSKU),
'                                                ByRef dicDeleteSKU_NO As Dictionary(Of String, clsSKU)) As Boolean
'    Try
'      '把料品主挡都捞出来
'      Dim gdicSKU As New Dictionary(Of String, clsSKU)
'      If gMain.objHandling.O_GetDB_dicSKUByAll(gdicSKU) = False Then
'        ret_strResultMsg = "无法取得料品主擋"
'        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If

'      '如果有要删除的SKU先捞一次库存
'      Dim gdicCarrierItem As New Dictionary(Of String, clsCarrierItem)
'      Dim tmp_dicUsedSKU As New Dictionary(Of String, String) '(料号,料号)
'      If dicDeleteSKU_NO.Any Then
'        If gMain.objHandling.GetCarrierItemByAll(gdicCarrierItem) = False Then
'          ret_strResultMsg = "无法取得库存资料"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        '把有库存的料号加进dic
'        For Each objCarrierItem In gdicCarrierItem.Values
'          If tmp_dicUsedSKU.ContainsKey(objCarrierItem.SKU_No) = False Then
'            tmp_dicUsedSKU.Add(objCarrierItem.SKU_No, objCarrierItem.SKU_No)
'          End If
'        Next
'        '把有单据的料号加进dic
'        Dim lstPO_DTL As New Dictionary(Of String, clsPO_DTL)
'        If gMain.objHandling.O_GetDB_dicPODTLByALL(lstPO_DTL) Then
'          For Each PO_DTL In lstPO_DTL.Values
'            If tmp_dicUsedSKU.ContainsKey(PO_DTL.SKU_NO) = False Then
'              tmp_dicUsedSKU.Add(PO_DTL.SKU_NO, PO_DTL.SKU_NO)
'            End If
'          Next
'        End If
'      End If

'      '检查新增资料
'      For Each objAddSKU As clsSKU In dicAddSKU_NO.Values
'        If gdicSKU.ContainsKey(objAddSKU.gid) Then
'          ret_strResultMsg = "料号：" & objAddSKU.SKU_NO & " 已存在无法新增"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next

'      '检查更新资料
'      For Each objUpdateSKU As clsSKU In dicUpdateSKU_NO.Values
'        Dim objSKU As clsSKU = Nothing
'        If gdicSKU.TryGetValue(objUpdateSKU.gid, objSKU) = False Then
'          ret_strResultMsg = "料号：" & objUpdateSKU.SKU_NO & " 不存在无法更新"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        '更新Excel中没有的
'        objUpdateSKU.CREATE_TIME = objSKU.CREATE_TIME
'        objUpdateSKU.WEIGHT_DIFFERENCE = objSKU.WEIGHT_DIFFERENCE
'        objUpdateSKU.HIGH_WATER = objSKU.HIGH_WATER
'        objUpdateSKU.LOW_WATER = objSKU.LOW_WATER
'        objUpdateSKU.AVAILABLE_DAYS = objSKU.AVAILABLE_DAYS
'        objUpdateSKU.SAVE_DAYS = objSKU.SAVE_DAYS
'      Next

'      '检查删除资料 (不能有库存，不能有单)
'      For Each objDeleteSKU As clsSKU In dicDeleteSKU_NO.Values
'        Dim objSKU As clsSKU = Nothing
'        If gdicSKU.TryGetValue(objDeleteSKU.gid, objSKU) = False Then
'          ret_strResultMsg = "料号：" & objDeleteSKU.SKU_NO & " 不存在无法删除"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        '如果有料号则检查是否有库存货单据
'        If tmp_dicUsedSKU.ContainsKey(objDeleteSKU.SKU_NO) Then
'          ret_strResultMsg = "料号：" & objDeleteSKU.SKU_NO & " 存在对应库存(或单据)"
'          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next

'      Return True
'    Catch ex As Exception
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'  'SQL
'  Private Function _Get_SQL(ByRef Result_Message As String,
'                            ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand),
'                            ByRef ret_dicAddPO As Dictionary(Of String, clsPO),
'                            ByRef ret_dicAddPO_DTL As Dictionary(Of String, clsPO_DTL),
'                            ByRef ret_dicAdd_PO_Line As Dictionary(Of String, clsPO_LINE),
'                            ByRef ret_dicDeletePO As Dictionary(Of String, clsPO),
'                            ByRef ret_dicDeletePO_DTL As Dictionary(Of String, clsPO_DTL),
'                            ByRef ret_dicDelete_PO_Line As Dictionary(Of String, clsPO_LINE),
'                            ByRef lstSql As List(Of String)) As Boolean
'    Try
'      ''取得要更新的SQL
'      For Each _Host_COMMAND In dicHost_Command.Values
'        If _Host_COMMAND.O_Add_Insert_SQLString(lstSql) = False Then
'          Result_Message = "Get Insert HOST_T_WMS_Command SQL Failed"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next
'      For Each objPO In ret_dicDeletePO.Values
'        If objPO.O_Add_Delete_SQLString(lstSql) = False Then
'          Result_Message = "Get Delete PO SQL Failed"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next
'      For Each objPODTL In ret_dicDeletePO_DTL.Values
'        If objPODTL.O_Add_Delete_SQLString(lstSql) = False Then
'          Result_Message = "Get Delete PODTL SQL Failed"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next
'      For Each objPOLINE In ret_dicDelete_PO_Line.Values
'        If objPOLINE.O_Add_Delete_SQLString(lstSql) = False Then
'          Result_Message = "Get Delete POLINE SQL Failed"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next

'      For Each objPO In ret_dicAddPO.Values
'        If objPO.O_Add_Insert_SQLString(lstSql) = False Then
'          Result_Message = "Get Insert PO SQL Failed"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next
'      For Each objPODTL In ret_dicAddPO_DTL.Values
'        If objPODTL.O_Add_Insert_SQLString(lstSql) = False Then
'          Result_Message = "Get Insert PODTL SQL Failed"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'      Next
'      For Each objPOLINE In ret_dicAdd_PO_Line.Values
'        If objPOLINE.O_Add_Insert_SQLString(lstSql) = False Then
'          Result_Message = "Get Insert POLINE SQL Failed"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
'                                      ByRef lstSql As List(Of String)) As Boolean
'    Try
'      '更新所有的SQL
'      If Common_DBManagement.BatchUpdate(lstSql) = False Then
'        '更新DB失敗則回傳False
'        Result_Message = "eHOST 更新资料库失败"
'        Return False
'      End If
'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  '處理 SKU Excel 解析等
'  Private Function I_SKUExeclImport(ByRef Result_Message As String,
'                                    ByRef FilePath As String,
'                                    ByRef dicAddSKU_NO As Dictionary(Of String, clsSKU),
'                                    ByRef dicUpdateSKU_NO As Dictionary(Of String, clsSKU),
'                                    ByRef dicDeleteSKU_NO As Dictionary(Of String, clsSKU),
'                                    ByRef ret_strPreview As String, ByVal workbook As IWorkbook) As Boolean
'    Try
'      Dim now_time = GetNewTime_DBFormat()

'      'Dim worksheet As Excel.Worksheet 'Worksheet 代表的是 Excel 工作表
'      'worksheet = workbook.Worksheets("Data") '讀取其中一張工作表
'      Dim worksheet = workbook.GetSheetAt(0)    '取得第一個Sheet

'      '從第二行 到最後一行
'      For row = 2 To worksheet.LastRowNum - 1
'        '判斷是否為最後一行
'        Dim IRow = worksheet.GetRow(row - 1)
'        If IRow.Cells(0).ToString = "" Then
'          Exit For
'        End If

'        If IsNothing(GetCellData(worksheet, row, 2)) Then Exit For
'        ' worksheet.Cells(1, row).Value() '讀取某一個欄位的值，第一個數字是行，第二個數字是列，如果欄位沒有東西會回傳Nothing
'        Dim SKU_NO As String = IIf(IsNothing(GetCellData(worksheet, row, 2)), "", GetCellData(worksheet, row, 2))
'        If SKU_NO.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        ''檢查資料庫是否有料，有則不理
'        'Dim dicSKU As New Dictionary(Of String, clsSKU)
'        'If gMain.objHandling.O_Get_SKU_NO_By_SKU_NO(SKU_NO, dicSKU) Then
'        '  SendMessageToLog(SKU_NO & " 已存在", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
'        '  If dicSKU.Any = True Then Continue For
'        'End If

'        Dim SKU_ID1 As String = IIf(IsNothing(GetCellData(worksheet, row, 2)), "", GetCellData(worksheet, row, 2)) 'IIf(IsNothing(worksheet.Cells(row, 1).Value()), "", worksheet.Cells(row, 1).Value())
'        If SKU_ID1.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_ID1 = " & SKU_ID1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_ID2 As String = IIf(IsNothing(GetCellData(worksheet, row, 2)), "", GetCellData(worksheet, row, 2)) '_SKU_TYPE2
'        Dim SKU_ID3 As String = ""
'        Dim SKU_ALIS1 As String = IIf(IsNothing(GetCellData(worksheet, row, 3)), "", GetCellData(worksheet, row, 3))
'        Dim SKU_ALIS2 As String = ""
'        Dim SKU_DESC As String = "" 'IIf(IsNothing(worksheet.Cells(row, 3).Value()), "", worksheet.Cells(row, 3).Value())
'        If SKU_ID1.Length > 400 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_DESC = " & SKU_DESC & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_CATALOG As enuSKU_CATALOG = enuSKU_CATALOG.A
'        Dim SKU_TYPE1 As String = "1" 'IIf(IsNothing(worksheet.Cells(row, 3).Value()), "", IIf(IsNumeric(worksheet.Cells(row, 3).Value()), worksheet.Cells(row, 3).Value(), ""))
'        If SKU_TYPE1.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_TYPE1 = " & SKU_TYPE1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        Else
'          If SKU_TYPE1 <> "1" AndAlso SKU_TYPE1 <> "2" AndAlso SKU_TYPE1 <> "3" Then
'            Result_Message = "料号 = " & SKU_NO & " 大分类 = " & SKU_TYPE1 & " 未定义"
'            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'            Return False
'          End If
'        End If
'        Dim SKU_TYPE2 As String = "1" 'IIf(IsNothing(worksheet.Cells(row, 11).Value()), "", worksheet.Cells(row, 11).Value())
'        If SKU_TYPE2.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_TYPE2 = " & SKU_TYPE2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_TYPE3 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 1).Value()), "", worksheet.Cells(row, 1).Value())
'        Dim SKU_COMMON1 As String = IIf(IsNothing(GetCellData(worksheet, row, 4)), "", GetCellData(worksheet, row, 4)) 'IIf(IsNothing(worksheet.Cells(row, 4).Value()), "", worksheet.Cells(row, 4).Value())
'        If SKU_COMMON1.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_COMMON1 = " & SKU_COMMON1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_COMMON2 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 15).Value()), "", worksheet.Cells(row, 15).Value())
'        If SKU_COMMON2.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_COMMON2 = " & SKU_COMMON2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_COMMON3 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 16).Value()), "", worksheet.Cells(row, 16).Value())
'        If SKU_COMMON1.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_COMMON3 = " & SKU_COMMON3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_COMMON4 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 4).Value()), "", worksheet.Cells(row, 4).Value())
'        If SKU_COMMON4.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_COMMON4 = " & SKU_COMMON4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_COMMON5 As String = ""
'        Dim SKU_COMMON6 As String = ""
'        Dim SKU_COMMON7 As String = ""
'        Dim SKU_COMMON8 As String = ""
'        Dim SKU_COMMON9 As String = ""
'        Dim SKU_COMMON10 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 5).Value()), "", IIf(IsNumeric(worksheet.Cells(row, 5).Value()), worksheet.Cells(row, 5).Value(), ""))

'        Dim SKU_L As Long = 0 'IIf(IsNothing(worksheet.Cells(row, 6).Value()), 0, worksheet.Cells(row, 6).Value())
'        If IsNumeric(SKU_L) = False Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_L = " & SKU_L & " 格式錯誤"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_W As Long = 0 ' IIf(IsNothing(worksheet.Cells(row, 7).Value()), 0, worksheet.Cells(row, 7).Value())
'        If IsNumeric(SKU_W) = False Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_W = " & SKU_W & " 格式錯誤"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_H As Long = 0 'IIf(IsNothing(worksheet.Cells(row, 8).Value()), 0, worksheet.Cells(row, 8).Value())
'        If IsNumeric(SKU_H) = False Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_H = " & SKU_H & " 格式錯誤"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_WEIGHT As Long = 0 'IIf(IsNothing(worksheet.Cells(row, 9).Value()), 0, worksheet.Cells(row, 9).Value())
'        If IsNumeric(SKU_WEIGHT) = False Then
'          Result_Message = "SKU_NO = " & SKU_NO & " SKU_WEIGHT = " & SKU_WEIGHT & " 格式錯誤"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim SKU_VALUE As Long = 0
'        Dim SKU_UNIT As String = "" 'IIf(IsNothing(worksheet.Cells(row, 10).Value()), "", worksheet.Cells(row, 10).Value())
'        Dim INBOUND_UNIT = ""
'        Dim OUTBOUND_UNIT = ""
'        Dim HIGH_WATER As Long = 0
'        Dim LOW_WATER As Long = 0
'        Dim AVAILABLE_DAYS As Long = 0
'        Dim SAVE_DAYS As Long = 0
'        Dim CREATE_TIME As String = now_time
'        Dim UPDATE_TIME As String = now_time
'        Dim WEIGHT_DIFFERENCE As Long = 0
'        Dim ENABLE As Boolean = IIf(IsNothing(GetCellData(worksheet, row, 5)), "", GetCellData(worksheet, row, 5)) 'IIf(IsNothing(worksheet.Cells(row, 5).Value()), False, IIf(worksheet.Cells(row, 5).Value() = 1, True, False))

'        Dim EFFECTIVE_DATE = ""
'        Dim FAILURE_DATE = ""
'        Dim QC_METHOD = ""

'        Dim COMMENTS As String = "由EXCEL匯入 " & IIf(IsNothing(GetCellData(worksheet, row, 6)), "", GetCellData(worksheet, row, 6)) 'IIf(IsNothing(worksheet.Cells(row, 6).Value()), "", worksheet.Cells(row, 6).Value())
'        If COMMENTS.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " COMMENTS = " & COMMENTS & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        '定义动作。1新增 2更新 3删除
'        Dim Action As String = IIf(IsNothing(GetCellData(worksheet, row, 7)), "", GetCellData(worksheet, row, 7)) 'IIf(IsNothing(worksheet.Cells(row, 17).Value()), "", worksheet.Cells(row, 17).Value())
'        If Action <> "1" AndAlso Action <> "2" AndAlso Action <> "3" Then
'          Result_Message = "SKU_NO = " & SKU_NO & " 新增/修改/删除 = " & Action & " 非定义"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        Dim New_objSKU As New clsSKU(SKU_NO, SKU_ID1, SKU_ID2, SKU_ID3, SKU_ALIS1, SKU_ALIS2, SKU_DESC, SKU_CATALOG, SKU_TYPE1, SKU_TYPE2, SKU_TYPE3, SKU_COMMON1,
'                                     SKU_COMMON2, SKU_COMMON3, SKU_COMMON4, SKU_COMMON5, SKU_COMMON6, SKU_COMMON7, SKU_COMMON8, SKU_COMMON9, SKU_COMMON10, SKU_L,
'                                     SKU_W, SKU_H, SKU_WEIGHT, SKU_VALUE, SKU_UNIT, INBOUND_UNIT, OUTBOUND_UNIT, HIGH_WATER, LOW_WATER, AVAILABLE_DAYS, SAVE_DAYS,
'                                     CREATE_TIME, UPDATE_TIME, WEIGHT_DIFFERENCE, ENABLE, EFFECTIVE_DATE, FAILURE_DATE, QC_METHOD, COMMENTS)
'        If Action = "1" Then
'          ret_strPreview += "新增"
'          If dicAddSKU_NO.ContainsKey(New_objSKU.gid) = False Then
'            dicAddSKU_NO.Add(New_objSKU.gid, New_objSKU)
'          End If
'        ElseIf Action = "3" Then
'          ret_strPreview += "修改"
'          If dicUpdateSKU_NO.ContainsKey(New_objSKU.gid) = False Then
'            dicUpdateSKU_NO.Add(New_objSKU.gid, New_objSKU)
'          End If
'        ElseIf Action = "2" Then
'          ret_strPreview += "删除"
'          If dicDeleteSKU_NO.ContainsKey(New_objSKU.gid) = False Then
'            dicDeleteSKU_NO.Add(New_objSKU.gid, New_objSKU)
'          End If
'        End If

'        ret_strPreview += " 料号:" & SKU_NO & " 描述:" & SKU_DESC & " 单位:" & SKU_UNIT & " 重量:" & SKU_WEIGHT
'        If SKU_COMMON10 = "0" Then
'          ret_strPreview += " 地上分检;"
'        ElseIf SKU_COMMON10 = "1" Then
'          ret_strPreview += " 线上分检;"
'        End If
'      Next
'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function

'  '處理 PO Excel 解析等
'  Private Function I_POExeclImport(ByVal Receive_Msg As MSG_T10F4U1_MainFileImport,
'                                    ByRef Result_Message As String,
'                                    ByRef FilePath As String,
'                                    ByRef dicAddPO As Dictionary(Of String, clsPO),
'                                    ByRef dicAddPO_DTL As Dictionary(Of String, clsPO_DTL),
'                                    ByRef dicAdd_PO_Line As Dictionary(Of String, clsPO_LINE),
'                                    ByRef ret_strPreview As String, ByVal workbook As Excel.Workbook) As Boolean
'    Try
'      Dim now_time = GetNewTime_DBFormat()
'      Dim USER = Receive_Msg.Header.ClientInfo.UserID
'      Dim worksheet As Excel.Worksheet 'Worksheet 代表的是 Excel 工作表
'      worksheet = workbook.Worksheets("Data") '讀取其中一張工作表
'      Dim PO_ID As String = ""
'      Dim WO_TYPE As String = ""
'      '第二行是PO資訊
'      For row = 2 To 2
'        If IsNothing(worksheet.Cells(row, 1).Value()) Then Exit For

'        PO_ID = IIf(IsNothing(worksheet.Cells(row, 1).Value()), "", worksheet.Cells(row, 1).Value()) '(單號)	
'        If PO_ID.Length > 50 Then
'          Result_Message = "PO_ID = " & PO_ID & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        WO_TYPE = IIf(IsNothing(worksheet.Cells(row, 2).Value()), "", worksheet.Cells(row, 2).Value()) '(單據類型)(1入庫、2出庫)
'        If CheckValueInEnum(Of enuWOType)(WO_TYPE) = False Then
'          Result_Message = "WO_TYPE = " & WO_TYPE & " 不符合規定"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        Dim PRIORITY As String = IIf(IsNothing(worksheet.Cells(row, 3).Value()), "", worksheet.Cells(row, 3).Value()) '(優先權)(1~100 越大約優先)
'        If IsNumeric(PRIORITY) = False OrElse PRIORITY > 100 OrElse PRIORITY < 0 Then
'          Result_Message = "PRIORITY = " & WO_TYPE & " 不符合規定"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim Auto_Bound As String = IntegerConvertToBoolean(IIf(IsNothing(worksheet.Cells(row, 4).Value()), "", worksheet.Cells(row, 4).Value()))
'        Dim SHIPPING_NO As String = IIf(IsNothing(worksheet.Cells(row, 5).Value()), "", worksheet.Cells(row, 5).Value())
'        Dim WRITE_OFF_NO As String = IIf(IsNothing(worksheet.Cells(row, 6).Value()), "", worksheet.Cells(row, 6).Value())
'        Dim PO_TYPE1 As String = IIf(WO_TYPE = enuWOType.Receipt, enuPOType_1.Combination_in, enuPOType_1.Picking_out) 'IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        Dim PO_TYPE2 As String = IIf(IsNothing(worksheet.Cells(row, 8).Value()), "", worksheet.Cells(row, 8).Value())
'        Dim PO_TYPE3 As String = IIf(IsNothing(worksheet.Cells(row, 9).Value()), "", worksheet.Cells(row, 9).Value())
'        Dim CUSTOMER_NO As String = IIf(IsNothing(worksheet.Cells(row, 10).Value()), "", worksheet.Cells(row, 10).Value())
'        Dim CLASS_NO As String = IIf(IsNothing(worksheet.Cells(row, 11).Value()), "", worksheet.Cells(row, 11).Value())
'        Dim H_PO_ORDER_TYPE As String = IIf(IsNothing(worksheet.Cells(row, 12).Value()), "", worksheet.Cells(row, 12).Value())
'        Dim H_PO1 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 13).Value()), "", worksheet.Cells(row, 13).Value())
'        Dim H_PO2 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 14).Value()), "", worksheet.Cells(row, 14).Value())
'        Dim H_PO3 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 15).Value()), "", worksheet.Cells(row, 15).Value())
'        Dim H_PO4 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 16).Value()), "", worksheet.Cells(row, 16).Value())
'        Dim H_PO5 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 17).Value()), "", worksheet.Cells(row, 17).Value())
'        Dim H_PO6 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 18).Value()), "", worksheet.Cells(row, 18).Value())
'        Dim H_PO7 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 19).Value()), "", worksheet.Cells(row, 19).Value())
'        Dim H_PO8 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 20).Value()), "", worksheet.Cells(row, 20).Value())
'        Dim H_PO9 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 21).Value()), "", worksheet.Cells(row, 21).Value())
'        Dim H_PO10 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 22).Value()), "", worksheet.Cells(row, 22).Value())
'        Dim H_PO11 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 23).Value()), "", worksheet.Cells(row, 23).Value())
'        Dim H_PO12 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 24).Value()), "", worksheet.Cells(row, 24).Value())
'        Dim H_PO13 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 25).Value()), "", worksheet.Cells(row, 25).Value())
'        Dim H_PO14 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 26).Value()), "", worksheet.Cells(row, 26).Value())
'        Dim H_PO15 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 27).Value()), "", worksheet.Cells(row, 27).Value())
'        Dim H_PO16 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 28).Value()), "", worksheet.Cells(row, 28).Value())
'        Dim H_PO17 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 29).Value()), "", worksheet.Cells(row, 29).Value())
'        Dim H_PO18 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 30).Value()), "", worksheet.Cells(row, 30).Value())
'        Dim H_PO19 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 31).Value()), "", worksheet.Cells(row, 31).Value())
'        Dim H_PO20 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 32).Value()), "", worksheet.Cells(row, 32).Value())

'        Dim Start_Time = ""
'        Dim Finish_Time = ""
'        Dim H_PO_FINISH_TIME = ""

'        If WO_TYPE = enuWOType.Receipt Then
'          ret_strPreview += "建立入庫單據，單號：" & PO_ID
'          PO_TYPE2 = enuPOType_2.m_general_in
'          H_PO_ORDER_TYPE = enuOrderType.m_general_in
'        Else
'          ret_strPreview += "建立出庫單據，單號：" & PO_ID
'          PO_TYPE2 = enuPOType_2.m_grneral_out
'          H_PO_ORDER_TYPE = enuOrderType.m_grneral_out
'        End If

'        'Dim objNewPO As New clsPO(PO_ID, PO_TYPE1, PO_TYPE2, PO_TYPE3, WO_TYPE, PRIORITY, now_time, Start_Time, Finish_Time, USER, CUSTOMER_NO, CLASS_NO,
'        '                          SHIPPING_NO, enuPOStatus.Queued, WRITE_OFF_NO, Auto_Bound, now_time, H_PO_FINISH_TIME, enuPOStatus.Queued, H_PO_ORDER_TYPE,
'        '                          H_PO1, H_PO2, H_PO3, H_PO4, H_PO5, H_PO6, H_PO7, H_PO8, H_PO9, H_PO10, H_PO11, H_PO12, H_PO13, H_PO14, H_PO15, H_PO16,
'        '                          H_PO17, H_PO18, H_PO19, H_PO20,)

'        'If dicAddPO.ContainsKey(objNewPO.gid) = False Then
'        '  dicAddPO.Add(objNewPO.gid, objNewPO)
'        'Else
'        'Result_Message = "存在相同訂單號 = " & PO_ID
'        '  SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        '  Return False
'        'End If

'      Next

'      'PO_DTL
'      For row = 4 To worksheet.Rows.Count - 1
'        If IsNothing(worksheet.Cells(row, 1).Value()) Then Exit For
'        ' worksheet.Cells(1, row).Value() '讀取某一個欄位的值，第一個數字是行，第二個數字是列，如果欄位沒有東西會回傳Nothing

'        Dim PO_SERIAL_NO As String = IIf(IsNothing(worksheet.Cells(row, 1).Value()), "", worksheet.Cells(row, 1).Value())
'        Dim WORKING_TYPE = ""
'        Dim WORKING_SERIAL_NO = ""
'        Dim WORKING_SERIAL_SEQ = ""
'        Dim SKU_NO As String = IIf(IsNothing(worksheet.Cells(row, 2).Value()), "", worksheet.Cells(row, 2).Value())
'        SKU_NO = SKU_NO.Trim()
'        Dim LOT_NO As String = IIf(IsNothing(worksheet.Cells(row, 3).Value()), "", worksheet.Cells(row, 3).Value())
'        LOT_NO = LOT_NO.Trim()
'        Dim QTY As String = IIf(IsNothing(worksheet.Cells(row, 4).Value()), "", worksheet.Cells(row, 4).Value())
'        Dim OWNER As String = IIf(IsNothing(worksheet.Cells(row, 5).Value()), "", worksheet.Cells(row, 5).Value())
'        OWNER = OWNER.Trim()
'        Dim SUBOWNER As String = IIf(IsNothing(worksheet.Cells(row, 6).Value()), "", worksheet.Cells(row, 6).Value())
'        SUBOWNER = SUBOWNER.Trim()
'        Dim PACKAGE_ID As String = "" ' IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        Dim ITEM_COMMON1 As String = "ERP" 'IIf(IsNothing(worksheet.Cells(row, 8).Value()), "", worksheet.Cells(row, 8).Value())
'        Dim ITEM_COMMON2 As String = IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        ITEM_COMMON2 = ITEM_COMMON2.Trim()
'        Dim ITEM_COMMON3 As String = "B6-00-00" 'IIf(IsNothing(worksheet.Cells(row, 10).Value()), "", worksheet.Cells(row, 10).Value())
'        Dim ITEM_COMMON4 As String = IIf(IsNothing(worksheet.Cells(row, 11).Value()), "", worksheet.Cells(row, 11).Value())
'        Dim ITEM_COMMON5 As String = IIf(IsNothing(worksheet.Cells(row, 12).Value()), "", worksheet.Cells(row, 12).Value())
'        Dim ITEM_COMMON6 As String = IIf(IsNothing(worksheet.Cells(row, 13).Value()), "", worksheet.Cells(row, 13).Value())
'        Dim ITEM_COMMON7 As String = IIf(IsNothing(worksheet.Cells(row, 14).Value()), "", worksheet.Cells(row, 14).Value())
'        Dim ITEM_COMMON8 As String = IIf(IsNothing(worksheet.Cells(row, 15).Value()), "", worksheet.Cells(row, 15).Value())
'        Dim ITEM_COMMON9 As String = IIf(IsNothing(worksheet.Cells(row, 16).Value()), "", worksheet.Cells(row, 16).Value())
'        Dim ITEM_COMMON10 As String = IIf(IsNothing(worksheet.Cells(row, 17).Value()), "", worksheet.Cells(row, 17).Value())
'        Dim SORT_ITEM_COMMON1 As String = IIf(IsNothing(worksheet.Cells(row, 18).Value()), "", worksheet.Cells(row, 18).Value())
'        Dim SORT_ITEM_COMMON2 As String = IIf(IsNothing(worksheet.Cells(row, 19).Value()), "", worksheet.Cells(row, 19).Value())
'        Dim SORT_ITEM_COMMON3 As String = IIf(IsNothing(worksheet.Cells(row, 20).Value()), "", worksheet.Cells(row, 20).Value())
'        Dim SORT_ITEM_COMMON4 As String = IIf(IsNothing(worksheet.Cells(row, 21).Value()), "", worksheet.Cells(row, 21).Value())
'        Dim SORT_ITEM_COMMON5 As String = IIf(IsNothing(worksheet.Cells(row, 22).Value()), "", worksheet.Cells(row, 22).Value())
'        Dim STORAGE_TYPE As String = enuStorageType.Store
'        Dim BND = enuBND.NB ' IIf(IsNothing(worksheet.Cells(row, 23).Value()), "", worksheet.Cells(row, 23).Value())
'        Dim QC_STATUS = enuQCStatus.NULL ' IIf(IsNothing(worksheet.Cells(row, 24).Value()), "", worksheet.Cells(row, 24).Value())
'        Dim H_POD1 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 25).Value()), "", worksheet.Cells(row, 25).Value())
'        Dim H_POD2 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 26).Value()), "", worksheet.Cells(row, 26).Value())
'        Dim H_POD3 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 27).Value()), "", worksheet.Cells(row, 27).Value())
'        Dim H_POD4 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 28).Value()), "", worksheet.Cells(row, 28).Value())
'        Dim H_POD5 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 29).Value()), "", worksheet.Cells(row, 29).Value())
'        Dim H_POD6 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 30).Value()), "", worksheet.Cells(row, 30).Value())
'        Dim H_POD7 As String = ITEM_COMMON3 'IIf(IsNothing(worksheet.Cells(row, 31).Value()), "", worksheet.Cells(row, 31).Value())
'        Dim H_POD8 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 32).Value()), "", worksheet.Cells(row, 32).Value())
'        Dim H_POD9 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 33).Value()), "", worksheet.Cells(row, 33).Value())
'        Dim H_POD10 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 34).Value()), "", worksheet.Cells(row, 34).Value())
'        Dim H_POD11 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 35).Value()), "", worksheet.Cells(row, 35).Value())
'        Dim H_POD12 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 36).Value()), "", worksheet.Cells(row, 36).Value())
'        Dim H_POD13 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 37).Value()), "", worksheet.Cells(row, 37).Value())
'        Dim H_POD14 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 38).Value()), "", worksheet.Cells(row, 38).Value())
'        Dim H_POD15 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 39).Value()), "", worksheet.Cells(row, 39).Value())
'        Dim H_POD16 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 40).Value()), "", worksheet.Cells(row, 40).Value())
'        Dim H_POD17 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 41).Value()), "", worksheet.Cells(row, 41).Value())
'        Dim H_POD18 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 42).Value()), "", worksheet.Cells(row, 42).Value())
'        Dim H_POD19 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 43).Value()), "", worksheet.Cells(row, 43).Value())
'        Dim H_POD20 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 44).Value()), "", worksheet.Cells(row, 44).Value())
'        Dim H_POD21 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 45).Value()), "", worksheet.Cells(row, 45).Value())
'        Dim H_POD22 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 46).Value()), "", worksheet.Cells(row, 46).Value())
'        Dim H_POD23 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 47).Value()), "", worksheet.Cells(row, 47).Value())
'        Dim H_POD24 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 48).Value()), "", worksheet.Cells(row, 48).Value())
'        Dim H_POD25 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 49).Value()), "", worksheet.Cells(row, 49).Value())
'        Dim COMMENTS = ""
'        Dim QTY_PROCESS = 0
'        Dim QTY_FINISH = 0
'        Dim PO_LINE_NO = PO_SERIAL_NO

'        Dim FROM_OWNER_ID = ""
'        Dim FROM_SUB_OWNER_ID = ""
'        Dim TO_OWNER_ID = ""
'        Dim TO_SUB_OWNER_ID = ""
'        '根據入出庫類型填入對應資訊
'        If WO_TYPE = enuWOType.Receipt Then
'          TO_OWNER_ID = OWNER
'          TO_SUB_OWNER_ID = SUBOWNER
'        Else
'          FROM_OWNER_ID = OWNER
'          FROM_SUB_OWNER_ID = SUBOWNER
'        End If
'        Dim FACTORY_ID = ""
'        Dim DEST_AREA_ID = ""
'        Dim DEST_LOCATION_ID = ""
'        Dim H_POD_STEP_NO = enuPOStatus.Queued
'        Dim H_POD_MOVE_TYPE = ""
'        Dim H_POD_FINISH_TIME = ""
'        Dim H_POD_BILLING_DATE = ""
'        Dim H_POD_CREATE_TIME = now_time
'        Dim PODTL_STATUS = enuPODTLStatus.Queued


'        If PO_SERIAL_NO.Length > 50 Then
'          Result_Message = "PO_SERIAL_NO = " & PO_SERIAL_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SKU_NO.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If LOT_NO.Length > 50 Then
'          Result_Message = "LOT_NO = " & LOT_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If IntegerCheckPositive(QTY) = False Then
'          Result_Message = "QTY = " & QTY & " 格式錯誤"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If OWNER.Length > 50 Then
'          Result_Message = "OWNER = " & OWNER & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SUBOWNER.Length > 50 Then
'          Result_Message = "SUBOWNER = " & SUBOWNER & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If PACKAGE_ID.Length > 50 Then
'          Result_Message = "PACKAGE_ID = " & PACKAGE_ID & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON1.Length > 50 Then
'          Result_Message = "ITEM_COMMON1 = " & ITEM_COMMON1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON2.Length > 50 Then
'          Result_Message = "ITEM_COMMON2 = " & ITEM_COMMON2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON3.Length > 50 Then
'          Result_Message = "ITEM_COMMON3 = " & ITEM_COMMON3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON4.Length > 50 Then
'          Result_Message = "ITEM_COMMON4 = " & ITEM_COMMON4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON5.Length > 50 Then
'          Result_Message = "ITEM_COMMON5 = " & ITEM_COMMON5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON6.Length > 50 Then
'          Result_Message = "ITEM_COMMON6 = " & ITEM_COMMON6 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON7.Length > 50 Then
'          Result_Message = "ITEM_COMMON7 = " & ITEM_COMMON7 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON8.Length > 50 Then
'          Result_Message = "ITEM_COMMON8 = " & ITEM_COMMON8 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON9.Length > 50 Then
'          Result_Message = "ITEM_COMMON9 = " & ITEM_COMMON9 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON10.Length > 50 Then
'          Result_Message = "ITEM_COMMON10 = " & ITEM_COMMON10 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON1.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON1 = " & SORT_ITEM_COMMON1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON2.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON2 = " & SORT_ITEM_COMMON2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON3.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON3 = " & SORT_ITEM_COMMON3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON4.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON4 = " & SORT_ITEM_COMMON4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON5.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON5 = " & SORT_ITEM_COMMON5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If STORAGE_TYPE.Length > 50 Then
'          Result_Message = "STORAGE_TYPE = " & ITEM_COMMON1 & " STORAGE_TYPE"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If CheckValueInEnum(Of enuBND)(BND) = False Then
'          Result_Message = "BND = " & BND & " is not defined"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If CheckValueInEnum(Of enuQCStatus)(QC_STATUS) = False Then
'          Result_Message = "QC_STATUS = " & QC_STATUS & " is not defined"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD1.Length > 50 Then
'          Result_Message = "H_POD1 = " & H_POD1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD2.Length > 50 Then
'          Result_Message = "H_POD2 = " & H_POD2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD3.Length > 50 Then
'          Result_Message = "H_POD3 = " & H_POD3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD4.Length > 50 Then
'          Result_Message = "H_POD4 = " & H_POD4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD5.Length > 50 Then
'          Result_Message = "H_POD5 = " & H_POD5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD6.Length > 50 Then
'          Result_Message = "H_POD6 = " & H_POD6 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD7.Length > 50 Then
'          Result_Message = "H_POD7 = " & H_POD7 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD8.Length > 50 Then
'          Result_Message = "H_POD8 = " & H_POD8 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD9.Length > 50 Then
'          Result_Message = "H_POD9 = " & H_POD9 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD10.Length > 50 Then
'          Result_Message = "H_POD10 = " & H_POD10 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD11.Length > 50 Then
'          Result_Message = "H_POD11 = " & H_POD11 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD12.Length > 50 Then
'          Result_Message = "H_POD12 = " & H_POD12 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD13.Length > 50 Then
'          Result_Message = "H_POD13 = " & H_POD13 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD14.Length > 50 Then
'          Result_Message = "H_POD14 = " & H_POD14 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD15.Length > 50 Then
'          Result_Message = "H_POD15 = " & H_POD15 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD16.Length > 50 Then
'          Result_Message = "H_POD16 = " & H_POD16 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD17.Length > 50 Then
'          Result_Message = "H_POD17 = " & H_POD17 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD18.Length > 50 Then
'          Result_Message = "H_POD18 = " & H_POD18 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD19.Length > 50 Then
'          Result_Message = "H_POD19 = " & H_POD19 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD20.Length > 50 Then
'          Result_Message = "H_POD20 = " & H_POD20 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD21.Length > 50 Then
'          Result_Message = "H_POD21 = " & H_POD21 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD22.Length > 50 Then
'          Result_Message = "H_POD22 = " & H_POD22 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD23.Length > 50 Then
'          Result_Message = "H_POD23 = " & H_POD23 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD24.Length > 50 Then
'          Result_Message = "H_POD24 = " & H_POD24 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD25.Length > 50 Then
'          Result_Message = "H_POD25 = " & H_POD25 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        Dim CLOSE_ABLE = False
'        Dim objNewPO_DTL As New clsPO_DTL(PO_ID, PO_LINE_NO, PO_SERIAL_NO, WORKING_TYPE, WORKING_SERIAL_NO, WORKING_SERIAL_SEQ, SKU_NO, LOT_NO, QTY, QTY_PROCESS, QTY_FINISH, COMMENTS, PACKAGE_ID, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6,
'                             ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, STORAGE_TYPE, BND, QC_STATUS,
'                             FROM_OWNER_ID, FROM_SUB_OWNER_ID, TO_OWNER_ID, TO_SUB_OWNER_ID, FACTORY_ID, DEST_AREA_ID, DEST_LOCATION_ID, H_POD_STEP_NO, H_POD_MOVE_TYPE, H_POD_FINISH_TIME, H_POD_BILLING_DATE,
'                             H_POD_CREATE_TIME, H_POD1, H_POD2, H_POD3, H_POD4, H_POD5, H_POD6, H_POD7, H_POD8, H_POD9, H_POD10, H_POD11, H_POD12, H_POD13, H_POD14, H_POD15, H_POD16, H_POD17, H_POD18, H_POD19,
'                             H_POD20, H_POD21, H_POD22, H_POD23, H_POD24, H_POD25, PODTL_STATUS, CLOSE_ABLE)
'        If dicAddPO_DTL.ContainsKey(objNewPO_DTL.gid) = False Then
'          dicAddPO_DTL.Add(objNewPO_DTL.gid, objNewPO_DTL)
'        Else
'          Result_Message = "存在相同訂單項次，訂單號 = " & PO_ID & " 項次 = " & PO_SERIAL_NO
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        Dim objNewPO_LINE As New clsPO_LINE(PO_ID, PO_LINE_NO, QTY, QTY_FINISH, 0, "", "", "", "", "")
'        If dicAdd_PO_Line.ContainsKey(objNewPO_LINE.gid) = False Then
'          dicAdd_PO_Line.Add(objNewPO_LINE.gid, objNewPO_LINE)
'        Else
'          Result_Message = "存在相同訂單項次，訂單號 = " & PO_ID & " 項次 = " & PO_SERIAL_NO
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If



'        ret_strPreview += " 項次:" & PO_SERIAL_NO & " 料号:" & SKU_NO & " 數量：" & QTY & " Batch：" & LOT_NO & " Plant：" & OWNER & " StorageLocation：" & SUBOWNER & " CompanyCode：" & ITEM_COMMON2 & " ;"


'      Next




'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function
'  '處理 Receipt Excel 解析等
'  Private Function I_ReceiptExeclImport(ByVal Receive_Msg As MSG_T10F4U1_MainFileImport,
'                                        ByRef Result_Message As String,
'                                        ByRef FilePath As String,
'                                        ByRef MSG_Receipt As MSG_T5F2U5_BatchCreateReceiptByPO,
'                                        ByRef ret_strPreview As String, ByVal workbook As Excel.Workbook) As Boolean
'    Try
'      Dim now_time = GetNewTime_DBFormat()
'      Dim USER = Receive_Msg.Header.ClientInfo.UserID
'      Dim worksheet As Excel.Worksheet 'Worksheet 代表的是 Excel 工作表
'      worksheet = workbook.Worksheets("Data") '讀取其中一張工作表
'      Dim PO_ID As String = ""
'      Dim WO_TYPE As String = ""
'      '第二行是PO資訊
'      For row = 2 To 2
'        If IsNothing(worksheet.Cells(row, 1).Value()) Then Exit For

'        PO_ID = IIf(IsNothing(worksheet.Cells(row, 1).Value()), "", worksheet.Cells(row, 1).Value()) '(單號)	
'        If PO_ID.Length > 50 Then
'          Result_Message = "PO_ID = " & PO_ID & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        WO_TYPE = IIf(IsNothing(worksheet.Cells(row, 2).Value()), "", worksheet.Cells(row, 2).Value()) '(單據類型)(1入庫、2出庫)
'        If CheckValueInEnum(Of enuWOType)(WO_TYPE) = False Then
'          Result_Message = "WO_TYPE = " & WO_TYPE & " 不符合規定"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        ElseIf WO_TYPE <> enuWOType.Receipt Then
'          Result_Message = "WO_TYPE = " & WO_TYPE & " 不為入庫"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        Dim PRIORITY As String = IIf(IsNothing(worksheet.Cells(row, 3).Value()), "", worksheet.Cells(row, 3).Value()) '(優先權)(1~100 越大約優先)
'        If IsNumeric(PRIORITY) = False OrElse PRIORITY > 100 OrElse PRIORITY < 0 Then
'          Result_Message = "PRIORITY = " & WO_TYPE & " 不符合規定"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        Dim Auto_Bound As String = IntegerConvertToBoolean(IIf(IsNothing(worksheet.Cells(row, 4).Value()), "", worksheet.Cells(row, 4).Value()))
'        'Dim SHIPPING_NO As String = IIf(IsNothing(worksheet.Cells(row, 5).Value()), "", worksheet.Cells(row, 5).Value())
'        'Dim WRITE_OFF_NO As String = IIf(IsNothing(worksheet.Cells(row, 6).Value()), "", worksheet.Cells(row, 6).Value())
'        'Dim PO_TYPE1 As String = IIf(WO_TYPE = enuWOType.Receipt, enuPOType_1.Combination_in, enuPOType_1.Picking_out) 'IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        'Dim PO_TYPE2 As String = IIf(IsNothing(worksheet.Cells(row, 8).Value()), "", worksheet.Cells(row, 8).Value())
'        'Dim PO_TYPE3 As String = IIf(IsNothing(worksheet.Cells(row, 9).Value()), "", worksheet.Cells(row, 9).Value())
'        'Dim CUSTOMER_NO As String = IIf(IsNothing(worksheet.Cells(row, 10).Value()), "", worksheet.Cells(row, 10).Value())
'        'Dim CLASS_NO As String = IIf(IsNothing(worksheet.Cells(row, 11).Value()), "", worksheet.Cells(row, 11).Value())
'        'Dim H_PO_ORDER_TYPE As String = IIf(IsNothing(worksheet.Cells(row, 12).Value()), "", worksheet.Cells(row, 12).Value())
'        'Dim H_PO1 As String = IIf(IsNothing(worksheet.Cells(row, 13).Value()), "", worksheet.Cells(row, 13).Value())
'        'Dim H_PO2 As String = IIf(IsNothing(worksheet.Cells(row, 14).Value()), "", worksheet.Cells(row, 14).Value())
'        'Dim H_PO3 As String = IIf(IsNothing(worksheet.Cells(row, 15).Value()), "", worksheet.Cells(row, 15).Value())
'        'Dim H_PO4 As String = IIf(IsNothing(worksheet.Cells(row, 16).Value()), "", worksheet.Cells(row, 16).Value())
'        'Dim H_PO5 As String = IIf(IsNothing(worksheet.Cells(row, 17).Value()), "", worksheet.Cells(row, 17).Value())
'        'Dim H_PO6 As String = IIf(IsNothing(worksheet.Cells(row, 18).Value()), "", worksheet.Cells(row, 18).Value())
'        'Dim H_PO7 As String = IIf(IsNothing(worksheet.Cells(row, 19).Value()), "", worksheet.Cells(row, 19).Value())
'        'Dim H_PO8 As String = IIf(IsNothing(worksheet.Cells(row, 20).Value()), "", worksheet.Cells(row, 20).Value())
'        'Dim H_PO9 As String = IIf(IsNothing(worksheet.Cells(row, 21).Value()), "", worksheet.Cells(row, 21).Value())
'        'Dim H_PO10 As String = IIf(IsNothing(worksheet.Cells(row, 22).Value()), "", worksheet.Cells(row, 22).Value())
'        'Dim H_PO11 As String = IIf(IsNothing(worksheet.Cells(row, 23).Value()), "", worksheet.Cells(row, 23).Value())
'        'Dim H_PO12 As String = IIf(IsNothing(worksheet.Cells(row, 24).Value()), "", worksheet.Cells(row, 24).Value())
'        'Dim H_PO13 As String = IIf(IsNothing(worksheet.Cells(row, 25).Value()), "", worksheet.Cells(row, 25).Value())
'        'Dim H_PO14 As String = IIf(IsNothing(worksheet.Cells(row, 26).Value()), "", worksheet.Cells(row, 26).Value())
'        'Dim H_PO15 As String = IIf(IsNothing(worksheet.Cells(row, 27).Value()), "", worksheet.Cells(row, 27).Value())
'        'Dim H_PO16 As String = IIf(IsNothing(worksheet.Cells(row, 28).Value()), "", worksheet.Cells(row, 28).Value())
'        'Dim H_PO17 As String = IIf(IsNothing(worksheet.Cells(row, 29).Value()), "", worksheet.Cells(row, 29).Value())
'        'Dim H_PO18 As String = IIf(IsNothing(worksheet.Cells(row, 30).Value()), "", worksheet.Cells(row, 30).Value())
'        'Dim H_PO19 As String = IIf(IsNothing(worksheet.Cells(row, 31).Value()), "", worksheet.Cells(row, 31).Value())
'        'Dim H_PO20 As String = IIf(IsNothing(worksheet.Cells(row, 32).Value()), "", worksheet.Cells(row, 32).Value())

'        'Dim Start_Time = ""
'        'Dim Finish_Time = ""
'        'Dim H_PO_FINISH_TIME = ""

'        'If WO_TYPE = enuWOType.Receipt Then
'        '  ret_strPreview += "建立入庫單據，單號：" & PO_ID
'        '  PO_TYPE2 = enuPOType_2.m_general_in
'        '  H_PO_ORDER_TYPE = enuOrderType.m_general_in
'        'Else
'        '  ret_strPreview += "建立出庫單據，單號：" & PO_ID
'        '  PO_TYPE2 = enuPOType_2.m_grneral_out
'        '  H_PO_ORDER_TYPE = enuOrderType.m_grneral_out
'        'End If

'        'Dim objNewPO As New clsPO(PO_ID, PO_TYPE1, PO_TYPE2, PO_TYPE3, WO_TYPE, PRIORITY, now_time, Start_Time, Finish_Time, USER, CUSTOMER_NO, CLASS_NO,
'        '                          SHIPPING_NO, enuPOStatus.Queued, WRITE_OFF_NO, Auto_Bound, now_time, H_PO_FINISH_TIME, enuPOStatus.Queued, H_PO_ORDER_TYPE,
'        '                          H_PO1, H_PO2, H_PO3, H_PO4, H_PO5, H_PO6, H_PO7, H_PO8, H_PO9, H_PO10, H_PO11, H_PO12, H_PO13, H_PO14, H_PO15, H_PO16,
'        '                          H_PO17, H_PO18, H_PO19, H_PO20)

'        'If dicAddPO.ContainsKey(objNewPO.gid) = False Then
'        '  dicAddPO.Add(objNewPO.gid, objNewPO)
'        'Else
'        '  Result_Message = "存在相同訂單號 = " & PO_ID
'        '  SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        '  Return False
'        'End If

'      Next


'      Dim Now_Data = GetNewDate_DBFormat()

'      'PO_DTL
'      For row = 4 To worksheet.Rows.Count - 1
'        If IsNothing(worksheet.Cells(row, 1).Value()) Then Exit For
'        ' worksheet.Cells(1, row).Value() '讀取某一個欄位的值，第一個數字是行，第二個數字是列，如果欄位沒有東西會回傳Nothing

'        Dim PO_SERIAL_NO As String = IIf(IsNothing(worksheet.Cells(row, 1).Value()), "", worksheet.Cells(row, 1).Value())
'        Dim SKU_NO As String = IIf(IsNothing(worksheet.Cells(row, 2).Value()), "", worksheet.Cells(row, 2).Value())
'        SKU_NO = SKU_NO.Trim()
'        Dim LOT_NO As String = IIf(IsNothing(worksheet.Cells(row, 3).Value()), "", worksheet.Cells(row, 3).Value())
'        LOT_NO = LOT_NO.Trim()
'        Dim QTY As String = IIf(IsNothing(worksheet.Cells(row, 4).Value()), "", worksheet.Cells(row, 4).Value())
'        Dim OWNER As String = IIf(IsNothing(worksheet.Cells(row, 5).Value()), "", worksheet.Cells(row, 5).Value())
'        OWNER = OWNER.Trim()
'        Dim SUBOWNER As String = IIf(IsNothing(worksheet.Cells(row, 6).Value()), "", worksheet.Cells(row, 6).Value())
'        SUBOWNER = SUBOWNER.Trim()
'        Dim PACKAGE_ID As String = "" ' IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        Dim ITEM_COMMON1 As String = "ERP" 'IIf(IsNothing(worksheet.Cells(row, 8).Value()), "", worksheet.Cells(row, 8).Value())
'        Dim ITEM_COMMON2 As String = IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        ITEM_COMMON2 = ITEM_COMMON2.Trim()
'        Dim ITEM_COMMON3 As String = "B6-00-00" 'IIf(IsNothing(worksheet.Cells(row, 10).Value()), "", worksheet.Cells(row, 10).Value())
'        Dim Carrier_id As String = IIf(IsNothing(worksheet.Cells(row, 8).Value()), "", worksheet.Cells(row, 8).Value())

'        Dim ITEM_COMMON4 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 11).Value()), "", worksheet.Cells(row, 11).Value())
'        Dim ITEM_COMMON5 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 12).Value()), "", worksheet.Cells(row, 12).Value())
'        Dim ITEM_COMMON6 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 13).Value()), "", worksheet.Cells(row, 13).Value())
'        Dim ITEM_COMMON7 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 14).Value()), "", worksheet.Cells(row, 14).Value())
'        Dim ITEM_COMMON8 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 15).Value()), "", worksheet.Cells(row, 15).Value())
'        Dim ITEM_COMMON9 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 16).Value()), "", worksheet.Cells(row, 16).Value())
'        Dim ITEM_COMMON10 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 17).Value()), "", worksheet.Cells(row, 17).Value())
'        Dim SORT_ITEM_COMMON1 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 18).Value()), "", worksheet.Cells(row, 18).Value())
'        Dim SORT_ITEM_COMMON2 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 19).Value()), "", worksheet.Cells(row, 19).Value())
'        Dim SORT_ITEM_COMMON3 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 20).Value()), "", worksheet.Cells(row, 20).Value())
'        Dim SORT_ITEM_COMMON4 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 21).Value()), "", worksheet.Cells(row, 21).Value())
'        Dim SORT_ITEM_COMMON5 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 22).Value()), "", worksheet.Cells(row, 22).Value())
'        Dim STORAGE_TYPE As String = enuStorageType.Store
'        Dim BND = enuBND.NB ' IIf(IsNothing(worksheet.Cells(row, 23).Value()), "", worksheet.Cells(row, 23).Value())
'        Dim QC_STATUS = enuQCStatus.NULL ' IIf(IsNothing(worksheet.Cells(row, 24).Value()), "", worksheet.Cells(row, 24).Value())
'        Dim H_POD1 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 25).Value()), "", worksheet.Cells(row, 25).Value())
'        Dim H_POD2 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 26).Value()), "", worksheet.Cells(row, 26).Value())
'        Dim H_POD3 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 27).Value()), "", worksheet.Cells(row, 27).Value())
'        Dim H_POD4 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 28).Value()), "", worksheet.Cells(row, 28).Value())
'        Dim H_POD5 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 29).Value()), "", worksheet.Cells(row, 29).Value())
'        Dim H_POD6 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 30).Value()), "", worksheet.Cells(row, 30).Value())
'        Dim H_POD7 As String = ITEM_COMMON3 'IIf(IsNothing(worksheet.Cells(row, 31).Value()), "", worksheet.Cells(row, 31).Value())
'        Dim H_POD8 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 32).Value()), "", worksheet.Cells(row, 32).Value())
'        Dim H_POD9 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 33).Value()), "", worksheet.Cells(row, 33).Value())
'        Dim H_POD10 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 34).Value()), "", worksheet.Cells(row, 34).Value())
'        Dim H_POD11 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 35).Value()), "", worksheet.Cells(row, 35).Value())
'        Dim H_POD12 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 36).Value()), "", worksheet.Cells(row, 36).Value())
'        Dim H_POD13 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 37).Value()), "", worksheet.Cells(row, 37).Value())
'        Dim H_POD14 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 38).Value()), "", worksheet.Cells(row, 38).Value())
'        Dim H_POD15 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 39).Value()), "", worksheet.Cells(row, 39).Value())
'        Dim H_POD16 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 40).Value()), "", worksheet.Cells(row, 40).Value())
'        Dim H_POD17 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 41).Value()), "", worksheet.Cells(row, 41).Value())
'        Dim H_POD18 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 42).Value()), "", worksheet.Cells(row, 42).Value())
'        Dim H_POD19 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 43).Value()), "", worksheet.Cells(row, 43).Value())
'        Dim H_POD20 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 44).Value()), "", worksheet.Cells(row, 44).Value())
'        Dim H_POD21 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 45).Value()), "", worksheet.Cells(row, 45).Value())
'        Dim H_POD22 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 46).Value()), "", worksheet.Cells(row, 46).Value())
'        Dim H_POD23 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 47).Value()), "", worksheet.Cells(row, 47).Value())
'        Dim H_POD24 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 48).Value()), "", worksheet.Cells(row, 48).Value())
'        Dim H_POD25 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 49).Value()), "", worksheet.Cells(row, 49).Value())
'        Dim COMMENTS = ""
'        Dim QTY_PROCESS = 0
'        Dim QTY_FINISH = 0
'        Dim PO_LINE_NO = PO_SERIAL_NO

'        Dim FROM_OWNER_ID = ""
'        Dim FROM_SUB_OWNER_ID = ""
'        Dim TO_OWNER_ID = ""
'        Dim TO_SUB_OWNER_ID = ""
'        '根據入出庫類型填入對應資訊
'        If WO_TYPE = enuWOType.Receipt Then
'          TO_OWNER_ID = OWNER
'          TO_SUB_OWNER_ID = SUBOWNER
'        Else
'          FROM_OWNER_ID = OWNER
'          FROM_SUB_OWNER_ID = SUBOWNER
'        End If
'        Dim FACTORY_ID = ""
'        Dim DEST_AREA_ID = ""
'        Dim DEST_LOCATION_ID = ""
'        Dim H_POD_STEP_NO = enuPOStatus.Queued
'        Dim H_POD_MOVE_TYPE = ""
'        Dim H_POD_FINISH_TIME = ""
'        Dim H_POD_BILLING_DATE = ""
'        Dim H_POD_CREATE_TIME = now_time
'        Dim PODTL_STATUS = enuPODTLStatus.Queued

'        If Carrier_id.Length > 50 Then
'          Result_Message = "膠籃編號 = " & Carrier_id & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If PO_SERIAL_NO.Length > 50 Then
'          Result_Message = "PO_SERIAL_NO = " & PO_SERIAL_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SKU_NO.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If LOT_NO.Length > 50 Then
'          Result_Message = "LOT_NO = " & LOT_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If IntegerCheckPositive(QTY) = False Then
'          Result_Message = "QTY = " & QTY & " 格式錯誤"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If OWNER.Length > 50 Then
'          Result_Message = "OWNER = " & OWNER & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SUBOWNER.Length > 50 Then
'          Result_Message = "SUBOWNER = " & SUBOWNER & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If PACKAGE_ID.Length > 50 Then
'          Result_Message = "PACKAGE_ID = " & PACKAGE_ID & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON1.Length > 50 Then
'          Result_Message = "ITEM_COMMON1 = " & ITEM_COMMON1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON2.Length > 50 Then
'          Result_Message = "ITEM_COMMON2 = " & ITEM_COMMON2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON3.Length > 50 Then
'          Result_Message = "ITEM_COMMON3 = " & ITEM_COMMON3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON4.Length > 50 Then
'          Result_Message = "ITEM_COMMON4 = " & ITEM_COMMON4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON5.Length > 50 Then
'          Result_Message = "ITEM_COMMON5 = " & ITEM_COMMON5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON6.Length > 50 Then
'          Result_Message = "ITEM_COMMON6 = " & ITEM_COMMON6 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON7.Length > 50 Then
'          Result_Message = "ITEM_COMMON7 = " & ITEM_COMMON7 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON8.Length > 50 Then
'          Result_Message = "ITEM_COMMON8 = " & ITEM_COMMON8 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON9.Length > 50 Then
'          Result_Message = "ITEM_COMMON9 = " & ITEM_COMMON9 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON10.Length > 50 Then
'          Result_Message = "ITEM_COMMON10 = " & ITEM_COMMON10 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON1.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON1 = " & SORT_ITEM_COMMON1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON2.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON2 = " & SORT_ITEM_COMMON2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON3.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON3 = " & SORT_ITEM_COMMON3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON4.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON4 = " & SORT_ITEM_COMMON4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON5.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON5 = " & SORT_ITEM_COMMON5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If STORAGE_TYPE.Length > 50 Then
'          Result_Message = "STORAGE_TYPE = " & ITEM_COMMON1 & " STORAGE_TYPE"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If CheckValueInEnum(Of enuBND)(BND) = False Then
'          Result_Message = "BND = " & BND & " is not defined"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If CheckValueInEnum(Of enuQCStatus)(QC_STATUS) = False Then
'          Result_Message = "QC_STATUS = " & QC_STATUS & " is not defined"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD1.Length > 50 Then
'          Result_Message = "H_POD1 = " & H_POD1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD2.Length > 50 Then
'          Result_Message = "H_POD2 = " & H_POD2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD3.Length > 50 Then
'          Result_Message = "H_POD3 = " & H_POD3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD4.Length > 50 Then
'          Result_Message = "H_POD4 = " & H_POD4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD5.Length > 50 Then
'          Result_Message = "H_POD5 = " & H_POD5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD6.Length > 50 Then
'          Result_Message = "H_POD6 = " & H_POD6 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD7.Length > 50 Then
'          Result_Message = "H_POD7 = " & H_POD7 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD8.Length > 50 Then
'          Result_Message = "H_POD8 = " & H_POD8 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD9.Length > 50 Then
'          Result_Message = "H_POD9 = " & H_POD9 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD10.Length > 50 Then
'          Result_Message = "H_POD10 = " & H_POD10 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD11.Length > 50 Then
'          Result_Message = "H_POD11 = " & H_POD11 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD12.Length > 50 Then
'          Result_Message = "H_POD12 = " & H_POD12 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD13.Length > 50 Then
'          Result_Message = "H_POD13 = " & H_POD13 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD14.Length > 50 Then
'          Result_Message = "H_POD14 = " & H_POD14 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD15.Length > 50 Then
'          Result_Message = "H_POD15 = " & H_POD15 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD16.Length > 50 Then
'          Result_Message = "H_POD16 = " & H_POD16 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD17.Length > 50 Then
'          Result_Message = "H_POD17 = " & H_POD17 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD18.Length > 50 Then
'          Result_Message = "H_POD18 = " & H_POD18 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD19.Length > 50 Then
'          Result_Message = "H_POD19 = " & H_POD19 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD20.Length > 50 Then
'          Result_Message = "H_POD20 = " & H_POD20 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD21.Length > 50 Then
'          Result_Message = "H_POD21 = " & H_POD21 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD22.Length > 50 Then
'          Result_Message = "H_POD22 = " & H_POD22 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD23.Length > 50 Then
'          Result_Message = "H_POD23 = " & H_POD23 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD24.Length > 50 Then
'          Result_Message = "H_POD24 = " & H_POD24 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD25.Length > 50 Then
'          Result_Message = "H_POD25 = " & H_POD25 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        ret_strPreview += " 項次:" & PO_SERIAL_NO & " 料号:" & SKU_NO & " 數量：" & QTY & " Batch：" & LOT_NO & " 料盒編號：" & " Plant：" & OWNER & " StorageLocation：" & SUBOWNER & " CompanyCode：" & ITEM_COMMON2 & " ;"

'        Dim ReceiptCarrierInfo As New MSG_T5F2U5_BatchCreateReceiptByPO.clsBody.clsReceiptCarrierList.clsReceiptCarrierInfo
'        Dim ReceiptInfo As New MSG_T5F2U5_BatchCreateReceiptByPO.clsBody.clsReceiptCarrierList.clsReceiptCarrierInfo.clsReceiptList.clsReceiptInfo
'        ReceiptCarrierInfo.CARRIER_ID = Carrier_id
'        ReceiptCarrierInfo.CARRIER_MODE = enuCarrierMode.Moth_Pallet
'        ReceiptCarrierInfo.CARRIER_TYPE = 1

'        ReceiptInfo.PO_ID = PO_ID
'        ReceiptInfo.PO_SERIAL_NO = PO_SERIAL_NO
'        ReceiptInfo.ITEM_KEY_NO = "NULL"
'        ReceiptInfo.SKU_NO = SKU_NO
'        ReceiptInfo.PACKAGE_ID = "NULL"
'        ReceiptInfo.QTY = QTY
'        ReceiptInfo.LOT_NO = LOT_NO
'        ReceiptInfo.ITEM_COMMON1 = "NULL"
'        ReceiptInfo.ITEM_COMMON2 = "NULL"
'        ReceiptInfo.ITEM_COMMON3 = "NULL"
'        ReceiptInfo.ITEM_COMMON4 = "NULL"
'        ReceiptInfo.ITEM_COMMON5 = "NULL"
'        ReceiptInfo.ITEM_COMMON6 = "NULL"
'        ReceiptInfo.ITEM_COMMON7 = "NULL"
'        ReceiptInfo.ITEM_COMMON8 = "NULL"
'        ReceiptInfo.ITEM_COMMON9 = "NULL"
'        ReceiptInfo.ITEM_COMMON10 = "NULL"
'        ReceiptInfo.SORT_ITEM_COMMON1 = "NULL"
'        ReceiptInfo.SORT_ITEM_COMMON2 = "NULL"
'        ReceiptInfo.SORT_ITEM_COMMON3 = "NULL"
'        ReceiptInfo.SORT_ITEM_COMMON4 = "NULL"
'        ReceiptInfo.SORT_ITEM_COMMON5 = "NULL"
'        ReceiptInfo.LENGTH = "NULL"
'        ReceiptInfo.WIDTH = "NULL"
'        ReceiptInfo.HEIGHT = "NULL"
'        ReceiptInfo.WEIGHT = "NULL"
'        ReceiptInfo.ITEM_VALUE = "NULL"
'        ReceiptInfo.CONTRACT_NO = "NULL"
'        ReceiptInfo.CONTRACT_SERIAL_NO = "NULL"
'        ReceiptInfo.STORAGE_TYPE = "NULL"
'        ReceiptInfo.BND = "NULL"
'        ReceiptInfo.QC_STATUS = "NULL"
'        ReceiptInfo.RECEIPT_DATE = Now_Data
'        ReceiptInfo.MANUFACETURE_DATE = Now_Data
'        ReceiptInfo.EXPIRED_DATE = Now_Data
'        ReceiptInfo.EFFECTIVE_DATE = Now_Data
'        ReceiptInfo.COMMENTS = "NULL"
'        ReceiptCarrierInfo.ReceiptList.ReceiptInfo.Add(ReceiptInfo)
'        MSG_Receipt.Body.ReceiptCarrierList.ReceiptCarrierInfo.Add(ReceiptCarrierInfo)

'      Next


'      Dim dicUUID As New Dictionary(Of String, clsUUID)
'      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
'        Result_Message = "Get UUID False"
'        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If
'      If dicUUID.Any = False Then
'        Result_Message = "Get UUID False"
'        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        Return False
'      End If
'      Dim objUUID = dicUUID.Values(0)
'      Dim UUID = objUUID.Get_NewUUID

'      Dim EventID = "T5F2U5_BatchCreateReceiptByPO"
'      '將單據宜並送給WMS 取得回復為OK後才將單據更新
'      MSG_Receipt.Header = New clsHeader
'      MSG_Receipt.Header.UUID = UUID
'      MSG_Receipt.Header.EventID = EventID
'      MSG_Receipt.Header.Direction = "Primary"
'      MSG_Receipt.Header.ClientInfo = New clsHeader.clsClientInfo
'      MSG_Receipt.Header.ClientInfo.ClientID = "Handler"
'      MSG_Receipt.Header.ClientInfo.UserID = ""
'      MSG_Receipt.Header.ClientInfo.IP = ""
'      MSG_Receipt.Header.ClientInfo.MachineID = ""

'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function



'  '處理 Stocktaking Excel 解析等
'  Private Function I_StocktakingExeclImport(ByVal Receive_Msg As MSG_T10F4U1_MainFileImport,
'                                    ByRef Result_Message As String,
'                                    ByRef FilePath As String,
'                                    ByRef MSG_Receipt_Stocktaking As MSG_T10F2U1_StocktakingManagement,
'                                    ByRef ret_strPreview As String, ByVal workbook As Excel.Workbook) As Boolean
'    Try
'      Dim User_ID = Receive_Msg.Header.ClientInfo.UserID
'      Dim Create_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()

'      '取得所有的盤點單號
'      Dim dicExistStocktaking As New Dictionary(Of String, clsTSTOCKTAKING)
'      gMain.objHandling.O_GetDB_dicStocktakingByAll(dicExistStocktaking)

'      Dim clsStocktakingInfo As New Dictionary(Of String, MSG_T10F2U1_StocktakingManagement.clsBody.clsStocktakingInfo)
'      Dim diclstStocktakingDTLInfo As New Dictionary(Of String, List(Of MSG_T10F2U1_StocktakingManagement.clsBody.clsStocktakingInfo.clsStocktakingDTLList.clsStocktakingDTLInfo))

'      Dim worksheet As Excel.Worksheet 'Worksheet 代表的是 Excel 工作表
'      worksheet = workbook.Worksheets("Data") '讀取其中一張工作表
'      Dim PO_ID As String = ""
'      Dim WO_TYPE As String = ""
'      Dim Stocktaking_ID As String = ""
'      '第二行是盤點資訊
'      For row = 2 To 2
'        If IsNothing(worksheet.Cells(row, 1).Value()) Then Exit For
'        PO_ID = IIf(IsNothing(worksheet.Cells(row, 1).Value()), "", worksheet.Cells(row, 1).Value()) '(單號)	
'        If PO_ID.Length > 50 Then
'          Result_Message = "PO_ID = " & PO_ID & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        'WO_TYPE = IIf(IsNothing(worksheet.Cells(row, 2).Value()), "", worksheet.Cells(row, 2).Value()) '(單據類型)(1入庫、2出庫)
'        'If CheckValueInEnum(Of enuWOType)(WO_TYPE) = False Then
'        '  Result_Message = "WO_TYPE = " & WO_TYPE & " 不符合規定"
'        '  SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        '  Return False
'        'End If

'        'Dim PRIORITY As String = IIf(IsNothing(worksheet.Cells(row, 3).Value()), "", worksheet.Cells(row, 3).Value()) '(優先權)(1~100 越大約優先)
'        'If IsNumeric(PRIORITY) = False OrElse PRIORITY > 100 OrElse PRIORITY < 0 Then
'        '  Result_Message = "PRIORITY = " & WO_TYPE & " 不符合規定"
'        '  SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        '  Return False
'        'End If
'        'Dim Auto_Bound As String = IntegerConvertToBoolean(IIf(IsNothing(worksheet.Cells(row, 4).Value()), "", worksheet.Cells(row, 4).Value()))
'        'Dim SHIPPING_NO As String = IIf(IsNothing(worksheet.Cells(row, 5).Value()), "", worksheet.Cells(row, 5).Value())
'        'Dim WRITE_OFF_NO As String = IIf(IsNothing(worksheet.Cells(row, 6).Value()), "", worksheet.Cells(row, 6).Value())
'        'Dim PO_TYPE1 As String = IIf(WO_TYPE = enuWOType.Receipt, enuPOType_1.Combination_in, enuPOType_1.Picking_out) 'IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        'Dim PO_TYPE2 As String = IIf(IsNothing(worksheet.Cells(row, 8).Value()), "", worksheet.Cells(row, 8).Value())
'        'Dim PO_TYPE3 As String = IIf(IsNothing(worksheet.Cells(row, 9).Value()), "", worksheet.Cells(row, 9).Value())
'        'Dim CUSTOMER_NO As String = IIf(IsNothing(worksheet.Cells(row, 10).Value()), "", worksheet.Cells(row, 10).Value())
'        'Dim CLASS_NO As String = IIf(IsNothing(worksheet.Cells(row, 11).Value()), "", worksheet.Cells(row, 11).Value())
'        'Dim H_PO_ORDER_TYPE As String = IIf(IsNothing(worksheet.Cells(row, 12).Value()), "", worksheet.Cells(row, 12).Value())
'        'Dim H_PO1 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 13).Value()), "", worksheet.Cells(row, 13).Value())
'        'Dim H_PO2 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 14).Value()), "", worksheet.Cells(row, 14).Value())
'        'Dim H_PO3 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 15).Value()), "", worksheet.Cells(row, 15).Value())
'        'Dim H_PO4 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 16).Value()), "", worksheet.Cells(row, 16).Value())
'        'Dim H_PO5 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 17).Value()), "", worksheet.Cells(row, 17).Value())
'        'Dim H_PO6 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 18).Value()), "", worksheet.Cells(row, 18).Value())
'        'Dim H_PO7 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 19).Value()), "", worksheet.Cells(row, 19).Value())
'        'Dim H_PO8 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 20).Value()), "", worksheet.Cells(row, 20).Value())
'        'Dim H_PO9 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 21).Value()), "", worksheet.Cells(row, 21).Value())
'        'Dim H_PO10 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 22).Value()), "", worksheet.Cells(row, 22).Value())
'        'Dim H_PO11 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 23).Value()), "", worksheet.Cells(row, 23).Value())
'        'Dim H_PO12 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 24).Value()), "", worksheet.Cells(row, 24).Value())
'        'Dim H_PO13 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 25).Value()), "", worksheet.Cells(row, 25).Value())
'        'Dim H_PO14 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 26).Value()), "", worksheet.Cells(row, 26).Value())
'        'Dim H_PO15 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 27).Value()), "", worksheet.Cells(row, 27).Value())
'        'Dim H_PO16 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 28).Value()), "", worksheet.Cells(row, 28).Value())
'        'Dim H_PO17 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 29).Value()), "", worksheet.Cells(row, 29).Value())
'        'Dim H_PO18 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 30).Value()), "", worksheet.Cells(row, 30).Value())
'        'Dim H_PO19 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 31).Value()), "", worksheet.Cells(row, 31).Value())
'        'Dim H_PO20 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 32).Value()), "", worksheet.Cells(row, 32).Value())

'        'Dim Start_Time = ""
'        'Dim Finish_Time = ""
'        'Dim H_PO_FINISH_TIME = ""

'        'If WO_TYPE = enuWOType.Receipt Then
'        '  ret_strPreview += "建立入庫單據，單號：" & PO_ID
'        '  PO_TYPE2 = enuPOType_2.m_general_in
'        '  H_PO_ORDER_TYPE = enuOrderType.m_general_in
'        'Else
'        '  ret_strPreview += "建立出庫單據，單號：" & PO_ID
'        '  PO_TYPE2 = enuPOType_2.m_grneral_out
'        '  H_PO_ORDER_TYPE = enuOrderType.m_grneral_out
'        'End If

'        '盤點單號 = 單號+盤點類型
'        Stocktaking_ID = PO_ID

'        'ERP盤點單會有多項次
'        '檢查盤點單號是否已存在
'        If dicExistStocktaking.ContainsKey(clsTSTOCKTAKING.Get_Combination_Key(Stocktaking_ID)) = True Then
'          Result_Message = "AccountingDocument already exists, AccountingDocument=" & Stocktaking_ID
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If


'        '建立盤點單
'        Dim StocktakingInfo As New MSG_T10F2U1_StocktakingManagement.clsBody.clsStocktakingInfo
'        Stocktaking_ID = PO_ID
'        StocktakingInfo.STOCKTAKING_ID = Stocktaking_ID
'        StocktakingInfo.STOCKTAKING_TYPE1 = enuStockTaking_Type1.Manual
'        StocktakingInfo.STOCKTAKING_TYPE2 = enuStockTaking_Type2.none
'        StocktakingInfo.STOCKTAKING_TYPE3 = enuStockTaking_Type3.none
'        If clsStocktakingInfo.ContainsKey(StocktakingInfo.STOCKTAKING_ID) = False Then
'          clsStocktakingInfo.Add(StocktakingInfo.STOCKTAKING_ID, StocktakingInfo)
'        End If

'      Next

'      'PO_DTL
'      For row = 4 To worksheet.Rows.Count - 1
'        If IsNothing(worksheet.Cells(row, 1).Value()) Then Exit For
'        ' worksheet.Cells(1, row).Value() '讀取某一個欄位的值，第一個數字是行，第二個數字是列，如果欄位沒有東西會回傳Nothing

'        Dim PO_SERIAL_NO As String = IIf(IsNothing(worksheet.Cells(row, 1).Value()), "", worksheet.Cells(row, 1).Value())
'        Dim SKU_NO As String = IIf(IsNothing(worksheet.Cells(row, 2).Value()), "", worksheet.Cells(row, 2).Value())
'        SKU_NO = SKU_NO.Trim()
'        Dim LOT_NO As String = IIf(IsNothing(worksheet.Cells(row, 3).Value()), "", worksheet.Cells(row, 3).Value())
'        LOT_NO = LOT_NO.Trim()
'        Dim QTY As String = IIf(IsNothing(worksheet.Cells(row, 4).Value()), "", worksheet.Cells(row, 4).Value())
'        Dim OWNER As String = IIf(IsNothing(worksheet.Cells(row, 5).Value()), "", worksheet.Cells(row, 5).Value())
'        OWNER = OWNER.Trim()
'        Dim SUBOWNER As String = IIf(IsNothing(worksheet.Cells(row, 6).Value()), "", worksheet.Cells(row, 6).Value())
'        SUBOWNER = SUBOWNER.Trim()
'        Dim PACKAGE_ID As String = "" ' IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        Dim ITEM_COMMON1 As String = "ERP" 'IIf(IsNothing(worksheet.Cells(row, 8).Value()), "", worksheet.Cells(row, 8).Value())
'        Dim ITEM_COMMON2 As String = IIf(IsNothing(worksheet.Cells(row, 7).Value()), "", worksheet.Cells(row, 7).Value())
'        ITEM_COMMON2 = ITEM_COMMON2.Trim()
'        Dim ITEM_COMMON3 As String = "B6-00-00" 'IIf(IsNothing(worksheet.Cells(row, 10).Value()), "", worksheet.Cells(row, 10).Value())
'        Dim ITEM_COMMON4 As String = IIf(IsNothing(worksheet.Cells(row, 11).Value()), "", worksheet.Cells(row, 11).Value())
'        Dim ITEM_COMMON5 As String = IIf(IsNothing(worksheet.Cells(row, 12).Value()), "", worksheet.Cells(row, 12).Value())
'        Dim ITEM_COMMON6 As String = IIf(IsNothing(worksheet.Cells(row, 13).Value()), "", worksheet.Cells(row, 13).Value())
'        Dim ITEM_COMMON7 As String = IIf(IsNothing(worksheet.Cells(row, 14).Value()), "", worksheet.Cells(row, 14).Value())
'        Dim ITEM_COMMON8 As String = IIf(IsNothing(worksheet.Cells(row, 15).Value()), "", worksheet.Cells(row, 15).Value())
'        Dim ITEM_COMMON9 As String = IIf(IsNothing(worksheet.Cells(row, 16).Value()), "", worksheet.Cells(row, 16).Value())
'        Dim ITEM_COMMON10 As String = IIf(IsNothing(worksheet.Cells(row, 17).Value()), "", worksheet.Cells(row, 17).Value())
'        Dim SORT_ITEM_COMMON1 As String = IIf(IsNothing(worksheet.Cells(row, 18).Value()), "", worksheet.Cells(row, 18).Value())
'        Dim SORT_ITEM_COMMON2 As String = IIf(IsNothing(worksheet.Cells(row, 19).Value()), "", worksheet.Cells(row, 19).Value())
'        Dim SORT_ITEM_COMMON3 As String = IIf(IsNothing(worksheet.Cells(row, 20).Value()), "", worksheet.Cells(row, 20).Value())
'        Dim SORT_ITEM_COMMON4 As String = IIf(IsNothing(worksheet.Cells(row, 21).Value()), "", worksheet.Cells(row, 21).Value())
'        Dim SORT_ITEM_COMMON5 As String = IIf(IsNothing(worksheet.Cells(row, 22).Value()), "", worksheet.Cells(row, 22).Value())
'        Dim STORAGE_TYPE As String = enuStorageType.Store
'        Dim BND = enuBND.NB ' IIf(IsNothing(worksheet.Cells(row, 23).Value()), "", worksheet.Cells(row, 23).Value())
'        Dim QC_STATUS = enuQCStatus.NULL ' IIf(IsNothing(worksheet.Cells(row, 24).Value()), "", worksheet.Cells(row, 24).Value())
'        Dim H_POD1 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 25).Value()), "", worksheet.Cells(row, 25).Value())
'        Dim H_POD2 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 26).Value()), "", worksheet.Cells(row, 26).Value())
'        Dim H_POD3 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 27).Value()), "", worksheet.Cells(row, 27).Value())
'        Dim H_POD4 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 28).Value()), "", worksheet.Cells(row, 28).Value())
'        Dim H_POD5 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 29).Value()), "", worksheet.Cells(row, 29).Value())
'        Dim H_POD6 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 30).Value()), "", worksheet.Cells(row, 30).Value())
'        Dim H_POD7 As String = ITEM_COMMON3 'IIf(IsNothing(worksheet.Cells(row, 31).Value()), "", worksheet.Cells(row, 31).Value())
'        Dim H_POD8 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 32).Value()), "", worksheet.Cells(row, 32).Value())
'        Dim H_POD9 As String = "" 'IIf(IsNothing(worksheet.Cells(row, 33).Value()), "", worksheet.Cells(row, 33).Value())
'        Dim H_POD10 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 34).Value()), "", worksheet.Cells(row, 34).Value())
'        Dim H_POD11 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 35).Value()), "", worksheet.Cells(row, 35).Value())
'        Dim H_POD12 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 36).Value()), "", worksheet.Cells(row, 36).Value())
'        Dim H_POD13 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 37).Value()), "", worksheet.Cells(row, 37).Value())
'        Dim H_POD14 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 38).Value()), "", worksheet.Cells(row, 38).Value())
'        Dim H_POD15 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 39).Value()), "", worksheet.Cells(row, 39).Value())
'        Dim H_POD16 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 40).Value()), "", worksheet.Cells(row, 40).Value())
'        Dim H_POD17 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 41).Value()), "", worksheet.Cells(row, 41).Value())
'        Dim H_POD18 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 42).Value()), "", worksheet.Cells(row, 42).Value())
'        Dim H_POD19 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 43).Value()), "", worksheet.Cells(row, 43).Value())
'        Dim H_POD20 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 44).Value()), "", worksheet.Cells(row, 44).Value())
'        Dim H_POD21 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 45).Value()), "", worksheet.Cells(row, 45).Value())
'        Dim H_POD22 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 46).Value()), "", worksheet.Cells(row, 46).Value())
'        Dim H_POD23 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 47).Value()), "", worksheet.Cells(row, 47).Value())
'        Dim H_POD24 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 48).Value()), "", worksheet.Cells(row, 48).Value())
'        Dim H_POD25 As String = "" ' IIf(IsNothing(worksheet.Cells(row, 49).Value()), "", worksheet.Cells(row, 49).Value())
'        Dim COMMENTS = ""
'        Dim QTY_PROCESS = 0
'        Dim QTY_FINISH = 0
'        Dim PO_LINE_NO = PO_SERIAL_NO

'        Dim FROM_OWNER_ID = ""
'        Dim FROM_SUB_OWNER_ID = ""
'        Dim TO_OWNER_ID = ""
'        Dim TO_SUB_OWNER_ID = ""
'        '根據入出庫類型填入對應資訊
'        'If WO_TYPE = enuWOType.Receipt Then
'        '  TO_OWNER_ID = OWNER
'        '  TO_SUB_OWNER_ID = SUBOWNER
'        'Else
'        '  FROM_OWNER_ID = OWNER
'        '  FROM_SUB_OWNER_ID = SUBOWNER
'        'End If
'        Dim FACTORY_ID = ""
'        Dim DEST_AREA_ID = ""
'        Dim DEST_LOCATION_ID = ""
'        Dim H_POD_STEP_NO = enuPOStatus.Queued
'        Dim H_POD_MOVE_TYPE = ""
'        Dim H_POD_FINISH_TIME = ""
'        Dim H_POD_BILLING_DATE = ""
'        Dim H_POD_CREATE_TIME = Create_Time
'        Dim PODTL_STATUS = enuPODTLStatus.Queued


'        If PO_SERIAL_NO.Length > 50 Then
'          Result_Message = "PO_SERIAL_NO = " & PO_SERIAL_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SKU_NO.Length > 50 Then
'          Result_Message = "SKU_NO = " & SKU_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If LOT_NO.Length > 50 Then
'          Result_Message = "LOT_NO = " & LOT_NO & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If IntegerCheckPositive(QTY) = False Then
'          Result_Message = "QTY = " & QTY & " 格式錯誤"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If OWNER.Length > 50 Then
'          Result_Message = "OWNER = " & OWNER & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SUBOWNER.Length > 50 Then
'          Result_Message = "SUBOWNER = " & SUBOWNER & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If PACKAGE_ID.Length > 50 Then
'          Result_Message = "PACKAGE_ID = " & PACKAGE_ID & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON1.Length > 50 Then
'          Result_Message = "ITEM_COMMON1 = " & ITEM_COMMON1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON2.Length > 50 Then
'          Result_Message = "ITEM_COMMON2 = " & ITEM_COMMON2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON3.Length > 50 Then
'          Result_Message = "ITEM_COMMON3 = " & ITEM_COMMON3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON4.Length > 50 Then
'          Result_Message = "ITEM_COMMON4 = " & ITEM_COMMON4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON5.Length > 50 Then
'          Result_Message = "ITEM_COMMON5 = " & ITEM_COMMON5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON6.Length > 50 Then
'          Result_Message = "ITEM_COMMON6 = " & ITEM_COMMON6 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON7.Length > 50 Then
'          Result_Message = "ITEM_COMMON7 = " & ITEM_COMMON7 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON8.Length > 50 Then
'          Result_Message = "ITEM_COMMON8 = " & ITEM_COMMON8 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON9.Length > 50 Then
'          Result_Message = "ITEM_COMMON9 = " & ITEM_COMMON9 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If ITEM_COMMON10.Length > 50 Then
'          Result_Message = "ITEM_COMMON10 = " & ITEM_COMMON10 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON1.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON1 = " & SORT_ITEM_COMMON1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON2.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON2 = " & SORT_ITEM_COMMON2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON3.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON3 = " & SORT_ITEM_COMMON3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON4.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON4 = " & SORT_ITEM_COMMON4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If SORT_ITEM_COMMON5.Length > 50 Then
'          Result_Message = "SORT_ITEM_COMMON5 = " & SORT_ITEM_COMMON5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If STORAGE_TYPE.Length > 50 Then
'          Result_Message = "STORAGE_TYPE = " & ITEM_COMMON1 & " STORAGE_TYPE"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If CheckValueInEnum(Of enuBND)(BND) = False Then
'          Result_Message = "BND = " & BND & " is not defined"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If CheckValueInEnum(Of enuQCStatus)(QC_STATUS) = False Then
'          Result_Message = "QC_STATUS = " & QC_STATUS & " is not defined"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD1.Length > 50 Then
'          Result_Message = "H_POD1 = " & H_POD1 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD2.Length > 50 Then
'          Result_Message = "H_POD2 = " & H_POD2 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD3.Length > 50 Then
'          Result_Message = "H_POD3 = " & H_POD3 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD4.Length > 50 Then
'          Result_Message = "H_POD4 = " & H_POD4 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD5.Length > 50 Then
'          Result_Message = "H_POD5 = " & H_POD5 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD6.Length > 50 Then
'          Result_Message = "H_POD6 = " & H_POD6 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD7.Length > 50 Then
'          Result_Message = "H_POD7 = " & H_POD7 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD8.Length > 50 Then
'          Result_Message = "H_POD8 = " & H_POD8 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD9.Length > 50 Then
'          Result_Message = "H_POD9 = " & H_POD9 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD10.Length > 50 Then
'          Result_Message = "H_POD10 = " & H_POD10 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD11.Length > 50 Then
'          Result_Message = "H_POD11 = " & H_POD11 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD12.Length > 50 Then
'          Result_Message = "H_POD12 = " & H_POD12 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD13.Length > 50 Then
'          Result_Message = "H_POD13 = " & H_POD13 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD14.Length > 50 Then
'          Result_Message = "H_POD14 = " & H_POD14 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD15.Length > 50 Then
'          Result_Message = "H_POD15 = " & H_POD15 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD16.Length > 50 Then
'          Result_Message = "H_POD16 = " & H_POD16 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD17.Length > 50 Then
'          Result_Message = "H_POD17 = " & H_POD17 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD18.Length > 50 Then
'          Result_Message = "H_POD18 = " & H_POD18 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD19.Length > 50 Then
'          Result_Message = "H_POD19 = " & H_POD19 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD20.Length > 50 Then
'          Result_Message = "H_POD20 = " & H_POD20 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD21.Length > 50 Then
'          Result_Message = "H_POD21 = " & H_POD21 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD22.Length > 50 Then
'          Result_Message = "H_POD22 = " & H_POD22 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD23.Length > 50 Then
'          Result_Message = "H_POD23 = " & H_POD23 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD24.Length > 50 Then
'          Result_Message = "H_POD24 = " & H_POD24 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If H_POD25.Length > 50 Then
'          Result_Message = "H_POD25 = " & H_POD25 & " 長度超過限制"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If

'        'Dim CLOSE_ABLE = False
'        'Dim objNewPO_DTL As New clsPO_DTL(PO_ID, PO_LINE_NO, PO_SERIAL_NO, SKU_NO, LOT_NO, QTY, QTY_PROCESS, QTY_FINISH, COMMENTS, PACKAGE_ID, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6,
'        '                     ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, STORAGE_TYPE, BND, QC_STATUS,
'        '                     FROM_OWNER_ID, FROM_SUB_OWNER_ID, TO_OWNER_ID, TO_SUB_OWNER_ID, FACTORY_ID, DEST_AREA_ID, DEST_LOCATION_ID, H_POD_STEP_NO, H_POD_MOVE_TYPE, H_POD_FINISH_TIME, H_POD_BILLING_DATE,
'        '                     H_POD_CREATE_TIME, H_POD1, H_POD2, H_POD3, H_POD4, H_POD5, H_POD6, H_POD7, H_POD8, H_POD9, H_POD10, H_POD11, H_POD12, H_POD13, H_POD14, H_POD15, H_POD16, H_POD17, H_POD18, H_POD19,
'        '                     H_POD20, H_POD21, H_POD22, H_POD23, H_POD24, H_POD25, PODTL_STATUS, CLOSE_ABLE)
'        'If dicAddPO_DTL.ContainsKey(objNewPO_DTL.gid) = False Then
'        '  dicAddPO_DTL.Add(objNewPO_DTL.gid, objNewPO_DTL)
'        'Else
'        '  Result_Message = "存在相同訂單項次，訂單號 = " & PO_ID & " 項次 = " & PO_SERIAL_NO
'        '  SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        '  Return False
'        'End If

'        'Dim objNewPO_LINE As New clsPO_LINE(PO_ID, PO_LINE_NO, QTY, QTY_FINISH, 0, "", "", "", "", "")
'        'If dicAdd_PO_Line.ContainsKey(objNewPO_LINE.gid) = False Then
'        '  dicAdd_PO_Line.Add(objNewPO_LINE.gid, objNewPO_LINE)
'        'Else
'        '  Result_Message = "存在相同訂單項次，訂單號 = " & PO_ID & " 項次 = " & PO_SERIAL_NO
'        '  SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'        '  Return False
'        'End If



'        ret_strPreview += " 項次:" & PO_SERIAL_NO & " 料号:" & SKU_NO & " Batch：" & LOT_NO & " Plant：" & OWNER & " StorageLocation：" & SUBOWNER & " CompanyCode：" & ITEM_COMMON2 & " ;"




'        Dim CompanyCode = "1700" '上位系統不會丟過來，目前
'        Dim OWNER_NO = OWNER
'        Dim SUB_OWNER_NO = SUBOWNER
'        Dim STOCKTAKING_SERIAL_NO = PO_SERIAL_NO

'        Dim clsStocktakingDTLInfo As New MSG_T10F2U1_StocktakingManagement.clsBody.clsStocktakingInfo.clsStocktakingDTLList.clsStocktakingDTLInfo

'        clsStocktakingDTLInfo.STOCKTAKING_SERIAL_NO = STOCKTAKING_SERIAL_NO 'ERP來的盤點單一次一項
'        clsStocktakingDTLInfo.BLOCK_NO = ""
'        clsStocktakingDTLInfo.SKU_NO = SKU_NO
'        clsStocktakingDTLInfo.OWNER_NO = OWNER_NO
'        clsStocktakingDTLInfo.SUB_OWNER_NO = SUB_OWNER_NO
'        clsStocktakingDTLInfo.STORAGE_TYPE = enuStorageType.Store
'        clsStocktakingDTLInfo.BND = BooleanConvertToInteger(False)
'        clsStocktakingDTLInfo.CARRIER_ID = ""
'        clsStocktakingDTLInfo.PERCENTAGE = "100" '比例
'        clsStocktakingDTLInfo.LOT_NO = LOT_NO
'        clsStocktakingDTLInfo.ITEM_COMMON1 = ITEM_COMMON1
'        clsStocktakingDTLInfo.ITEM_COMMON2 = ITEM_COMMON2
'        clsStocktakingDTLInfo.ITEM_COMMON3 = ITEM_COMMON3
'        clsStocktakingDTLInfo.ITEM_COMMON4 = ""
'        clsStocktakingDTLInfo.ITEM_COMMON5 = ""
'        clsStocktakingDTLInfo.ITEM_COMMON6 = ""
'        clsStocktakingDTLInfo.ITEM_COMMON7 = ""
'        clsStocktakingDTLInfo.ITEM_COMMON8 = ""
'        clsStocktakingDTLInfo.ITEM_COMMON9 = ""
'        clsStocktakingDTLInfo.ITEM_COMMON10 = ""
'        clsStocktakingDTLInfo.SORT_ITEM_COMMON1 = ""
'        clsStocktakingDTLInfo.SORT_ITEM_COMMON2 = ""
'        clsStocktakingDTLInfo.SORT_ITEM_COMMON3 = ""
'        clsStocktakingDTLInfo.SORT_ITEM_COMMON4 = ""
'        clsStocktakingDTLInfo.SORT_ITEM_COMMON5 = ""
'        clsStocktakingDTLInfo.RECEIPT_DATE = ""
'        clsStocktakingDTLInfo.ERP_QTY = QTY '上位系統的數量
'        If diclstStocktakingDTLInfo.ContainsKey(Stocktaking_ID) = False Then
'          diclstStocktakingDTLInfo.Add(Stocktaking_ID, New List(Of MSG_T10F2U1_StocktakingManagement.clsBody.clsStocktakingInfo.clsStocktakingDTLList.clsStocktakingDTLInfo)({clsStocktakingDTLInfo}))
'        Else
'          diclstStocktakingDTLInfo.Item(Stocktaking_ID).Add(clsStocktakingDTLInfo)
'        End If
'      Next

'      '根據單號排序、組事件
'      For Each objStocktakingInfo In clsStocktakingInfo.Values
'        '建立command
'        '取得流水號
'        Dim dicUUID As New Dictionary(Of String, clsUUID)
'        If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
'          Result_Message = "Get UUID False"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        If dicUUID.Any = False Then
'          Result_Message = "Get UUID False"
'          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
'          Return False
'        End If
'        Dim objUUID = dicUUID.Values(0)

'        Dim UUID = objUUID.Get_NewUUID
'        '將單據宜並送給WMS 取得回復為OK後才將單據更新
'        MSG_Receipt_Stocktaking.Header = New clsHeader
'        MSG_Receipt_Stocktaking.Header.UUID = UUID
'        MSG_Receipt_Stocktaking.Header.EventID = "T10F2U1_StocktakingManagement"
'        MSG_Receipt_Stocktaking.Header.Direction = "Primary"
'        MSG_Receipt_Stocktaking.Header.ClientInfo = New clsHeader.clsClientInfo
'        MSG_Receipt_Stocktaking.Header.ClientInfo.ClientID = "Handler"
'        MSG_Receipt_Stocktaking.Header.ClientInfo.UserID = User_ID
'        MSG_Receipt_Stocktaking.Header.ClientInfo.IP = ""
'        MSG_Receipt_Stocktaking.Header.ClientInfo.MachineID = ""
'        MSG_Receipt_Stocktaking.Body = New MSG_T10F2U1_StocktakingManagement.clsBody
'        MSG_Receipt_Stocktaking.Body.Action = enuAction.Create.ToString
'        MSG_Receipt_Stocktaking.Body.StocktakingInfo.STOCKTAKING_ID = objStocktakingInfo.STOCKTAKING_ID
'        MSG_Receipt_Stocktaking.Body.StocktakingInfo.STOCKTAKING_TYPE1 = objStocktakingInfo.STOCKTAKING_TYPE1
'        MSG_Receipt_Stocktaking.Body.StocktakingInfo.STOCKTAKING_TYPE2 = objStocktakingInfo.STOCKTAKING_TYPE2
'        MSG_Receipt_Stocktaking.Body.StocktakingInfo.STOCKTAKING_TYPE3 = objStocktakingInfo.STOCKTAKING_TYPE3

'        '表身
'        For Each objStocktakingDTLInfo In diclstStocktakingDTLInfo.Item(objStocktakingInfo.STOCKTAKING_ID)
'          MSG_Receipt_Stocktaking.Body.StocktakingInfo.StocktakingDTLList.StocktakingDTLInfo.Add(objStocktakingDTLInfo)
'        Next


'      Next



'      Return True
'    Catch ex As Exception
'      Result_Message = ex.ToString
'      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
'      Return False
'    End Try
'  End Function


'End Module
