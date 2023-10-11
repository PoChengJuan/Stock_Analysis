Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports eCA_HostObject
Imports eCA_TransactionMessage

''' <summary>
''' 20181117
''' V1.0.0
''' Mark
''' 处理所有自动处理的程序
''' </summary>
Module Module_Auto_Excute
  '自动执行
  Public Sub O_thr_Auto_Excute()
    Const SleepTime As Integer = 1000 '3600000 '一小時
    Dim AutoDownLoad_SKU_Timer As Long = 0 '料品計數器，每秒加一次
    Dim AutoDownLoad_PO_Timer As Long = 0 '單據計數器，每秒加一次
    'Dim AutoStocktakingDate = "-1" '过账日期 

    '載入設定 若更新設定則須重新載入
    Dim objBusiness As clsBusiness_Rule = Nothing
    If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.Report_Time, objBusiness) Then
      For Each _ReportTime In objBusiness.Rule_Value.Split(",")
        ReportTime.Add(_ReportTime)
      Next
    End If
    Dim int_I_Clear_Out_Time_CT_PO_DTL_coount = 0

    While True
      Try

        Dim dictest As New Dictionary(Of String, clsPO)



        If AutoDownLoad_SKU_Timer >= DownLoad_Interval_SKU And f_AutoDownLoad = True Then


          '這裡的作法，假如上面進了CATCH，下一秒會馬上再跑一次，一直到正常作完為止。
          '要是DB的資料有誤或壞掉，就會有大量的FALSE造成的LOG紀錄
          AutoDownLoad_SKU_Timer = 0
        Else
          AutoDownLoad_SKU_Timer += 1
        End If

        If AutoDownLoad_PO_Timer >= DownLoad_Interval_PO And f_AutoDownLoad = True Then
          '單據主檔 每一分鐘處理一次
          Dim strResultMsg As String = ""
          '出通單
          Dim dicEPSXB As New Dictionary(Of String, clsEPSXB)
          If gMain.objHandling.O_GetDB_dicEPSXBByXB010_IS_ZERO(dicEPSXB) = True Then
            Dim dicPO_ID As New Dictionary(Of String, String)
            For Each objEPSXB In dicEPSXB.Values
              Dim PO_ID = objEPSXB.XB001 & "_" & objEPSXB.XB002 & "_" & objEPSXB.XB010 & "_" & objEPSXB.XB015
              If dicPO_ID.ContainsKey(PO_ID) = False Then
                dicPO_ID.Add(PO_ID, PO_ID)
              End If
            Next
            Module_POManagement_HTG_Sell.O_POManagement_HTG_Sell(dicPO_ID, dicEPSXB, strResultMsg)
          End If



          '領料單
          Dim dicMOCXD As New Dictionary(Of String, clsMOCXD)
          If gMain.objHandling.O_GetDB_dicMOCXDByXD011_IS_ZERO(dicMOCXD) = True Then
            Dim dicPO_ID As New Dictionary(Of String, String)
            For Each objMOCXD In dicMOCXD.Values
              Dim PO_ID = objMOCXD.XD001 & "_" & objMOCXD.XD002 & "_" & objMOCXD.XD011
              If dicPO_ID.ContainsKey(PO_ID) = False Then
                dicPO_ID.Add(PO_ID, PO_ID)
              End If
            Next
            Module_POManagement_HTG_Pickup.O_POManagement_HTG_Pickup(dicPO_ID, dicMOCXD, strResultMsg)
          End If

          '轉撥單
          Dim dicINVXF As New Dictionary(Of String, clsINVXF)
          If gMain.objHandling.O_GetDB_dicINVXFByXF009_IS_ZERO(dicINVXF) = True Then
            Dim dicPO_ID As New Dictionary(Of String, String)
            For Each objINVXF In dicINVXF.Values
              Dim PO_ID = objINVXF.XF001 & "_" & objINVXF.XF002 & "_" & objINVXF.XF009
              If dicPO_ID.ContainsKey(PO_ID) = False Then
                dicPO_ID.Add(PO_ID, PO_ID)
              End If
            Next
            Module_POManagement_HTG_Transfer.O_POManagement_HTG_Transfer(dicPO_ID, dicINVXF, strResultMsg)
          End If
          '這裡的作法，假如上面進了CATCH，下一秒會馬上再跑一次，一直到正常作完為止。
          '要是DB的資料有誤或壞掉，就會有大量的FALSE造成的LOG紀錄
          AutoDownLoad_PO_Timer = 0
        Else
          AutoDownLoad_PO_Timer += 1
        End If

        'int_I_Clear_Out_Time_CT_PO_DTL_coount += 1
        'If int_I_Clear_Out_Time_CT_PO_DTL_coount >= 60 Then '一小時一次
        '  ''環鴻ERP單據 多餘的排除
        '  I_Clear_Out_Time_CT_PO_DTL()
        '  int_I_Clear_Out_Time_CT_PO_DTL_coount = 0
        'End If

        ''一分鐘一次
        ''發送對帳後的信件 '有on再送 '先產生excel 再發
        'If bln_SendAccountMail = True Then
        '  If I_Auto_Send_Account_Mail() = True Then
        '    bln_SendAccountMail = False '發送成功則off掉
        '  End If
        'End If

        ''檢查入庫短缺、結單短缺 定時回報 一天一次
        'If CkeckClockOn(enuSystemStatus.LastReportTime, ReportTime) Then
        '  I_Auto_Report_Inbound_Lack() '回報入庫短缺
        'End If



        ''執行上報的生產線生產資訊收集，並把數量配置到工單中
        'I_Auto_SetProductionInfo()
        ''執行保養資訊的更新
        'I_Auto_SetMaintenanceStatus()
        ''檢查是否需要發送保養異常
        'I_Auto_SetLineInfo()
        ''進行班次生產數量的Reset
        'I_Auto_ResetLineProductionByClass()

      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_Auto_Excute End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub

  Public Sub O_thr_DownLoad_MainFile()
    Const SleepTime As Integer = 1000 '3600000 '一小時
    Dim AutoDownLoad_SKU_Timer As Long = 0 '料品計數器，每秒加一次
    Dim AutoDownLoad_PO_Timer As Long = 0 '單據計數器，每秒加一次
    'Dim AutoStocktakingDate = "-1" '过账日期 

    '載入設定 若更新設定則須重新載入
    Dim objBusiness As clsBusiness_Rule = Nothing
    If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.Report_Time, objBusiness) Then
      For Each _ReportTime In objBusiness.Rule_Value.Split(",")
        ReportTime.Add(_ReportTime)
      Next
    End If
    Dim int_I_Clear_Out_Time_CT_PO_DTL_coount = 0

    While True
      Try

        Dim dictest As New Dictionary(Of String, clsPO)



        If AutoDownLoad_SKU_Timer >= DownLoad_Interval_SKU And f_AutoDownLoad = True Then


          '這裡的作法，假如上面進了CATCH，下一秒會馬上再跑一次，一直到正常作完為止。
          '要是DB的資料有誤或壞掉，就會有大量的FALSE造成的LOG紀錄
          AutoDownLoad_SKU_Timer = 0
          Dim strResultMsg As String = ""
          Dim dicINVXB As New Dictionary(Of String, clsINVXB)
          If gMain.objHandling.O_GetDB_dicINVXBByXB008_IS_ZERO(dicINVXB) = True Then
            SendMessageToLog("開始處理料號下載，筆數：" & dicINVXB.Count, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            Module_SKUManagement_INVXB.O_SKUManagement(dicINVXB, strResultMsg)
          End If
        Else
          AutoDownLoad_SKU_Timer += 1
        End If




      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_DownLoad_MainFile End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub


  Private Sub I_Clear_Out_Time_CT_PO_DTL()
    SendMessageToLog("Start Process I_Clear_Out_Time_CT_PO_DTL", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)             'Vito_12b30
    SyncLock gMain.objHandling.objCT_PO_DTLLock
      SendMessageToLog("Auto_Excute objCT_PO_DTLLock Locked", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)  'Vito_12b30
      Try
        SendMessageToLog("I_Clear_Out_Time_CT_PO_DTL", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        Dim newTime = GetNewTime_DBFormat()
        '刪除超過固定小時的單據
        Dim Over_Hour = 96 '暫定72小時
        Dim dicDeleteCT_PO_DTL As New Dictionary(Of String, clsHOST_CT_TMP_PO_DTL)
        Dim lstSQL As New List(Of String)

        For Each objCT_PO_DTL In gMain.objHandling.gdicCT_PO_DTL.Values
          '-有時間格式限制
          Dim _nowTime As DateTime = DateTime.Parse(newTime)
          Dim _beforeTime As DateTime = DateTime.Parse(objCT_PO_DTL.CREATE_TIME)
          Dim ts As TimeSpan = _nowTime.Subtract(_beforeTime)
          If ts.TotalHours > Over_Hour Then '大於時間 可以刪除
            If dicDeleteCT_PO_DTL.ContainsKey(objCT_PO_DTL.gid) = False Then
              dicDeleteCT_PO_DTL.Add(objCT_PO_DTL.gid, objCT_PO_DTL)
            End If
          End If
        Next

        For Each obj In dicDeleteCT_PO_DTL.Values
          obj.O_Add_Delete_SQLString(lstSQL)
          '移除記憶體中的資料
          Dim objDelete As clsHOST_CT_TMP_PO_DTL = Nothing
          If gMain.objHandling.gdicCT_PO_DTL.TryGetValue(obj.gid, objDelete) Then
            objDelete.Remove_Relationship()
          End If
        Next
        If lstSQL.Any = True Then
          Common_DBManagement.BatchUpdate(lstSQL)
        End If

      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      End Try
    End SyncLock
    SendMessageToLog("End Process I_Clear_Out_Time_CT_PO_DTL End", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)  'Vito_12b30
  End Sub
  Private Function I_Auto_Send_Account_Mail() As Boolean
    Try
      SendMessageToLog("I_Auto_Send_Account_Mail", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      If gMain.objHandling.O_GetDB_INVENTORY_COMPARISON_GetEndFlagByDB = False Then
        Return False
      End If
      Dim newTime = GetNewTime_DBFormat()
      Dim bln_sendFile = True '是否傳送檔案
      Dim str_FilePath As String = "" '要發送的檔案位置
      '1. 產生檔案

      Dim csv_ExportPath = "<CSVOption>
	<FilePath>C:\Log\Export</FilePath>
	<SQList>
		<SQLInfo>
			<FileName>SAP_自動倉庫存比對差異報表</FileName>
			<SQL>select a.ACC_COMMON1 as 棧板編號,c.Location_ALIS as 位置,a.ACC_COMMON3 as 料籃數量,a.SKU_NO as 物料編號,a.LOT_NO as Batch,a.OWNER_NO as 廠別,a.SUB_OWNER_NO as 庫別 ,TRIM(to_char(a.ERP_STOCK_QTY,'999,999,999,990')) as SAP帳上數量,TRIM(to_char(a.ERP_UNFINISH_QTY,'999,999,999,990')) as SAP已銷未發,TRIM(to_char(a.ERP_COMPARSON_QTY,'999,999,999,990')) as SAP比對數量,TRIM(to_char(a.WMS_STOCK_QTY,'999,999,999,990')) as 自動倉帳上數量,TRIM(to_char(a.WMS_UNFINISH_QTY,'999,999,999,990')) as 自動倉已發未結,TRIM(to_char(a.WMS_COMPARSON_QTY,'999,999,999,990')) as 自動倉比對數量,TRIM(to_char(a.QUANTITY_VARIANCE,'999,999,999,990')) as 差異數量,b.sku_desc as 物料名稱 from WMS_CT_INVENTORY_COMPARISON a inner join WMS_M_SKU b on a.SKU_NO=b.sku_no inner join WMS_M_LOCATION c on c.Location_no = a.ACC_COMMON2 order by a.QUANTITY_VARIANCE</SQL>
			<DateTimeRange>
				<FieldName/>
				<HourPeriod>24</HourPeriod>
			</DateTimeRange>
		</SQLInfo>
	</SQList>
</CSVOption>"

      Dim objCsvInfo As CSVOption = Nothing
      If I_Load_EXPORT(csv_ExportPath, objCsvInfo) = True Then
        For Each info As CSVOption.SQLInfo In objCsvInfo.SQList
          Dim sdate = $"{Now.AddHours(-info.DateTimeRange.HourPeriod):yyyy/MM/dd HH:mm:ss}"
          Dim edate = $"{Now:yyyy/MM/dd HH:mm:ss}"
          ' get data
          Dim dt As DataTable = GetCSVData(info.SQL, info.DateTimeRange.FieldName, sdate, edate)
          If dt.Columns.Count = 0 Then '如果0行 則不發送檔案
            bln_sendFile = False
            Return False
          End If
          ' create csv file
          CreateCSVFile(dt, info, objCsvInfo, str_FilePath)
        Next

        SendMessageToLog("Export completed.", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      End If

      '2. 發送檔案
      Dim objT1F1M1_SendMessage As New MSG_T1F1M1_SendMessage

      '發送的訊息
      Dim objMessageList As New MSG_T1F1M1_SendMessage.BodyInfo.MessageDataList
      Dim objMessageInfo As New MSG_T1F1M1_SendMessage.BodyInfo.MessageDataInfo
      objMessageInfo.MESSAGE_TITLE = "SAP_自動倉庫存比對差異報表"
      objMessageInfo.MESSAGE_TEXT = "Dear Sir,
附件為 SAP_自動倉庫存比對差異報表 " & newTime


      objMessageInfo.ATTACHMENT_PATH = IIf(bln_sendFile, str_FilePath, "")
      objMessageList.MessageInfo.Add(objMessageInfo) '僅一組 寫死

      '發送的人員
      Dim objSendDataList As New MSG_T1F1M1_SendMessage.BodyInfo.SendDataList
      Dim objSendDataInfo As New MSG_T1F1M1_SendMessage.BodyInfo.SendDataInfo


      'gMain.objHandling.

      objT1F1M1_SendMessage.Body.MessageList = objMessageList
      'objT1F1M1_SendMessage.Body.SendList = Nothing

      'str_FilePath

      'Send_Maill(objT1F1M1_SendMessage, enuMessageType.Account)

      '更新狀態


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function I_Auto_Report_Inbound_Lack() As Boolean
    Try
      SendMessageToLog("I_Auto_Report_Inbound_Lack", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Dim newTime = GetNewTime_DBFormat()
      Dim bln_sendFile As Boolean = True '是否發送檔案
      Dim str_FilePath As String = "" '要發送的檔案位置
      Dim F_Date = Date.Parse(newTime).AddDays(-1).ToString(DBTimeFormat)
      Dim T_Date = newTime
      '1. 產生檔案
      Dim csv_ExportPath = String.Format("<CSVOption>
	<FilePath>C:\Log\Export</FilePath>
	<SQList>
		<SQLInfo>
			<FileName>自動倉入庫短缺資訊</FileName>
			<SQL>select b.PO_ID 單號,b.PO_SERIAL_NO 項次,b.sku_no as 料號, b.Lot_no as 料批, TRIM(to_char(b.QTY,'999,999,999,990'))as 計劃數量,TRIM(to_char(c.QTY_FINISH,'999,999,999,990'))as 完成數量, TRIM(to_char((b.qty-c.QTY_FINISH),'999,999,999,990')) as 差異數量 ,'' as 結單時間 from WMS_T_PO a inner join WMS_T_PO_DTL b on a.PO_ID = b.PO_ID inner join WMS_T_PO_MERGE c on c.po_id = a.PO_id and c.PO_SERIAL_NO=b.PO_SERIAL_NO where a.WO_TYPE='1' and (b.qty-c.QTY_FINISH)>0 and c.QTY_FINISH>0 and b.PODTL_STATUS not in (3) UNION select c.PO_ID,c.PO_SERIAL_NO,c.sku_no as 料號, c.Lot_no as 料批,TRIM(to_char(c.QTY,'999,999,999,990')) as 計劃數量, TRIM(to_char(c.QTY_PROCESS,'999,999,999,990')) as 完成數量,TRIM(to_char((c.qty-c.QTY_PROCESS),'999,999,999,990')) as 差異數量 , c.CLOSE_time as 結單時間 from wms_H_WO_HIST c where c.CLOSE_time BETWEEN '{0}' and '{1}' and c.WO_TYPE='1' and (c.qty-c.QTY_PROCESS)>0 and c.QTY_PROCESS>0 UNION select b.PO_ID 單號,b.PO_SERIAL_NO 項次,b.sku_no as 料號, b.Lot_no as 料批, TRIM(to_char(b.QTY,'999,999,999,990'))as 計劃數量,TRIM(to_char(b.QTY_FINISH,'999,999,999,990'))as 完成數量, TRIM(to_char((b.qty-b.QTY_FINISH),'999,999,999,990')) as 差異數量 ,c.CLOSE_TIME as 結單時間 from WMS_T_PO a inner join WMS_T_PO_DTL b on a.PO_ID = b.PO_ID inner join WMS_H_WO_HIST c on b.PO_ID=C.HIST_COMMON1 and b.Lot_No=C.Lot_No where a.WO_TYPE='1' and (b.qty-b.QTY_FINISH)>0 and b.QTY_FINISH>0 and b.PODTL_STATUS=3 and c.CLOSE_TIME BETWEEN '{0}' and '{1}'</SQL>
			<DateTimeRange>
				<FieldName/>
				<HourPeriod>24</HourPeriod>
			</DateTimeRange>
		</SQLInfo>
	</SQList>
</CSVOption>", F_Date, T_Date)

      Dim objCsvInfo As CSVOption = Nothing
      If I_Load_EXPORT(csv_ExportPath, objCsvInfo) = True Then
        For Each info As CSVOption.SQLInfo In objCsvInfo.SQList
          Dim sdate = $"{Now.AddHours(-info.DateTimeRange.HourPeriod):yyyy/MM/dd HH:mm:ss}"
          Dim edate = $"{Now:yyyy/MM/dd HH:mm:ss}"
          ' get data
          Dim dt As DataTable = GetCSVData(info.SQL, info.DateTimeRange.FieldName, sdate, edate)
          If dt.Columns.Count = 0 Then '如果0行 則不發送檔案
            bln_sendFile = False
          End If
          ' create csv file
          CreateCSVFile(dt, info, objCsvInfo, str_FilePath)
        Next

        SendMessageToLog("Export completed.", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)


        '2. 發送檔案
        Dim objT1F1M1_SendMessage As New MSG_T1F1M1_SendMessage

        '發送的訊息
        Dim objMessageList As New MSG_T1F1M1_SendMessage.BodyInfo.MessageDataList
        Dim objMessageInfo As New MSG_T1F1M1_SendMessage.BodyInfo.MessageDataInfo
        objMessageInfo.MESSAGE_TITLE = GetNewTime_ByDataTimeFormat(DBDayFormat) & " 自動倉入庫短缺資訊"
        If bln_sendFile = True Then
          objMessageInfo.MESSAGE_TEXT = "Dear Sir,
附件為 " & F_Date & " ~ " & T_Date & " 自動倉入庫短缺資訊。"
        Else
          objMessageInfo.MESSAGE_TEXT = "Dear Sir,
" & F_Date & " ~ " & T_Date & " 無入庫短缺資訊。"
        End If
        objMessageInfo.ATTACHMENT_PATH = IIf(bln_sendFile, str_FilePath, "")
        objMessageList.MessageInfo.Add(objMessageInfo) '僅一組 寫死

        '發送的人員
        Dim objSendDataList As New MSG_T1F1M1_SendMessage.BodyInfo.SendDataList
        Dim objSendDataInfo As New MSG_T1F1M1_SendMessage.BodyInfo.SendDataInfo


        'gMain.objHandling.

        objT1F1M1_SendMessage.Body.MessageList = objMessageList
        'objT1F1M1_SendMessage.Body.SendList = Nothing

        'str_FilePath

        'Send_Maill(objT1F1M1_SendMessage, enuMessageType.InboundN)


      End If


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Sub I_Auto_SetProductionInfo()
    SyncLock gMain.objHandling.objLineProduction_InfoLock
      SyncLock gMain.objHandling.objProduction_InfoLock
        Try
          '先排序取出所有可以沖銷的單據，之後再進行Line的數量調整
          Dim strLog As String = ""
          strLog = "執行自動分配生產線生產數量到工單"
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Dim Now_Time As String = GetNewTime_DBFormat()
          Dim dicUpdateLineProductionInfo As New Dictionary(Of String, clsLineProduction_Info)
          Dim dicUpdateProductionInfo As New Dictionary(Of String, clsProduce_Info)
          Dim tmp_dicProductionInfo As New Dictionary(Of String, clsProduce_Info)
          If O_Get_dicProductionInfoByStutusNotEnd_Sort(gMain.objHandling.gdicProduce_Info, tmp_dicProductionInfo) = True Then
            Dim tmp_dciCheckCmpKey As New Dictionary(Of String, String)  '記錄檢查過的Factory_No+Area_No
            For Each objProductionInfo As clsProduce_Info In tmp_dicProductionInfo.Values
              Dim Factory_No As String = objProductionInfo.Factory_No
              Dim Area_No As String = objProductionInfo.Area_No
              Dim Qty As Double = objProductionInfo.Qty '需求數量
              Dim Qty_Process As Double = objProductionInfo.Qty_Process ' 己生產數量
              Dim Qty_NG As Double = objProductionInfo.Qty_NG 'NG的數量
              Dim Status As enuProduceStatus = objProductionInfo.Status
              Dim CheckCmpKey As String = Factory_No & LinkKey & Area_No
              If tmp_dciCheckCmpKey.ContainsKey(CheckCmpKey) = False Then '沒有在tmp_dciCheckCmpKey裡的才進行確認
                '取得這個Facotry_No、Area_No內的Line
                Dim tmp_dicLineProductionInfo As New Dictionary(Of String, clsLineProduction_Info)
                If O_Get_dicCLineProductionInfoByFacotryNo_AreaNo(gMain.objHandling.gdicLineProduction_Info, Factory_No, Area_No, tmp_dicLineProductionInfo) = True Then
                  Dim ProduciotnCheckCmp As Boolean = False
                  Dim ProcuctionChenge As Boolean = False
                  For Each objLineProductionInfo As clsLineProduction_Info In tmp_dicLineProductionInfo.Values
                    Dim Tmp_objLineProductionInfo As clsLineProduction_Info = Nothing
                    If dicUpdateLineProductionInfo.TryGetValue(objLineProductionInfo.gid, Tmp_objLineProductionInfo) = False Then
                      Tmp_objLineProductionInfo = objLineProductionInfo.Clone()
                    End If
                    Dim LineQty_Process As Double = Tmp_objLineProductionInfo.Qty_Process
                    Dim LinePreviousQty_Process As Double = Tmp_objLineProductionInfo.Previous_Qty_Process
                    Dim LineResetQty_Process As Double = Tmp_objLineProductionInfo.Reset_Qty_Process
                    Dim LineQty_Modify As Double = Tmp_objLineProductionInfo.Qty_Modify
                    Dim LinePreviousQty_Modify As Double = Tmp_objLineProductionInfo.Previous_Qty_Modify
                    Dim LineResetQty_Modify As Double = Tmp_objLineProductionInfo.Reset_Qty_Modify
                    Dim LineQty_NG As Double = Tmp_objLineProductionInfo.Qty_NG
                    Dim LinePreviousQty_NG As Double = Tmp_objLineProductionInfo.Previous_Qty_NG
                    Dim LineResetQty_NG As Double = Tmp_objLineProductionInfo.Reset_Qty_NG
                    Dim NewQtyProcess As Double = LineQty_Process - LinePreviousQty_Process + LineResetQty_Process '這一次新增Process的數量
                    Dim NewQtyModify As Double = LineQty_Modify - LinePreviousQty_Modify + LineResetQty_Modify '這一次新增Modify的數量
                    Dim NewQtyNG As Double = LineQty_NG - LinePreviousQty_NG + LineResetQty_NG  '這一次新增的NG數量
                    If NewQtyProcess <> 0 OrElse NewQtyModify <> 0 OrElse NewQtyNG <> 0 Then
                      If (Qty - Qty_Process) > 0 Then
                        If (Qty - Qty_Process) >= (NewQtyProcess + NewQtyModify) Then '機台增加的數量可以完全配到單據裡
                          LinePreviousQty_Process = LineQty_Process
                          LineResetQty_Process = 0
                          LinePreviousQty_Modify = LineQty_Modify
                          LineResetQty_Modify = 0
                          Qty_Process = Qty_Process + NewQtyProcess + NewQtyModify
                        Else  '機台增加的數量超過一張單據的需求數量
                          Dim Need_Qty_Process As Double = Qty - Qty_Process
                          '優先配到Modify的數量，因為Modify的數量有可能是負的，優先配會好處理
                          If Need_Qty_Process >= NewQtyModify Then
                            LinePreviousQty_Modify = LineQty_Modify
                            LineResetQty_Modify = 0
                            Qty_Process = Qty_Process + NewQtyModify
                            Need_Qty_Process = Need_Qty_Process - NewQtyModify
                          Else
                            If Need_Qty_Process >= LineResetQty_Modify Then
                              Qty_Process = Qty_Process + LineResetQty_Modify
                              Need_Qty_Process = Need_Qty_Process - LineResetQty_Modify
                              LineResetQty_Modify = 0
                            Else
                              Qty_Process = Qty
                              LineResetQty_Modify = LineResetQty_Modify - Need_Qty_Process
                            End If
                            Dim LineCreateModify As Double = LineQty_Modify - LinePreviousQty_Modify
                            If Need_Qty_Process >= LineCreateModify Then
                              Qty_Process = Qty_Process + LineCreateModify
                              Need_Qty_Process = Need_Qty_Process - LineCreateModify
                              LinePreviousQty_Modify = LineQty_Modify
                            Else
                              Qty_Process = Qty
                              Need_Qty_Process = 0
                              LinePreviousQty_Modify = LinePreviousQty_Modify + LineCreateModify
                            End If
                          End If
                          '分配Process的數量
                          If Need_Qty_Process >= NewQtyProcess Then
                            LinePreviousQty_Process = LineQty_Process
                            LineResetQty_Process = 0
                            Qty_Process = Qty_Process + NewQtyProcess
                            Need_Qty_Process = Need_Qty_Process - NewQtyProcess
                          Else
                            If Need_Qty_Process >= LineResetQty_Process Then
                              Qty_Process = Qty_Process + LineResetQty_Process
                              Need_Qty_Process = Need_Qty_Process - LineResetQty_Process
                              LineResetQty_Process = 0
                            Else
                              Qty_Process = Qty
                              LineResetQty_Process = LineResetQty_Process - Need_Qty_Process
                            End If
                            Dim LineCreateProcess As Double = LineQty_Process - LinePreviousQty_Process
                            If Need_Qty_Process >= LineCreateProcess Then
                              Qty_Process = Qty_Process + LineCreateProcess
                              Need_Qty_Process = Need_Qty_Process - LineCreateProcess
                              LinePreviousQty_Process = LineQty_Process
                            Else
                              Qty_Process = Qty
                              Need_Qty_Process = 0
                              LinePreviousQty_Process = LinePreviousQty_Process + LineCreateProcess
                            End If
                          End If
                          ProduciotnCheckCmp = True
                        End If
                        Qty_NG = Qty_NG + NewQtyNG
                        LinePreviousQty_NG = LineQty_NG
                        LineResetQty_NG = 0
                        '把資料回寫到Tmp_objLineProductionInfo
                        Tmp_objLineProductionInfo.Previous_Qty_Process = LinePreviousQty_Process
                        Tmp_objLineProductionInfo.Reset_Qty_Process = LineResetQty_Process
                        Tmp_objLineProductionInfo.Previous_Qty_Modify = LinePreviousQty_Modify
                        Tmp_objLineProductionInfo.Reset_Qty_Modify = LineResetQty_Modify
                        Tmp_objLineProductionInfo.Previous_Qty_NG = LinePreviousQty_NG
                        Tmp_objLineProductionInfo.Reset_Qty_NG = LineResetQty_NG
                      Else
                        ProduciotnCheckCmp = True
                      End If
                      If dicUpdateLineProductionInfo.ContainsKey(Tmp_objLineProductionInfo.gid) = False Then
                        dicUpdateLineProductionInfo.Add(Tmp_objLineProductionInfo.gid, Tmp_objLineProductionInfo)
                      End If
                      ProcuctionChenge = True
                    End If
                    If ProduciotnCheckCmp = True Then
                      Exit For
                    End If
                  Next
                  If ProcuctionChenge = True Then
                    '把修改的資料寫入dicUpdateProductionInfo
                    If dicUpdateProductionInfo.ContainsKey(objProductionInfo.gid) = False Then
                      Dim objNewProductionInfo = objProductionInfo.Clone()
                      Dim tmp_dicLineArea As New Dictionary(Of String, clsLine_Area)
                      If O_Get_dicLastCLineAreaByFacotryNo_AreaNo(gMain.objHandling.gdicLine_Area, Factory_No, Area_No, tmp_dicLineArea) = True Then
                        objNewProductionInfo.Previous_Area_No = tmp_dicLineArea.First.Value.Area_No  '-第一個生產線  上一個是庫別 會填空的				
                      End If

                      objNewProductionInfo.Qty_Process = Qty_Process
                      objNewProductionInfo.Qty_NG = Qty_NG
                      objNewProductionInfo.Update_Time = Now_Time
                      If objNewProductionInfo.Qty_Process >= Qty Then
                        objNewProductionInfo.Status = enuProduceStatus.NormalEnd
                        If objNewProductionInfo.Start_Time = "" Then
                          objNewProductionInfo.Start_Time = Now_Time
                        End If
                        If objNewProductionInfo.Finish_Time = "" Then
                          objNewProductionInfo.Finish_Time = Now_Time
                        End If
                      ElseIf objNewProductionInfo.Qty_Process > 0 Then
                        objNewProductionInfo.Status = enuProduceStatus.Process
                        If objNewProductionInfo.Start_Time = "" Then
                          objNewProductionInfo.Start_Time = Now_Time
                        End If
                      End If
                      dicUpdateProductionInfo.Add(objNewProductionInfo.gid, objNewProductionInfo)
                    End If
                  End If
                  If ProduciotnCheckCmp = False Then  '表示所有的LineProduction都處理完了，但是Production的數量還沒有滿足(此Area不再進行確認)
                    tmp_dciCheckCmpKey.Add(CheckCmpKey, CheckCmpKey)
                  End If
                Else '找不到表示有該Area的資料(此Area不再進行確認)
                  tmp_dciCheckCmpKey.Add(CheckCmpKey, CheckCmpKey)
                End If
              End If
            Next
            '把更新的資料寫入資料庫
            Dim lstSQL As New List(Of String)
            '取得SQL
            For Each obj As clsLineProduction_Info In dicUpdateLineProductionInfo.Values
              obj.O_Add_Update_SQLString(lstSQL)
            Next
            For Each obj As clsProduce_Info In dicUpdateProductionInfo.Values
              obj.O_Add_Update_SQLString(lstSQL)
            Next
            '   
            If Common_DBManagement.BatchUpdate(lstSQL) = False Then
              '更新DB失敗則回傳False
              'ret_ResultMsg = "WMS Update DB Failed"
              strLog = "WMS 更新资料库失败"
              SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Else
              '更新記憶體資料
              '1.更新
              For Each objNew As clsLineProduction_Info In dicUpdateLineProductionInfo.Values
                Dim obj As clsLineProduction_Info = Nothing
                If gMain.objHandling.gdicLineProduction_Info.TryGetValue(objNew.gid, obj) = True Then
                  obj.Update_To_Memory(objNew)
                End If
              Next
              '2.更新
              For Each objNew As clsProduce_Info In dicUpdateProductionInfo.Values
                Dim obj As clsProduce_Info = Nothing
                If gMain.objHandling.gdicProduce_Info.TryGetValue(objNew.gid, obj) = True Then
                  obj.Update_To_Memory(objNew)
                End If
              Next

            End If
          End If
        Catch ex As Exception
          SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End Try
      End SyncLock
    End SyncLock
  End Sub

  '自動設定MaintenanceStatus
  Private Sub I_Auto_SetMaintenanceStatus()
    Try
      Dim strLog As String = ""
      strLog = "執行MaintenanceStatus更新"
      SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Dim Maintenance_ID As String = "Qty_Maintenance" '暫時先這樣寫
      Dim Function_ID As String = "Qty_Process" '暫時先這樣寫
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim dicUpdateMaterialStatus As New Dictionary(Of String, clsMAINTENANCE_STATUS)
      For Each objLineProduction As clsLineProduction_Info In gMain.objHandling.gdicLineProduction_Info.Values
        With objLineProduction
          Dim objMaintenanceStatus As clsMAINTENANCE_STATUS = Nothing
          If gMain.objHandling.O_Get_MaintenanceStatus(.Factory_No, .Device_No, .Area_No, .Unit_ID, Maintenance_ID, Function_ID, objMaintenanceStatus) = True Then
            If IsNumeric(objMaintenanceStatus.VALUE) AndAlso CDbl(objMaintenanceStatus.VALUE) = .Qty_Process Then
              Continue For
            End If
            Dim objNewMaterialStatus As clsMAINTENANCE_STATUS = objMaintenanceStatus.Clone()
            objNewMaterialStatus.VALUE = objLineProduction.Qty_Process
            objNewMaterialStatus.UPDATE_TIME = Now_Time
            If dicUpdateMaterialStatus.ContainsKey(objNewMaterialStatus.gid) = False Then
              dicUpdateMaterialStatus.Add(objNewMaterialStatus.gid, objNewMaterialStatus)
            End If
          End If
        End With
      Next

      '寫入DB並更新記憶體
      If dicUpdateMaterialStatus.Any = True Then
        Dim lstSQL As New List(Of String)
        For Each objMaterialStatus As clsMAINTENANCE_STATUS In dicUpdateMaterialStatus.Values
          objMaterialStatus.O_Add_Update_SQLString(lstSQL)
        Next
        If Common_DBManagement.BatchUpdate(lstSQL) = False Then
          '更新DB失敗則回傳False
          'ret_ResultMsg = "WMS Update DB Failed"
          strLog = "WMS 更新资料库失败"
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Else
          '更新記憶體資料
          '1.更新
          For Each objNew As clsMAINTENANCE_STATUS In dicUpdateMaterialStatus.Values
            Dim obj As clsMAINTENANCE_STATUS = Nothing
            If gMain.objHandling.gdicMaintenance_Status.TryGetValue(objNew.gid, obj) = True Then
              obj.Update_To_Memory(objNew)
            End If
          Next
        End If
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '檢查是否需要發送保養異常
  Private Sub I_Auto_SetLineInfo()
    Try
      Dim strLog As String = ""
      strLog = "執行MaintenanceStatus檢查"
      SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      '取得所有Maintenance和MaintenanceStatus
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim dicUpdateMaintenanceStatus As New Dictionary(Of String, clsMAINTENANCE_STATUS)
      Dim dicAddLineInfo As New Dictionary(Of String, clsLineInfo)
      For Each objMaintenanceStatus As clsMAINTENANCE_STATUS In gMain.objHandling.gdicMaintenance_Status.Values
        Dim Factory_No As String = objMaintenanceStatus.FACTORY_NO
        Dim Device_No As String = objMaintenanceStatus.DEVICE_NO
        Dim Area_No As String = objMaintenanceStatus.AREA_NO
        Dim Unit_ID As String = objMaintenanceStatus.UNIT_ID
        Dim Maintenance_ID As String = objMaintenanceStatus.MAINTENANCE_ID
        Dim Function_ID As String = objMaintenanceStatus.FUNCTION_ID
        Dim Value As String = objMaintenanceStatus.VALUE
        Dim Maintenance_Set As Boolean = objMaintenanceStatus.MAINTENANCE_SET
        Dim NeedSetMaintenance As Boolean = False '檢查後是否需要進行保養
        If Maintenance_Set = False Then '沒有過的才判斷
          '判斷保養警告是否啟用
          Dim objMaintenance As clsMAINTENANCE = Nothing
          If gMain.objHandling.O_Get_Maintenance(Factory_No, Device_No, Area_No, Unit_ID, Maintenance_ID, objMaintenance) = True Then
            If objMaintenance.ENABLE = True Then
              '判斷是否超標
              Dim objMaintenanceDTL As clsMAINTENANCE_DTL = Nothing
              If gMain.objHandling.O_Get_MaintenanceDTL(Factory_No, Device_No, Area_No, Unit_ID, Maintenance_ID, Function_ID, objMaintenanceDTL) = True Then
                '根據Maintenance設定的是數值或是日期進行對應的檢查
                Dim MaintenanceValueType As enuMaintenanceValueType = objMaintenanceDTL.VALUE_TYPE
                Dim MaintenanceNoticeType As enuMaintenanceNoticeType = objMaintenanceDTL.NOTICE_TYPE
                Dim HighValue As String = objMaintenanceDTL.HIGH_WATER_VALUE
                Dim LowValue As String = objMaintenanceDTL.LOW_WATER_VALUE
                Select Case MaintenanceValueType
                  Case enuMaintenanceValueType.IsNumber '數字的檢查
                    '先檢查Value是否為數字，如果不為數字則不進行判斷
                    If IsNumeric(Value) = True Then
                      Dim NewValue As Double = CDbl(Value)
                      '數字根據是高標、低標、高低標進行判斷
                      Select Case MaintenanceNoticeType
                        Case enuMaintenanceNoticeType.HighLowCheck
                          '高於高標
                          If IsNumeric(HighValue) AndAlso CDbl(HighValue) < NewValue Then
                            NeedSetMaintenance = True
                          End If
                          '低於低標
                          If IsNumeric(LowValue) AndAlso CDbl(LowValue) > NewValue Then
                            NeedSetMaintenance = True
                          End If
                        Case enuMaintenanceNoticeType.HighCheck
                          '高於高標
                          If IsNumeric(HighValue) AndAlso CDbl(HighValue) < NewValue Then
                            NeedSetMaintenance = True
                          End If
                        Case enuMaintenanceNoticeType.LowCheck
                          '低於低標
                          If IsNumeric(LowValue) AndAlso CDbl(LowValue) > NewValue Then
                            NeedSetMaintenance = True
                          End If
                      End Select
                    End If
                  Case enuMaintenanceValueType.IsDate '日期的檢查
                    '先檢查傳入的值是否為標準的日期格式
                    Dim NewValueDate As String = ParseTime(Value, DBTimeFormat)
                    If NewValueDate <> "" Then
                      '日期根據是高標、低標、高低標進行判斷
                      Select Case MaintenanceNoticeType
                        Case enuMaintenanceNoticeType.HighLowCheck
                          '高於高標
                          If IsNumeric(HighValue) AndAlso SubTractTime_Day(Now_Time, NewValueDate) > HighValue Then
                            NeedSetMaintenance = True
                          End If
                          '低於低標
                          If IsNumeric(LowValue) AndAlso SubTractTime_Day(Now_Time, NewValueDate) < LowValue Then
                            NeedSetMaintenance = True
                          End If
                        Case enuMaintenanceNoticeType.HighCheck
                          '高於高標
                          If IsNumeric(HighValue) AndAlso SubTractTime_Day(Now_Time, NewValueDate) > HighValue Then
                            NeedSetMaintenance = True
                          End If
                        Case enuMaintenanceNoticeType.LowCheck
                          '低於低標
                          If IsNumeric(LowValue) AndAlso SubTractTime_Day(Now_Time, NewValueDate) < LowValue Then
                            NeedSetMaintenance = True
                          End If
                      End Select
                    End If
                End Select
              End If
              '如果需要進行保養警告，則寫入保養警告的Table
              If NeedSetMaintenance = True Then
                Dim objNewMaintenanceStatus As clsMAINTENANCE_STATUS = objMaintenanceStatus.Clone()
                objNewMaintenanceStatus.MAINTENANCE_SET = True
                objNewMaintenanceStatus.MAINTENANCE_TIME = Now_Time
                If dicUpdateMaintenanceStatus.ContainsKey(objNewMaintenanceStatus.gid) = False Then
                  dicUpdateMaintenanceStatus.Add(objNewMaintenanceStatus.gid, objNewMaintenanceStatus)
                  '寫入保養警告的Table
                  Dim objLineInfo As New clsLineInfo(Factory_No, Area_No, Device_No, Unit_ID, Now_Time, objMaintenanceDTL.MAINTENANCE_MESSAGE, Maintenance_ID, Function_ID)
                  If gMain.objHandling.gdicLineInfo.ContainsKey(objLineInfo.gid) = False AndAlso dicAddLineInfo.ContainsKey(objLineInfo.gid) = False Then
                    dicAddLineInfo.Add(objLineInfo.gid, objLineInfo)
                  End If
                End If
              End If
            End If
          End If
        End If
      Next
      '當有產生需要保養資訊時，寫入DB
      If dicUpdateMaintenanceStatus.Any = True OrElse dicAddLineInfo.Any = True Then
        Dim lstSQL As New List(Of String)
        For Each objMaterialStatus As clsMAINTENANCE_STATUS In dicUpdateMaintenanceStatus.Values
          objMaterialStatus.O_Add_Update_SQLString(lstSQL)
        Next
        For Each objLineInfo As clsLineInfo In dicAddLineInfo.Values
          objLineInfo.O_Add_Insert_SQLString(lstSQL)
        Next
        If Common_DBManagement.BatchUpdate(lstSQL) = False Then
          '更新DB失敗則回傳False
          'ret_ResultMsg = "WMS Update DB Failed"
          strLog = "WMS 更新资料库失败"
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Else
          '更新記憶體資料
          '1.更新
          For Each objNew As clsMAINTENANCE_STATUS In dicUpdateMaintenanceStatus.Values
            Dim obj As clsMAINTENANCE_STATUS = Nothing
            If gMain.objHandling.gdicMaintenance_Status.TryGetValue(objNew.gid, obj) = True Then
              obj.Update_To_Memory(objNew)
            End If
          Next
          '1.更新
          For Each objNew As clsLineInfo In dicAddLineInfo.Values
            objNew.Add_Relationship(gMain.objHandling)
          Next
        End If
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '進行班次生產數量的Reset
  Public Sub I_Auto_ResetLineProductionByClass()
    Try
      Dim strLog As String = ""
      strLog = "執行ResetLineProduction的檢查"
      SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Dim lstSQL As New List(Of String)
      Dim Now_date As String = GetNewTime_ByDataTimeFormat(DBDayFormat)
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim dicUpdateSystemStatus As New Dictionary(Of String, clsSystemStatus)
      Dim lstAdd_ClassProductionHist As New List(Of clsClassProduction_HIST)
      '取得當下的時間是屬於哪一個Class
      Dim Now_Class_No As String = O_Get_Now_Class()
      If Now_Class_No <> "" Then
        '如果NowClass_No不是空的，檢查取得NowClass_No和目前的Class_No是否相同
        Dim CurrentClass_No As String = ""
        Dim objSystemStatus As clsSystemStatus = Nothing
        If gMain.objHandling.O_Get_SystemStatus(enuSystemStatus.CurrentClassNo, objSystemStatus) Then
          CurrentClass_No = objSystemStatus.STATUS_VALUE
          '如果NowClass_No和目前的Class_No不相同，則進行Reset的處理
          If CurrentClass_No <> Now_Class_No Then
            SendMessageToLog("當班資訊更新、機台數量Reset", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            Dim objCurrentClassNoSystemStatus = objSystemStatus.Clone()
            '取得前一次的PreviousClassNo
            If gMain.objHandling.O_Get_SystemStatus(enuSystemStatus.PreviousClassNo, objSystemStatus) Then
              Dim objPreviousClassNoSystemStatus = objSystemStatus.Clone()
              '把前一班的Class_No改成當班的Class_No
              objPreviousClassNoSystemStatus.STATUS_VALUE = CurrentClass_No
              objPreviousClassNoSystemStatus.UPDATE_TIME = Now_Time
              If dicUpdateSystemStatus.ContainsKey(objPreviousClassNoSystemStatus.gid) = False Then
                dicUpdateSystemStatus.Add(objPreviousClassNoSystemStatus.gid, objPreviousClassNoSystemStatus)
              End If
              '把當班的Class_No改成目前取得的Class_No
              objCurrentClassNoSystemStatus.STATUS_VALUE = Now_Class_No
              objCurrentClassNoSystemStatus.UPDATE_TIME = Now_Time
              If dicUpdateSystemStatus.ContainsKey(objCurrentClassNoSystemStatus.gid) = False Then
                dicUpdateSystemStatus.Add(objCurrentClassNoSystemStatus.gid, objCurrentClassNoSystemStatus)
              End If

              '當前班次
              'CurrentClass_No
              '取得當班時間區間(跨日要注意!)
              Dim objCurrentClass As clsClass = Nothing
              If gMain.objHandling.O_Get_Class(CurrentClass_No, objCurrentClass) = False Then
                SendMessageToLog("無法取得當前班別資料，班別：" & CurrentClass_No, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return
              End If
              '撈WMS_CH_LINE_PRODUCTION_HIST 統整時間區間內的數量
              Dim lstLineProductionHIST As New List(Of clsLineProduction_Hist)
              '取得日期
              Dim StartTime As String = ""
              Dim EndTime As String = ""
              '跨日
              If objCurrentClass.CLASS_START_TIME > objCurrentClass.CLASS_END_TIME Then
                StartTime = DateTime.Parse(Now_date).AddDays(-1).ToString(DBTimeFormat) & " " & objCurrentClass.CLASS_START_TIME
                EndTime = Now_date & " " & objCurrentClass.CLASS_END_TIME
              Else '不跨日
                StartTime = Now_date & " " & objCurrentClass.CLASS_START_TIME
                EndTime = Now_date & " " & objCurrentClass.CLASS_END_TIME
              End If
              If gMain.objHandling.O_GetDB_Line_Production_HIST(StartTime, EndTime, lstLineProductionHIST) = False Then
                SendMessageToLog("無法取得線別生產歷史，開始時間：" & StartTime & " 結束時間：" & EndTime, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return
              End If
              '組成dic
              Dim dicLPHIST As New Dictionary(Of String, List(Of clsLineProduction_Hist))
              For Each objLPHIST In lstLineProductionHIST
                Dim _Factory_No As String = objLPHIST.Factory_No
                Dim _Area_No As String = objLPHIST.Area_No
                Dim _Device_No As String = objLPHIST.Device_No
                Dim _Unit_ID As String = objLPHIST.Unit_ID
                Dim Key = _Factory_No & "_" & _Area_No & "_" & _Device_No & "_" & _Unit_ID
                If dicLPHIST.ContainsKey(Key) = False Then
                  dicLPHIST.Add(Key, New List(Of clsLineProduction_Hist)({objLPHIST}))
                Else
                  dicLPHIST.Item(Key).Add(objLPHIST)
                End If
              Next
              For Each objLPHIST In dicLPHIST
                '更新Production的資訊
                Dim Factory_No As String = objLPHIST.Key.Split("_")(0)
                Dim Area_No As String = objLPHIST.Key.Split("_")(1)
                Dim Device_No As String = objLPHIST.Key.Split("_")(2)
                Dim Unit_ID As String = objLPHIST.Key.Split("_")(3)
                Dim Last_Qty_Total As Double = 0
                Dim _Hist_Time As String = ""
                Dim sum_qty_process As Double = 0
                Dim sum_qty_modify As Double = 0
                Dim sum_qty_ng As Double = 0

                For Each _LPHIST In objLPHIST.Value
                  Dim Last_Qty_Process As Double = CDbl(_LPHIST.Qty_Process)
                  Dim Last_Qty_Modify As Double = CDbl(_LPHIST.Qty_Modify)
                  Dim Last_Qty_NG As Double = CDbl(_LPHIST.Qty_NG)
                  If _Hist_Time = "" Then
                    _Hist_Time = _LPHIST.Hist_Time
                  ElseIf _Hist_Time < _LPHIST.Hist_Time Then '抓最後一個時間
                    _Hist_Time = _LPHIST.Hist_Time
                    Last_Qty_Total = CDbl(_LPHIST.QTY_TOTAL)
                  End If
                  sum_qty_process += Last_Qty_Process
                  sum_qty_modify += Last_Qty_Modify
                  sum_qty_ng += Last_Qty_NG
                Next
                '新增班別生產數量的歷史記錄
                Dim objClassProductionHist As New clsClassProduction_HIST(Factory_No, Area_No, Device_No, Unit_ID, CurrentClass_No, Last_Qty_Total, sum_qty_process, sum_qty_modify, sum_qty_ng, Now_Time)
                lstAdd_ClassProductionHist.Add(objClassProductionHist)
              Next
              '取得SQL
              For Each lst_CPHIST In lstAdd_ClassProductionHist
                lst_CPHIST.O_Add_Insert_SQLString(lstSQL)
              Next
              For Each obj As clsSystemStatus In dicUpdateSystemStatus.Values
                obj.O_Add_Update_SQLString(lstSQL)
              Next
              If Common_DBManagement.BatchUpdate(lstSQL) = False Then
                SendMessageToLog("I_Auto_ResetLineProductionByClass Update DB failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If



              ''不下命令給WCS了，人員會自己reset。程序去加總數量就可以了 20190312
              ''取得流水號
              'Dim dicUUID As New Dictionary(Of String, clsUUID)
              'If gMain.objHandling.O_GetDB_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
              '  strLog = "Get UUID False"
              '  SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              '  Exit Sub
              'End If
              'If dicUUID.Any = False Then
              '  strLog = "Get UUID False"
              '  SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              '  Exit Sub
              'End If
              'Dim objUUID = dicUUID.Values(0)
              'Dim lstMCSSql As New List(Of String)
              ''下Reset的Message給MCS
              'For Each objLineProductionInfo As clsLineProduction_Info In gMain.objHandling.gdicLineProduction_Info.Values
              '  Dim lstUnit = New eCA_TransactionMessage.MSG_T3F5S1_LineProductionResetRequest.clsBody.clsUnitList
              '  Dim Factory_No As String = objLineProductionInfo.Factory_No
              '  Dim Area_No As String = objLineProductionInfo.Area_No
              '  Dim Device_No As String = objLineProductionInfo.Device_No
              '  Dim Unit_ID As String = objLineProductionInfo.Unit_ID
              '  Dim objUnitInfo As New eCA_TransactionMessage.MSG_T3F5S1_LineProductionResetRequest.clsBody.clsUnitList.clsUnitInfo
              '  objUnitInfo.FACTORY_NO = Factory_No
              '  objUnitInfo.AREA_NO = Area_No
              '  objUnitInfo.DEVICE_NO = Device_No
              '  objUnitInfo.UNIT_ID = Unit_ID
              '  lstUnit.UnitInfo.Add(objUnitInfo)
              '  '組成Message發送給MCS
              '  Dim UUID = objUUID.Get_NewUUID
              '  Dim objT3F1S1 As New eCA_TransactionMessage.MSG_T3F5S1_LineProductionResetRequest
              '  objT3F1S1.Header = New eCA_TransactionMessage.clsHeader
              '  objT3F1S1.Header.UUID = UUID
              '  objT3F1S1.Header.EventID = "T3F5S1_LineProductionResetRequest"
              '  objT3F1S1.Header.Direction = "Primary"
              '  objT3F1S1.Header.ClientInfo = New eCA_TransactionMessage.clsHeader.clsClientInfo
              '  objT3F1S1.Header.ClientInfo.ClientID = "Handler"
              '  objT3F1S1.Header.ClientInfo.UserID = ""
              '  objT3F1S1.Header.ClientInfo.IP = ""
              '  objT3F1S1.Header.ClientInfo.MachineID = ""
              '  objT3F1S1.Body = New eCA_TransactionMessage.MSG_T3F5S1_LineProductionResetRequest.clsBody
              '  objT3F1S1.Body.UnitList = lstUnit
              '  '將物件轉成xml
              '  Dim strXML = ""
              '  If eCA_TransactionMessage.CombinationXmlString.PrepareMessage_T3F5S1_LineProductionResetRequest(strXML, objT3F1S1, strLog) = False Then
              '    If strLog = "" Then
              '      strLog = "轉XML錯誤(MSG_T3F5S1_LineProductionResetRequest)"
              '    End If
              '    SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              '    Exit Sub
              '  End If
              '  '寫Command 
              '  Dim Host_Command = New clsToMCSCommand(UUID, enuSystemType.HostHandler, enuSystemType.BigDataSystem, "T3F5S1_LineProductionResetRequest", 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "", "")
              '  '取得要送給WMS的CMD
              '  If Host_Command.O_Add_Insert_SQLString(lstMCSSql) = False Then
              '    strLog = "Get Insert WMS_T_MCS_COMMAND SQL Failed"
              '    SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              '    Exit Sub
              '  End If
              'Next
              ''取得SQL
              'Dim lstSql As New List(Of String)
              'For Each obj As clsSystemStatus In dicUpdateSystemStatus.Values
              '  obj.O_Add_Update_SQLString(lstSql)
              'Next
              ''發送給MCS Message
              'If MCS_T_CommandManagement.BatchUpdate(lstMCSSql) = True Then
              '  If Common_DBManagement.BatchUpdate(lstSql) = False Then
              '    strLog = "Update DB Failed, I_Auto_ResetLineProductionByClass"
              '    SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              '  Else
              '    '更新記憶體資訊
              '    For Each objNew As clsSystemStatus In dicUpdateSystemStatus.Values
              '      Dim obj As clsSystemStatus = Nothing
              '      If gMain.objHandling.gdicSystemStatus.TryGetValue(objNew.gid, obj) Then
              '        obj.Update_To_Memory(objNew)
              '      End If
              '    Next
              '  End If
              'End If
            End If
          End If
        End If
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Function O_Get_Now_Class() As String
    Try
      Dim strLog As String = ""
      Dim NowClass_No As String = ""
      Dim Check_Time As String = GetNewTime_ByDataTimeFormat(DBOnlyHMFormat)
      '取出所有Class進行判斷，進行時間的判斷，取得目前的Class
      For Each objClass As clsClass In gMain.objHandling.gdicClass.Values
        Dim Class_No As String = objClass.CLASS_NO
        Dim Start_Time As String = objClass.CLASS_START_TIME
        Dim End_Time As String = objClass.CLASS_END_TIME
        If Start_Time < End_Time Then
          If Check_Time >= Start_Time AndAlso Check_Time < End_Time Then
            NowClass_No = Class_No
            Exit For
          End If
        Else
          If Not (Check_Time < Start_Time AndAlso Check_Time > End_Time) Then
            NowClass_No = Class_No
            Exit For
          End If
        End If
      Next
      Return NowClass_No
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

#Region "匯出csv"
  Private Function I_Load_EXPORT(ByVal ExportPath As String, ByRef CsvInfo As CSVOption) As Boolean
    Dim result = False

    Try

      'If Not File.Exists(ExportPath) Then
      '  SendMessageToLog($"exports file is not exist. path={ExportPath}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return result
      'End If

      Dim xmlstr = ExportPath ' File.ReadAllText(ExportPath)
      CsvInfo = DeserializeObject(Of CSVOption)(xmlstr)

      If CsvInfo Is Nothing Then
        SendMessageToLog("read exports file fail.", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return result
      End If

      result = True
    Catch ex As Exception
      Console.WriteLine(ex.ToString())
    End Try

    Return result
  End Function
  Public Function DeserializeObject(Of T As {Class, New})(xmlstr As String) As T
    Dim result As T = Nothing

    Try

      Dim serialize As New XmlSerializer(GetType(T))

      Using reader As New StringReader(xmlstr)
        result = serialize.Deserialize(reader)
      End Using

    Catch ex As Exception
      Console.WriteLine(ex.ToString())
    End Try

    Return result
  End Function
  Private Function GetCSVData(sqlstr As String, colname As String, stdate As String, edate As String) As DataTable
    Dim ds As New DataSet
    Dim dt As DataTable = Nothing

    Try
      ' execute sql string
      Dim rtnmsg As String = String.Empty
      Dim tempstr = $"SELECT * FROM ({sqlstr})"

      If Not String.IsNullOrWhiteSpace(colname) Then
        tempstr = $"{tempstr} WHERE {colname} BETWEEN '{stdate}' AND '{edate}'"
      End If

      SendMessageToLog($"Sql Execute: {tempstr}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Common_DBManagement.DBTool.SQLExcute_DynamicConnection(tempstr, ds) ',,,,, True,,, rtnmsg)

      ' check data set exist
      If ds IsNot Nothing Then
        dt = ds.Tables(0)
      Else
        ' recrod sql syntax error
        If Not String.IsNullOrWhiteSpace(rtnmsg) Then
          SendMessageToLog($"{rtnmsg}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        SendMessageToLog($"Sql Execute Error:{tempstr}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      End If

    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try

    Return dt
  End Function
  Private Sub CreateCSVFile(ByRef dt As DataTable, info As CSVOption.SQLInfo, ByVal CsvInfo As CSVOption, ByRef str_FilePath As String)
    Try
      If dt Is Nothing Then
        Exit Sub
      End If

      ' if folder not exist, create new folder
      If Not Directory.Exists($"{CsvInfo.FilePath}\{Now:yyyyMMdd}") Then
        Directory.CreateDirectory($"{CsvInfo.FilePath}\{Now:yyyyMMdd}")
      End If

      Dim path = $"{CsvInfo.FilePath}\{Now:yyyyMMdd}\{info.FileName}_{Now:yyyyMMddHHmmss}.csv"
      str_FilePath = path
      SendMessageToLog($"File create path:{path}", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Dim fs As New FileStream(path, FileMode.OpenOrCreate)
      Dim sw = New StreamWriter(fs, Text.Encoding.GetEncoding(950)) 'EncodingNumber

      If Not String.IsNullOrWhiteSpace(info.DateTimeRange.FieldName) Then
        ' header
        sw.WriteLine("Query conditions")
        sw.WriteLine()

        Dim startDate = $"{Now.AddHours(-info.DateTimeRange.HourPeriod):yyyy/MM/dd HH:mm:ss}"
        Dim endDate = $"{Now:yyyy/MM/dd HH:mm:ss}"
        sw.WriteLine($"Time,{startDate} ~ {endDate}")
        sw.WriteLine()
      End If
      Dim lstdatacol = dt.Columns.OfType(Of DataColumn)
      Dim collist = dt.Columns.OfType(Of DataColumn).Select(Function(s) s.ColumnName).ToList()
      sw.WriteLine($"{String.Join(",", collist.Select(Function(s) $"""{s}"""))}")

      ' content
      For ri = 0 To dt.Rows.Count - 1
        Dim row = dt.Rows(ri)
        'Dim valstr As String = String.Join(",", collist.Select(Function(s) $"=""{row.Item(s)}"""))
        Dim valstr As String = String.Join(",", collist.Select(Function(s)
                                                                 Dim col = lstdatacol.FirstOrDefault(Function(f) f.ColumnName = s)
                                                                 If row.Item(s).ToString.Length > 0 Then
                                                                   If row.Item(s).ToString.Chars(0) = "0" Then
                                                                     Return $"=""{row.Item(s)}"""
                                                                   End If
                                                                 End If
                                                                 Return $"""{row.Item(s)}"""
                                                               End Function)).ToString
        sw.WriteLine(valstr)
      Next

      sw.Flush()
      sw.Close()
      fs.Close()
      dt.Dispose()
      dt = Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
#End Region


End Module
