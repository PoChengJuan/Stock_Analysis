Imports eCA_TransactionMessage
Imports eCA_HostObject
''' <summary>
''' 20181119
''' Mark
''' 狀態:Open(進行初步確認)
''' 設備上報生產線數據重置
''' </summary>
Module Module_T3F5R4_LineInProductionInfoReset
  Public Function O_Process_Message(ByVal Receive_Msg As MSG_T3F5R4_LineInProductionInfoReset,
                                    ByRef ret_strResultMsg As String) As Boolean
    SyncLock gMain.objHandling.objLineProduction_InfoLock
      Try
        Dim dicUpdateLineProductionInfo As New Dictionary(Of String, clsLineProduction_Info)
        Dim lstClassProductionHist As New List(Of clsClassProduction_HIST)
        Dim lstAddLineProductionHist As New List(Of clsLineProduction_Hist)
        Dim lstSQL As New List(Of String)
        Dim lstQueueSQL As New List(Of String)
        '檢查資料是否正確
        If I_Check_Data(Receive_Msg, ret_strResultMsg) = False Then
          Return False
        End If
        '取得要更新的資料
        If I_Get_Data(Receive_Msg, dicUpdateLineProductionInfo, lstClassProductionHist, lstAddLineProductionHist, ret_strResultMsg) = False Then
          Return False
        End If
        '取得要更新的SQL
        If I_Get_SQL(ret_strResultMsg, dicUpdateLineProductionInfo, lstClassProductionHist, lstAddLineProductionHist, lstSQL, lstQueueSQL) = False Then
          Return False
        End If
        '執行資料更新
        If I_Execute_DataUpdate(ret_strResultMsg, dicUpdateLineProductionInfo, lstSQL, lstQueueSQL) = False Then
          Return False
        End If
        Return True
      Catch ex As Exception
        ret_strResultMsg = ex.ToString
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End Try
    End SyncLock
  End Function
  '檢查傳入的資料是否正確
  Private Function I_Check_Data(ByVal Receive_Msg As MSG_T3F5R4_LineInProductionInfoReset,
                                ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim strLog As String = ""
      For Each ProductionInfo In Receive_Msg.Body.ProductionList.ProductionInfo
        Dim Factory_No As String = ProductionInfo.FACTORY_NO
        Dim Area_No As String = ProductionInfo.AREA_NO
        Dim Device_No As String = ProductionInfo.DEVICE_NO
        Dim Unit_ID As String = ProductionInfo.UNIT_ID
        Dim Qty_Total As String = ProductionInfo.QTY_TOTAL
        Dim Qty_Process As String = ProductionInfo.QTY_PROCESS
        Dim Qty_Modify As String = ProductionInfo.QTY_MODIFY
        Dim Qty_NG As String = ProductionInfo.QTY_NG
        '檢查位置是否存在
        If gMain.objHandling.O_Get_CLine(Factory_No, Area_No, Device_No, Unit_ID) = False Then
          strLog = String.Format("CLine is not Exist, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          ret_strResultMsg = String.Format("查無此生產線位置, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>;", Factory_No, Area_No, Device_No, Unit_ID)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '檢查數量上報格式是否正確
        If IsNumeric(Qty_Total) = False Then
          strLog = String.Format("Qty_Total format error, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Qty_Total = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Qty_Total)
          ret_strResultMsg += String.Format("數量格式異常, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Qty_Total = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Qty_Total)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        If IsNumeric(Qty_Process) = False Then
          strLog = String.Format("Qty_Process format error, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Qty_Process = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Qty_Process)
          ret_strResultMsg += String.Format("數量格式異常, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Qty_Process = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Qty_Process)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        If IsNumeric(Qty_Modify) = False Then
          strLog = String.Format("Qty_Modify format error, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Qty_Process = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Qty_Modify)
          ret_strResultMsg += String.Format("數量格式異常, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Qty_Process = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Qty_Modify)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        If IsNumeric(Qty_NG) = False Then
          strLog = String.Format("Qty_NG format error, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Qty_Process = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Qty_NG)
          ret_strResultMsg += String.Format("數量格式異常, Factory_No = <{0}>, Area_No = <{1}>, Device_No = <{2}>, Unit_ID = <{3}>, Qty_Process = <{4}>;", Factory_No, Area_No, Device_No, Unit_ID, Qty_NG)
          SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得資料庫新
  Private Function I_Get_Data(ByVal Receive_Msg As MSG_T3F5R4_LineInProductionInfoReset,
                              ByRef ret_dicUpdateLineProductionInfo As Dictionary(Of String, clsLineProduction_Info),
                              ByRef ret_lstClassProductionHist As List(Of clsClassProduction_HIST),
                              ByRef ret_lstLineProductionHist As List(Of clsLineProduction_Hist),
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim strLog As String = ""
      Dim Now_Time As String = GetNewTime_DBFormat()

      '取得Class_No
      Dim Class_No As String = ""
      Dim objSystemStatus As clsSystemStatus = Nothing
      If gMain.objHandling.O_Get_SystemStatus(enuSystemStatus.PreviousClassNo, objSystemStatus) Then
        Class_No = objSystemStatus.STATUS_VALUE
      End If
      For Each ProductionInfo In Receive_Msg.Body.ProductionList.ProductionInfo
        '更新Production的資訊
        Dim Factory_No As String = ProductionInfo.FACTORY_NO
        Dim Area_No As String = ProductionInfo.AREA_NO
        Dim Device_No As String = ProductionInfo.DEVICE_NO
        Dim Unit_ID As String = ProductionInfo.UNIT_ID
        Dim Last_Qty_Total As Double = CDbl(ProductionInfo.QTY_TOTAL)
        Dim Last_Qty_Process As Double = CDbl(ProductionInfo.QTY_PROCESS)
        Dim Last_Qty_Modify As Double = CDbl(ProductionInfo.QTY_MODIFY)
        Dim Last_Qty_NG As Double = CDbl(ProductionInfo.QTY_NG)
        Dim objLineProductionInfo As clsLineProduction_Info = Nothing
        If gMain.objHandling.O_Get_CLineProduction_Info(Factory_No, Area_No, Device_No, Unit_ID, objLineProductionInfo) = True Then '如果已經有資料了就進行資料更新
          '把數據全部清成0，並且把未更新到單據上的數量加到Reset_Qty上
          Dim objNewLineProductionInfo As clsLineProduction_Info = objLineProductionInfo
          Dim New_Reset_Qty_Process As Double = Last_Qty_Process - objNewLineProductionInfo.Previous_Qty_Process
          'Dim New_Reset_Qty_Modify As Double = Last_Qty_Modify - objNewLineProductionInfo.Previous_Qty_Modify
          Dim New_Reset_Qty_NG As Double = Last_Qty_NG - objNewLineProductionInfo.Previous_Qty_NG
          objNewLineProductionInfo.Reset_Qty_Process = objNewLineProductionInfo.Reset_Qty_Process + New_Reset_Qty_Process
          'objNewLineProductionInfo.Reset_Qty_Modify = objNewLineProductionInfo.Reset_Qty_Modify + Last_Qty_Modify
          objNewLineProductionInfo.Reset_Qty_NG = objNewLineProductionInfo.Reset_Qty_NG + Last_Qty_NG
          objNewLineProductionInfo.QTY_TOTAL = Last_Qty_Total
          objNewLineProductionInfo.Qty_Process = 0
          'objNewLineProductionInfo.Qty_Modify = 0
          objNewLineProductionInfo.Qty_NG = 0
          objNewLineProductionInfo.Previous_Qty_Process = 0
          'objNewLineProductionInfo.Previous_Qty_Modify = 0
          objNewLineProductionInfo.Previous_Qty_NG = 0
          objNewLineProductionInfo.Update_Time = Now_Time
          If ret_dicUpdateLineProductionInfo.ContainsKey(objNewLineProductionInfo.gid) = False Then
            ret_dicUpdateLineProductionInfo.Add(objNewLineProductionInfo.gid, objNewLineProductionInfo)
          End If
          ''Auto Excute 早八晚八自行觸發
          ''新增班別生產數量的歷史記錄
          'Dim objClassProductionHist As New clsClassProduction_HIST(Factory_No, Area_No, Device_No, Unit_ID, Class_No, Last_Qty_Total, Last_Qty_Process, Last_Qty_Modify, Last_Qty_NG, Now_Time)
          'ret_lstClassProductionHist.Add(objClassProductionHist)
          ''記錄此次上報增加了多少數量
          'Dim CurrentQty_Process As Double = Last_Qty_Process - objLineProductionInfo.Qty_Process
          'Dim CurrentQty_NG As Double = Last_Qty_NG - objLineProductionInfo.Qty_NG
          'Dim CurrentQty_Modify As Double = Last_Qty_Modify - objLineProductionInfo.Qty_Modify
          'If CurrentQty_Process > 0 OrElse CurrentQty_NG > 0 OrElse CurrentQty_Modify > 0 Then
          '  Dim objLineProductionHist As New clsLineProduction_Hist(Factory_No, Area_No, Device_No, Unit_ID, CurrentQty_Process, CurrentQty_Modify, CurrentQty_NG, Now_Time, Last_Qty_Total)
          '  ret_lstLineProductionHist.Add(objLineProductionHist)
          'End If
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得更新的SQL
  Private Function I_Get_SQL(ByRef ret_strResultMsg As String,
                             ByRef ret_dicUpdateLineProductionInfo As Dictionary(Of String, clsLineProduction_Info),
                             ByRef ret_lstClassProductionHist As List(Of clsClassProduction_HIST),
                             ByRef ret_lstLineProductionHist As List(Of clsLineProduction_Hist),
                             ByRef lstSQL As List(Of String),
                             ByRef lstQueueSQL As List(Of String)) As Boolean
    Try
      '取得LineProductionInfo的Update SQL
      For Each obj In ret_dicUpdateLineProductionInfo.Values
        If Not obj.O_Add_Update_SQLString(lstSQL) Then
          ret_strResultMsg = "Get Update WMS_CT_LINE_PRODUCTION_INFO SQL Failed"
          Return False
        End If
      Next

      '取得ClassProductionHist的Insert SQL
      For Each obj In ret_lstClassProductionHist
        If Not obj.O_Add_Insert_SQLString(lstQueueSQL) Then
          ret_strResultMsg = "Get Insert WMS_CH_CLASS_PRODUCTION_HIST SQL Failed"
          Return False
        End If
      Next

      '取得LineProductionHist的Insert SQL
      For Each obj In ret_lstLineProductionHist
        If Not obj.O_Add_Insert_SQLString(lstQueueSQL) Then
          ret_strResultMsg = "Get Insert WMS_CH_LINE_PRODUCTION_HIST SQL Failed"
          Return False
        End If
      Next

      Return True
    Catch ex As Exception
      ret_strResultMsg = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行資料更新
  Private Function I_Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                        ByRef ret_dicUpdateLineProductionInfo As Dictionary(Of String, clsLineProduction_Info),
                                        ByRef lstSQL As List(Of String),
                                        ByRef lstQueueSQL As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSQL.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
        '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
        ret_strResultMsg = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        '更新DB失敗則回傳False
        'ret_strResult_Message = "WMS Update DB Failed"
        ret_strResultMsg = "WMS 更新资料库失败"
        Return False
      End If
      Common_DBManagement.AddQueued(lstQueueSQL)
      '修改記憶體資料
      '新增LineProductionInf資訊
      '更新LineProductionInfo資訊
      For Each objNew As clsLineProduction_Info In ret_dicUpdateLineProductionInfo.Values
        Dim obj As clsLineProduction_Info = Nothing
        If gMain.objHandling.gdicLineProduction_Info.TryGetValue(objNew.gid, obj) Then
          obj.Update_To_Memory(objNew)
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = "Other Error"
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
