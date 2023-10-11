Imports eCA_TransactionMessage
Imports eCA_HostObject
''' <summary>
''' 20181117
''' V1.0.0
''' Mark
''' 處理所有傳入的Message
''' </summary>
Module Module_ProcessMessage
  Public Function O_ProcessCommand(ByVal strFunction_ID As String,
                                   ByVal strXmlMessage As String,
                                   ByRef ret_strResultMsg As String,
                                   ByVal contentType As enuHTTPContentType,
                                   Optional ByRef ret_strWait_UUID As String = "") As Boolean
    Try
      SendMessageToLog("ProcessCommand Function_ID=" & strFunction_ID & ", Message=" & strXmlMessage, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Dim blnProcessResult As Boolean = False

      'TODO SELECT CASE裡的 EXCUTE 應該要改寫。
      'EXCUTE 要改成呼叫 EX:O_Send_ToNSCommand 功能即可，讓開關去控制回傳的方法

      Select Case strFunction_ID
#Region "Message"
        'WMS
        Case enuWMSMessageFunctionID.T5F1U90_WOExcuting.ToString
          If O_ProcessResult_T5F1U90_WOExcuting(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        'MCS
        Case enuMCSMessageFunctionID.T3F4R2_DeviceAlarmReport.ToString
          If O_Process_T3F4R2_DeviceAlarmReport(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case enuMCSMessageFunctionID.T3F5R1_LineStatusChangeReport.ToString
          If O_Process_T3F5R1_LineStatusChangeReport(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case enuMCSMessageFunctionID.T3F5R2_LineInfoReport.ToString
          If O_Process_T3F5R2_LineInfoReport(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case enuMCSMessageFunctionID.T3F5R3_LineInProductionInfoReport.ToString
          If O_Process_T3F5R3_LineInProductionInfoReport(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case enuMCSMessageFunctionID.T3F5R4_LineInProductionInfoReset.ToString
          If O_Process_T3F5R4_LineInProductionInfoReset(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        'UI
        Case enuGUICommandFunctionID.T3F5U1_MaintenanceSet.ToString
          If O_Process_T3F5U1_MaintenanceSet(strXmlMessage, ret_strResultMsg) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T3F5U2_Maintenance.ToString
          If O_Process_T3F5U2_Maintenance(strXmlMessage, ret_strResultMsg) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T3F5U3_LineBigDataAlarmSet.ToString
          If O_Process_T3F5U3_LineBigDataAlarmSet(strXmlMessage, ret_strResultMsg) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T3F5U4_ProductionCountSet.ToString
          If O_Process_T3F5U4_ProductionCountSet(strXmlMessage, ret_strResultMsg) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T3F5U5_ClassProductionSet.ToString
          If O_Process_T3F5U5_ClassProductionSet(strXmlMessage, ret_strResultMsg) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T11F1U11_ProducePOExecution.ToString
          If O_Process_T11F1U11_ProducePOExecution(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If

        Case enuGUICommandFunctionID.T11F1U1_PODownload.ToString
          If O_Process_T11F1U1_PODownload(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T5F1U4_PODownload.ToString
          If O_Process_T5F1U4_PODownload(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T11F2U1_InventoryComparison.ToString
          If O_Process_T11F2U1_InventoryComparison(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T11F1U12_StocktakingExecution.ToString
          If O_Process_T11F1U12_StocktakingExecution(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T11F1U2_POExecution.ToString
          If O_Process_T11F1U2_POExecution(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T5F1U11_POExecution.ToString
          If O_Process_T5F1U11_POExecution(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T10F4U1_MainFileImport.ToString
          If O_Process_T10F4U1_MainFileImport(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T11F3U1_StocktakingDownload.ToString
          If O_Process_T11F3U1_StocktakingDownload(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) Then
            blnProcessResult = True
          End If

        Case enuGUICommandFunctionID.T5F1U18_POToWOOneToOne.ToString
          If O_Process_T5F1U18_POToWOOneToOne(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T5F1U19_POToWOOneSerialToOneWO.ToString
          If O_Process_T5F1U19_POToWOOneSerialToOneWO(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T6F5U1_ItemLabelManagement.ToString
          If O_Process_T6F5U1_ItemLabelManagement(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
        Case enuGUICommandFunctionID.T6F5U2_ItemLabelPrint.ToString
          If O_Process_T6F5U2_ItemLabelPrint(strXmlMessage, ret_strResultMsg, ret_strWait_UUID) = True Then
            blnProcessResult = True
          End If
#End Region
#Region "WCF"

#End Region
#Region "WebAPI"

#End Region
        Case Else
          blnProcessResult = False
          ret_strResultMsg = "Not Defines Function_ID, Function_ID=" & strFunction_ID
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      End Select
      Return blnProcessResult
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '處理所有回覆結果的Message
  Public Function O_ProcessCommandResult(ByVal UUID As String,
                                         ByVal strFunction_ID As String,
                                         ByVal strXmlMessage As String,
                                         ByRef strRejectReason As String,
                                         ByRef blnResult As Boolean,
                                         ByRef ret_strResultMsg As String) As Boolean
    Try
      SendMessageToLog("ProcessCommandResult Function_ID=" & strFunction_ID & ", Result=" & blnResult.ToString, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Dim blnProcessResult As Boolean = False
      Select Case strFunction_ID
        'Case enuHostCommandFunctionID.T5F3U23_POToWO.ToString
        '  If O_ProcessResult_T5F3U23_POToWO(strXmlMessage, ret_strResultMsg, strRejectReason, blnResult) = True Then
        '    blnProcessResult = True
        '  End If
        'Case enuHostCommandFunctionID.T5F1U1_POManagement.ToString
        '  If O_ProcessResult_T5F1U1_POManagement(strXmlMessage, ret_strResultMsg, strRejectReason, blnResult) = True Then
        '    blnProcessResult = True
        '  End If
        'Case enuHostCommandFunctionID.T5F2U62_AutoInbound.ToString
        '  If O_ProcessResult_T5F2U62_AutoInbound(strXmlMessage, ret_strResultMsg, strRejectReason, blnResult) = True Then
        '    blnProcessResult = True
        '  End If
        Case enuHostCommandFunctionID.T5F1U1_POManagement.ToString
          If O_ProcessResult_T5F1U1_PO_Management(strXmlMessage, ret_strResultMsg, strRejectReason, blnResult) Then
            blnProcessResult = True
          End If
        Case enuHostCommandFunctionID.T5F5U1_TransactionOederManagement.ToString
          If O_ProcessResult_T5F5U1_TransactionOederManagement(strXmlMessage, ret_strResultMsg, strRejectReason, blnResult) Then
            blnProcessResult = True
          End If
        Case enuHostCommandFunctionID.T2F3U1_SKUManagement.ToString
          If O_ProcessResult_T2F3U1_SKUManagement(strXmlMessage, ret_strResultMsg, strRejectReason, blnResult) Then
            blnProcessResult = True
          End If
        Case Else
          blnProcessResult = blnResult
      End Select

      '檢查是否需要Report
      If O_Set_WaitCommandResult(UUID, blnResult, strRejectReason) = False Then
        Return False
      End If

      Return blnProcessResult
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '處理所有要傳出去給Host的Message
  Public Function O_Send_ToWMSCommand(ByVal strMessage As String,
                                       ByVal HeaderInfo As clsHeader,
                                      ByVal dicHost_Command As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try
      Dim strLog As String = ""
      'Select Case ModuleDeclaration.WMSToHostHandlerInterfaceType
      Select Case enuHandlingInterfaceType.DB       '目前先暫定使用DB的方式
        Case enuHandlingInterfaceType.DB
          If O_Send_MessageToWMS_ByDB(strMessage, HeaderInfo, dicHost_Command) = True Then
            strLog = String.Format("Send Message To MCS Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog("Send Message To MCS Success, ", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          End If
        Case enuHandlingInterfaceType.MQ

        Case enuHandlingInterfaceType.WebService

      End Select
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '舊版(O_Send_ToWMSCommand)有 dicHost_Command => 因為舊版要把DIC回傳之後才做SQL動作，新版的統一由這隻函式做，因此不需要了
  Public Function O_Send_ToWMSCommand_N(ByRef strMessage As String,
                                        ByVal HeaderInfo As clsHeader) As Boolean
    Try
      Dim strLog As String = ""
      Select Case ModuleDeclaration.HandlingToWMSInterfaceType
        Case enuHandlingInterfaceType.DB
          If O_Send_MessageToOther_ByDB(strMessage, HeaderInfo, enuSystemType.WMS) = True Then
            strLog = String.Format("Send Message To WMS By DB Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case enuHandlingInterfaceType.MQ
          If O_Send_MessageToOther_ByMQ(strMessage, HeaderInfo, enuSystemType.WMS) = True Then
            strLog = String.Format("Send Message To WMS By MQ Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case enuHandlingInterfaceType.WebAPI
          '因為除了DB跟MQ以外的方式，都是用HTTP REQUEST做，所以用ELSE把剩下的包起來就好
          If O_Send_MessageToOther_ByHTTP(strMessage, HeaderInfo, enuSystemType.WMS, enuHTTPContentType.XML) = True Then
            strLog = String.Format("Send Message To WMS By WebAPI Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case Else
          SendMessageToLog($"沒有對應的拋送方式，InterfaceType : {ModuleDeclaration.HandlingToWMSInterfaceType}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
      End Select
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '處理所有要傳出去給MCS的Message
  Public Function O_Send_ToMCSCommand(ByVal strMessage As String,
                                     ByVal HeaderInfo As clsHeader) As Boolean
    Try
      Dim strLog As String = ""
      Select Case ModuleDeclaration.HandlingToMCSInterfaceType
        Case enuHandlingInterfaceType.DB
          If O_Send_MessageToOther_ByDB(strMessage, HeaderInfo, enuSystemType.MCS) = True Then
            strLog = String.Format("Send Message To MCS By DB Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case enuHandlingInterfaceType.MQ
          If O_Send_MessageToOther_ByMQ(strMessage, HeaderInfo, enuSystemType.MCS) = True Then
            strLog = String.Format("Send Message To MCS By MQ Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case enuHandlingInterfaceType.WebAPI
          If O_Send_MessageToOther_ByHTTP(strMessage, HeaderInfo, enuSystemType.MCS, enuHTTPContentType.XML) = True Then
            strLog = String.Format("Send Message To MCS By WebAPI Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case Else
          SendMessageToLog($"沒有對應的拋送方式，InterfaceType : {ModuleDeclaration.HandlingToWMSInterfaceType}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
      End Select
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '處理所有要傳出去給MCS的Message
  Public Function O_Send_ToGUICommand(ByVal strMessage As String,
                                     ByVal HeaderInfo As clsHeader) As Boolean
    Try
      Dim strLog As String = ""
      Select Case ModuleDeclaration.HandlingToGUIInterfaceType
        Case enuHandlingInterfaceType.DB
          If O_Send_MessageToOther_ByDB(strMessage, HeaderInfo, enuSystemType.GUI) = True Then
            strLog = String.Format("Send Message To GUI By DB Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case enuHandlingInterfaceType.MQ
          If O_Send_MessageToOther_ByMQ(strMessage, HeaderInfo, enuSystemType.GUI) = True Then
            strLog = String.Format("Send Message To GUI By MQ Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case enuHandlingInterfaceType.WebAPI
          If O_Send_MessageToOther_ByHTTP(strMessage, HeaderInfo, enuSystemType.GUI, enuHTTPContentType.XML) = True Then
            strLog = String.Format("Send Message To GUI By WebAPI Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case Else
          SendMessageToLog($"沒有對應的拋送方式，InterfaceType : {ModuleDeclaration.HandlingToWMSInterfaceType}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
      End Select
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Send_ToNSCommand(ByVal strMessage As String,
                                     ByVal HeaderInfo As clsHeader) As Boolean
    Try
      Dim strLog As String = ""
      Select Case ModuleDeclaration.HandlingToNSInterfaceType
        Case enuHandlingInterfaceType.DB
          If O_Send_MessageToOther_ByDB(strMessage, HeaderInfo, enuSystemType.NS) = True Then
            strLog = String.Format("Send Message To NS By DB Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case enuHandlingInterfaceType.MQ
          If O_Send_MessageToOther_ByMQ(strMessage, HeaderInfo, enuSystemType.NS) = True Then
            strLog = String.Format("Send Message To NS By MQ Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case enuHandlingInterfaceType.WebAPI
          If O_Send_MessageToOther_ByHTTP(strMessage, HeaderInfo, enuSystemType.NS, enuHTTPContentType.XML) = True Then
            strLog = String.Format("Send Message To NS By WebAPI Sucess, UUID=<{0}>, EventID=<{1}>, Message=<{2}>", HeaderInfo.UUID, HeaderInfo.EventID, strMessage)
            SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          Else
            Return False
          End If
        Case Else
          SendMessageToLog($"沒有對應的拋送方式，InterfaceType : {ModuleDeclaration.HandlingToWMSInterfaceType}", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
      End Select

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~處理來自異質系統的Message~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

  'MCS
  ''' <summary>
  ''' 處理MCS上報的T3F4R2_DeviceAlarmReport事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <param name="blnResult"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F4R2_DeviceAlarmReport(ByVal strXmlMessage As String,
                                                     ByRef ret_strResultMsg As String,
                                                     Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T3F4R2_DeviceAlarmReport = Nothing
      If ParseXmlString.ParseMessage_T3F4R2_DeviceAlarmReport(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F4R2_DeviceAlarmReport.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("Module_T3F4R2_DeviceAlarmReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F4R2_DeviceAlarmReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理MCS上報的T3F5R1_LineStatusChangeReport事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <param name="blnResult"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5R1_LineStatusChangeReport(ByVal strXmlMessage As String,
                                                          ByRef ret_strResultMsg As String,
                                                          Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T3F5R1_LineStatusChangeReport = Nothing
      If ParseXmlString.ParseMessage_T3F5R1_LineStatusChangeReport(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5R1_LineStatusChangeReport.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("Module_T3F5R1_LineStatusChangeReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5R1_LineStatusChangeReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理MCS上報的T3F5R2_LineInfoReport事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <param name="blnResult"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5R2_LineInfoReport(ByVal strXmlMessage As String,
                                                  ByRef ret_strResultMsg As String,
                                                  Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T3F5R2_LineInfoReport = Nothing
      If ParseXmlString.ParseMessage_T3F5R2_LineInfoReport(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5R2_LineInfoReport.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("Module_T3F5R2_LineInfoReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5R2_LineInfoReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理MCS上報的T3F5R3_LineInProductionInfoReport事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <param name="blnResult"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5R3_LineInProductionInfoReport(ByVal strXmlMessage As String,
                                                              ByRef ret_strResultMsg As String,
                                                              Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T3F5R3_LineInProductionInfoReport = Nothing
      If ParseXmlString.ParseMessage_T3F5R3_LineInProductionInfoReport(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5R3_LineInProductionInfoReport.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("Module_T3F5R3_LineInProductionInfoReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5R3_LineInProductionInfoReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理MCS上報的T3F5R4_LineInProductionInfoReset事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <param name="blnResult"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5R4_LineInProductionInfoReset(ByVal strXmlMessage As String,
                                                             ByRef ret_strResultMsg As String,
                                                             Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T3F5R4_LineInProductionInfoReset = Nothing
      If ParseXmlString.ParseMessage_T3F5R4_LineInProductionInfoReset(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5R4_LineInProductionInfoReset.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("Module_T3F5R4_LineInProductionInfoReset Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5R4_LineInProductionInfoReset Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''' <summary>
  ''' 處理GUI報的T3F5U1_MaintenanceSet事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5U1_MaintenanceSet(ByVal strXmlMessage As String, ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T3F5U1_MaintenanceSet = Nothing
      If ParseXmlString.ParseMessage_T3F5U1_MaintenanceSet(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5U1_MaintenanceSet.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T3F5U1_MaintenanceSet Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5U1_MaintenanceSetFailed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理GUI報的T3F5U2_Maintenance事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5U2_Maintenance(ByVal strXmlMessage As String, ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T3F5U2_Maintenance = Nothing
      If ParseXmlString.ParseMessage_T3F5U2_Maintenance(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5U2_Maintenance.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T3F5U2_Maintenance Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5U2_MaintenanceFailed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理GUI報的T3F5U3_LineBigDataAlarmSet事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5U3_LineBigDataAlarmSet(ByVal strXmlMessage As String, ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T3F5U3_LineBigDataAlarmSet = Nothing
      If ParseXmlString.ParseMessage_T3F5U3_LineBigDataAlarmSet(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5U3_LineBigDataAlarmSet.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T3F5U3_LineBigDataAlarmSet Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5U3_LineBigDataAlarmSetFailed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理GUI報的T3F5U4_ProductionCountSet事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5U4_ProductionCountSet(ByVal strXmlMessage As String, ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T3F5U4_ProductionCountSet = Nothing
      If ParseXmlString.ParseMessage_T3F5U4_ProductionCountSet(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5U4_ProductionCountSet.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T3F5U4_ProductionCountSet Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5U4_ProductionCountSetFailed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理GUI報的T3F5U5_ClassProductionSet事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <returns></returns>
  Public Function O_Process_T3F5U5_ClassProductionSet(ByVal strXmlMessage As String, ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T3F5U5_ClassProductionSet = Nothing
      If ParseXmlString.ParseMessage_T3F5U5_ClassProductionSet(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T3F5U5_ClassProductionSet.O_Process_Message(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T3F5U5_ClassProductionSet Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T3F5U5_ClassProductionSetFailed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理GUI報的T11F1U11_ProducePOExecution事件
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <param name="ret_strWaitUUID"></param>
  ''' <returns></returns>
  Public Function O_Process_T11F1U11_ProducePOExecution(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T11F1U11_ProducePOExecution = Nothing
      If ParseXmlString.ParseMessage_T11F1U11_ProducePOExecution(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T11F1U11_ProducePOExecution.O_Process_Message(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T11F1U11_ProducePOExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T11F1U11_ProducePOExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T11F1U1_PODownload(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T11F1U1_PODownload = Nothing
      If ParseXmlString.ParseMessage_T11F1U1_PODownload(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T11F1U1_PODownload.O_T11F1U1_PODownload(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T11F1U1_PODownload Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T11F1U1_PODownload Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T5F1U4_PODownload(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T5F1U4_PODownload = Nothing
      If ParseXmlString.ParseMessage_T5F1U4_PODownload(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T5F1U4_PODownload.O_T5F1U4_PODownload(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T5F1U4_PODownload Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T5F1U4_PODownload Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T11F2U1_InventoryComparison(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T11F2U1_InventoryComparison = Nothing
      If ParseXmlString.ParseMessage_T11F2U1_InventoryComparison(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T11F2U1_InventoryComparison.O_Process(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T11F2U1_InventoryComparison Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T11F2U1_InventoryComparison, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function



  Public Function O_Process_T11F1U12_StocktakingExecution(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T11F1U12_StocktakingExecution = Nothing
      If ParseXmlString.ParseMessage_T11F1U12_StocktakingExecution(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T11F1U12_StocktakingExecution.O_Process_Message(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message T11F1U12_StocktakingExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T11F1U12_StocktakingExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  Public Function O_Process_T11F1U2_POExecution(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T11F1U2_POExecution = Nothing
      If ParseXmlString.ParseMessage_T11F1U2_POExecution(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T11F1U2_POExecution.O_T11F1U2_POExecution(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T11F1U2_POExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T11F1U2_POExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T5F1U11_POExecution(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T5F1U11_POExecution = Nothing
      If ParseXmlString.ParseMessage_T5F1U11_POExecution(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T5F1U11_POExecution.O_T5F1U11_POExecution(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T5F1U11_POExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T5F1U11_POExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T5F1U18_POToWOOneToOne(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T5F1U18_POToWOOneToOne = Nothing
      Dim lstSql As New List(Of String)
      If ParseXmlString.ParseMessage_T5F1U18_POToWOOneToOne(strXmlMessage, obj, ret_strResultMsg) = True Then
        Try 'Host不檢查內容 直接送給WMS
          For Each POInfo In obj.Body.POList.POInfo
            If ExcutePO.ContainsKey(POInfo.PO_ID) = False Then ExcutePO.Add(POInfo.PO_ID, POInfo.PO_ID) '排除時間差問題 20190628
          Next
        Catch ex As Exception
          ret_strResultMsg = "POInfo 資料異常"
          Return False
        End Try

        Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
        If O_Send_ToWMSCommand(strXmlMessage, obj.Header, Host_Command) = False Then
          gMain.SendMessageToLog("轉發 Module_T5F1U18_POToWOOneToOne Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        For Each objHost_Command In Host_Command.Values
          If objHost_Command.O_Add_Insert_SQLString(lstSql) = False Then
            ret_strResultMsg = "Get Insert HOST_T_WMS_Command SQL Failed"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        Next
        If Common_DBManagement.BatchUpdate(lstSql) = False Then
          '更新DB失敗則回傳False
          ret_strResultMsg = "eHOST 更新资料库失败"
          Return False
        End If
        Return True
      Else
        gMain.SendMessageToLog("ParseMessage_T5F1U18_POToWOOneToOne Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      Return False
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T5F1U19_POToWOOneSerialToOneWO(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T5F1U19_POToWOOneSerialToOneWO = Nothing
      Dim lstSql As New List(Of String)
      If ParseXmlString.ParseMessage_T5F1U19_POToWOOneSerialToOneWO(strXmlMessage, obj, ret_strResultMsg) = True Then
        Try 'Host不檢查內容 直接送給WMS
          For Each POInfo In obj.Body.POList.POInfo
            If ExcutePO.ContainsKey(POInfo.PO_ID) = False Then ExcutePO.Add(POInfo.PO_ID, POInfo.PO_ID) '排除時間差問題 20190628
          Next
        Catch ex As Exception
          ret_strResultMsg = "POInfo 資料異常"
          Return False
        End Try

        Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
        If O_Send_ToWMSCommand(strXmlMessage, obj.Header, Host_Command) = False Then
          gMain.SendMessageToLog("轉發 Module_T5F1U19_POToWOOneSerialToOneWO Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        For Each objHost_Command In Host_Command.Values
          If objHost_Command.O_Add_Insert_SQLString(lstSql) = False Then
            ret_strResultMsg = "Get Insert HOST_T_WMS_Command SQL Failed"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        Next
        If Common_DBManagement.BatchUpdate(lstSql) = False Then
          '更新DB失敗則回傳False
          ret_strResultMsg = "eHOST 更新资料库失败"
          Return False
        End If
        Return True
      Else
        gMain.SendMessageToLog("ParseMessage_T5F1U19_POToWOOneSerialToOneWO Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      Return False

    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  Public Function O_Process_T10F4U1_MainFileImport(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T10F4U1_MainFileImport = Nothing
      If ParseXmlString.ParseMessage_T10F4U1_MainFileImport(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T10F4U1_MainFileImport.O_T10F4U1_MainFileImport(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T10F4U1_MainFileImport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T10F4U1_MainFileImport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Process_T11F3U1_StocktakingDownload(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T11F3U1_StocktakingDownload = Nothing
      If ParseXmlString.ParseMessage_T11F3U1_StocktakingDownload(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T11F3U1_StocktakingDownload.O_T11F3U1_StocktakingDownload(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T11F3U1_StocktakingDownload Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T11F3U1_StocktakingDownload Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T6F5U1_ItemLabelManagement(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T6F5U1_ItemLabelManagement = Nothing
      If ParseXmlString.ParseMessage_T6F5U1_ItemLabelManagement(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T6F5U1_ItemLabelManagement.O_T6F5U1_ItemLabelManagement(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T6F5U1_ItemLabelManagement Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T6F5U1_ItemLabelManagement Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T6F5U2_ItemLabelPrint(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef ret_strWaitUUID As String) As Boolean
    Try
      Dim obj As MSG_T6F5U2_ItemLabelPrint = Nothing
      If ParseXmlString.ParseMessage_T6F5U2_ItemLabelPrint(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T6F5U2_ItemLabelPrint.O_T6F5U2_ItemLabelPrint(obj, ret_strResultMsg, ret_strWaitUUID) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_Message Module_T6F5U1_ItemLabelManagement Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T6F5U1_ItemLabelManagement Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~處理發送給異質系統後的回覆~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  Public Function O_ProcessResult_T5F3U23_POToWO(ByVal strXmlMessage As String,
                                                 ByRef ret_strResultMsg As String,
                                                 ByRef strRejectReason As String,
                                                 Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T5F3U23_POToWO = Nothing
      '失败则不解格式
      If blnResult = False Then
        ret_strResultMsg = "T5F3U23_POToWO(詢問執行PO轉WO) Failed, WMS回复资讯：" & strRejectReason
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '返回成功 往下继续做
      If ParseXmlString.ParseMessage_T5F3U23_POToWO(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T5F3U23_POToWO.O_CheckMessageResult(obj, ret_strResultMsg, strRejectReason, blnResult) = True Then
          Return True
        Else
          gMain.SendMessageToLog("T5F3U23_POToWO O_CheckMessageResult Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T5F3U23_POToWO Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~處理發送給異質系統後的回覆~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  Public Function O_ProcessResult_T5F1U1_POManagement(ByVal strXmlMessage As String,
                                                 ByRef ret_strResultMsg As String,
                                                 ByRef strRejectReason As String,
                                                 Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T5F1U1_PO_Management = Nothing
      '失败则不解格式
      If blnResult = False Then
        ret_strResultMsg = "MSG_T5F1U1_PO_Management(建單流程) Failed, WMS回复资讯：" & strRejectReason
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '返回成功 往下继续做
      If ParseXmlString.ParseMessage_T5F1U1_POManagement(strXmlMessage, obj, ret_strResultMsg) = True Then
        Dim ret_strWaitUUID = String.Empty
        '執行相對應的事件
        If Module_T5F1U1_POManagement.O_CheckMessageResult(obj, ret_strResultMsg, strRejectReason, ret_strWaitUUID, blnResult) = True Then
          Return True
        Else
          gMain.SendMessageToLog("T5F1U1_PO_Management O_CheckMessageResult Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T5F1U1_PO_Management Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~處理發送給異質系統後的回覆~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  Public Function O_ProcessResult_T5F1U1_PO_Management(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef strRejectReason As String,
                                                        Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T5F1U1_PO_Management = Nothing
      '失败则不解格式
      If blnResult = False Then
        ret_strResultMsg = "T5F1U1_PO_Management(單據新增、刪除、修改) Failed, WMS回复资讯：" & strRejectReason
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        'Return False
      End If
      '返回成功 往下继续做
      If ParseXmlString.ParseMessage_T5F1U1_POManagement(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_POManagement_HTG_Result.O_CheckMessageResult(obj, blnResult, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("T5F1U1_PO_Management O_CheckMessageResult Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("T5F1U1_PO_Management Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_ProcessResult_T5F5U1_TransactionOederManagement(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef strRejectReason As String,
                                                        Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T5F5U1_TransactionOederManagement = Nothing
      '失败则不解格式
      If blnResult = False Then
        ret_strResultMsg = "T5F5U1_TransactionOederManagement(單據新增、刪除、修改) Failed, WMS回复资讯：" & strRejectReason
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '返回成功 往下继续做
      If ParseXmlString.ParseMessage_T5F5U1_TransactionOederManagement(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_TransactionOederManagement_HTG_Result.O_CheckMessageResult(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_ProcessResult_T5F5U1_TransactionOederManagement O_CheckMessageResult Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("O_ProcessResult_T5F5U1_TransactionOederManagement Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_ProcessResult_T2F3U1_SKUManagement(ByVal strXmlMessage As String,
                                                        ByRef ret_strResultMsg As String,
                                                        ByRef strRejectReason As String,
                                                        Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T2F3U1_SKUManagement = Nothing
      '失败则不解格式
      If blnResult = False Then
        ret_strResultMsg = "T2F3U1_SKUManagement(單據新增、刪除、修改) Failed, WMS回复资讯：" & strRejectReason
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '返回成功 往下继续做
      If ParseXmlString.ParseMessage_T2F3U1_SKUManagement(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_SKUManagement_HTG_Result.O_CheckMessageResult(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_ProcessResult_T2F3U1_SKUManagement O_CheckMessageResult Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("O_ProcessResult_T2F3U1_SKUManagement Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_ProcessResult_T5F1U62_AutoInbound(ByVal strXmlMessage As String,
                                                 ByRef ret_strResultMsg As String,
                                                 ByRef strRejectReason As String,
                                                 Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T5F2U62_AutoInbound = Nothing
      '失败则不解格式
      If blnResult = False Then
        ret_strResultMsg = "T5F2U62_AutoInbound(詢問執行PO轉WO) Failed, WMS回复资讯：" & strRejectReason
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '返回成功 往下继续做
      If ParseXmlString.ParseMessage_T5F2U62_AutoInbound(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T5F2U62_AutoInbound.O_CheckMessageResult(obj, ret_strResultMsg, strRejectReason, blnResult) = True Then
          Return True
        Else
          gMain.SendMessageToLog("T5F2U62_AutoInbound O_CheckMessageResult Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T5F2U62_AutoInbound Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If

    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_ProcessResult_T5F1U90_WOExcuting(ByVal strXmlMessage As String,
                                                     ByRef ret_strResultMsg As String,
                                                     Optional ByRef blnResult As Boolean = True) As Boolean
    Try
      Dim obj As MSG_T5F1S90_WOExcuting = ParseXmlStringToClass(Of MSG_T5F1S90_WOExcuting)(strXmlMessage)
      If obj IsNot Nothing Then
        '執行相對應的事件
        If Module_T5F1U90_WOExcuting.O_T5F1U90_WOExcuting(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("Module_T5F1U90_WOExcuting Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T5F1U90_WOExcuting Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''' <summary>
  ''' 取得等待回傳的Command，回傳結果
  ''' </summary>
  ''' <param name="UUID"></param>
  ''' <param name="bln_Result"></param>
  ''' <param name="str_RejectReason"></param>
  ''' <returns></returns>
  Public Function O_Set_WaitCommandResult(ByVal UUID As String, ByVal bln_Result As Boolean, ByVal str_RejectReason As String) As Boolean
    Try
      '檢查是否有回傳
      Dim objHOST_T_COMMAND_REPORT As clsHOST_T_COMMAND_REPORT = Nothing
      If O_Get_objHOST_T_COMMAND_REPORTByUUID(gMain.objHandling.gdicHOST_T_COMMAND_REPORT, UUID, objHOST_T_COMMAND_REPORT) = False Then
        Return True '沒有需要處理的
      End If

      'DB的部分自己處理掉了，不需要使用QUEUE處理

      '回覆對應的系統
      Select Case objHOST_T_COMMAND_REPORT.REPORT_SYSTEM_TYPE
        Case enuSystemType.GUI
          Select Case ModuleDeclaration.HandlingToGUIInterfaceType
            Case enuHandlingInterfaceType.DB
              'O_Set_GUI_DBWaitCommandResult(objCommandReport, bln_Result, str_RejectReason)
            Case enuHandlingInterfaceType.MQ
              O_Set_GUI_MQWaitCommandResult(objHOST_T_COMMAND_REPORT, UUID, bln_Result, str_RejectReason)
            Case enuHandlingInterfaceType.WebAPI

          End Select
        Case enuSystemType.WMS
          Select Case ModuleDeclaration.HandlingToWMSInterfaceType
            Case enuHandlingInterfaceType.DB
              'O_Set_HOST_DBWaitCommandResult(objCommandReport, bln_Result, str_RejectReason)
            Case enuHandlingInterfaceType.MQ
              O_Set_WMS_MQWaitCommandResult(objHOST_T_COMMAND_REPORT, UUID, bln_Result, str_RejectReason)
            Case enuHandlingInterfaceType.WebAPI

          End Select
        Case enuSystemType.MCS
          Select Case ModuleDeclaration.HandlingToMCSInterfaceType
            Case enuHandlingInterfaceType.DB
              'O_Set_MCS_DBWaitCommandResult(objCommandReport, bln_Result, str_RejectReason)
            Case enuHandlingInterfaceType.MQ
              O_Set_MCS_MQWaitCommandResult(objHOST_T_COMMAND_REPORT, UUID, bln_Result, str_RejectReason)
            Case enuHandlingInterfaceType.WebAPI

          End Select
      End Select

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module