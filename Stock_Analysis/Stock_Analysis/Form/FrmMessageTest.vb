Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports ClsConfigTool
Imports System.IO
'Imports NDde
'Imports NDde.Client
Imports System.Net.Mail
Imports eCA_TransactionMessage
Imports System.Xml.Serialization

Public Class FrmMessageTest
  Public MessageName As enuMessageName
  '判別是否被開啟過的參數
  Private Shared dicFormType As Dictionary(Of enuMessageName, Object) = New Dictionary(Of enuMessageName, Object)

  Sub New(ByVal New_MessageName As enuMessageName)
    Try
      MessageName = New_MessageName
      InitializeComponent()
      '設定Form的標頭
      Me.Text = MessageName.ToString

    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Private Sub btn_SendMessage_Click(sender As Object, e As EventArgs) Handles btn_SendMessage.Click
    Try
      Dim New_Msg As String = txt_SendMessage.Text
      Dim RetMsg As String = ""
      Dim strLog As String = String.Format("XML String ={0}", New_Msg)
      SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Select Case MessageName
        Case enuMessageName.T3F4R2_DeviceAlarmReport
          Dim objMSG As eCA_TransactionMessage.MSG_T3F4R2_DeviceAlarmReport = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F4R2_DeviceAlarmReport(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F4R2_DeviceAlarmReport.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5R1LineStatusChangeReport
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5R1_LineStatusChangeReport = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5R1_LineStatusChangeReport(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5R1_LineStatusChangeReport.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5R2_LineInfoReport
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5R2_LineInfoReport = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5R2_LineInfoReport(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5R2_LineInfoReport.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5R3_LineInProductionInfoReport
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5R3_LineInProductionInfoReport = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5R3_LineInProductionInfoReport(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5R3_LineInProductionInfoReport.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5R4LineInProductionInfoReset
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5R4_LineInProductionInfoReset = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5R4_LineInProductionInfoReset(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5R4_LineInProductionInfoReset.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5U1_MaintenanceSet
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5U1_MaintenanceSet = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5U1_MaintenanceSet(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5U1_MaintenanceSet.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5U2_Maintenance
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5U2_Maintenance = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5U2_Maintenance(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5U2_Maintenance.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5U3_LineBigDataAlarmSet
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5U3_LineBigDataAlarmSet = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5U3_LineBigDataAlarmSet(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5U3_LineBigDataAlarmSet.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5U4_ProductionCountSet
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5U4_ProductionCountSet = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5U4_ProductionCountSet(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5U4_ProductionCountSet.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T3F5U5_ClassProductionSet
          Dim objMSG As eCA_TransactionMessage.MSG_T3F5U5_ClassProductionSet = Nothing
          If eCA_TransactionMessage.ParseMessage_T3F5U5_ClassProductionSet(New_Msg, objMSG, RetMsg) = True Then
            If Module_T3F5U5_ClassProductionSet.O_Process_Message(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T11F1U11_ProducePOExecution
          Dim objMSG As eCA_TransactionMessage.MSG_T11F1U11_ProducePOExecution = Nothing
          If eCA_TransactionMessage.ParseMessage_T11F1U11_ProducePOExecution(New_Msg, objMSG, RetMsg) = True Then
            If Module_T11F1U11_ProducePOExecution.O_Process_Message(objMSG, RetMsg, "") = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T5F1U11_POExecution
          Dim objMSG As eCA_TransactionMessage.MSG_T5F1U11_POExecution = Nothing
          If eCA_TransactionMessage.ParseMessage_T5F1U11_POExecution(New_Msg, objMSG, RetMsg) = True Then
            If Module_T5F1U11_POExecution.O_T5F1U11_POExecution(objMSG, RetMsg, "") = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T10F4U1_MainFileImport
          Dim objMSG As eCA_TransactionMessage.MSG_T10F4U1_MainFileImport = Nothing
          If eCA_TransactionMessage.ParseMessage_T10F4U1_MainFileImport(New_Msg, objMSG, RetMsg) = True Then
            If Module_T10F4U1_MainFileImport.O_T10F4U1_MainFileImport(objMSG, RetMsg, "") = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T6F5U1_ItemLabelManagement
          Dim objMSG As eCA_TransactionMessage.MSG_T6F5U1_ItemLabelManagement = Nothing
          If eCA_TransactionMessage.ParseMessage_T6F5U1_ItemLabelManagement(New_Msg, objMSG, RetMsg) = True Then
            If Module_T6F5U1_ItemLabelManagement.O_T6F5U1_ItemLabelManagement(objMSG, RetMsg, "") = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T6F5U2_ItemLabelPrint
          Dim objMSG As eCA_TransactionMessage.MSG_T6F5U2_ItemLabelPrint = Nothing
          If eCA_TransactionMessage.ParseMessage_T6F5U2_ItemLabelPrint(New_Msg, objMSG, RetMsg) = True Then
            If Module_T6F5U2_ItemLabelPrint.O_T6F5U2_ItemLabelPrint(objMSG, RetMsg, "") = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T11F1U1_PODownload
          Dim obj As MSG_T11F1U1_PODownload = Nothing
          If ParseXmlString.ParseMessage_T11F1U1_PODownload(New_Msg, obj, RetMsg) = True Then
            '執行相對應的事件
            If Module_T11F1U1_PODownload.O_T11F1U1_PODownload(obj, RetMsg, "") = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T11F1U2_POExecution
          Dim objMSG As eCA_TransactionMessage.MSG_T11F1U2_POExecution = Nothing

          If eCA_TransactionMessage.ParseMessage_T11F1U2_POExecution(New_Msg, objMSG, RetMsg) = True Then
            If Module_T11F1U2_POExecution.O_T11F1U2_POExecution(objMSG, RetMsg, "") = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If

        Case enuMessageName.T10F2S1_StocktakingReport
          Dim objMSG As eCA_TransactionMessage.MSG_T10F2S1_StocktakingReport = Nothing
          If eCA_TransactionMessage.ParseMessage_T10F2S1_StocktakingReport(New_Msg, objMSG, RetMsg) = True Then
            If Module_T10F2S1_StocktakingReport.O_T10F2S1_StocktakingReport(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.T5F1S1WOClose
          Dim objMSG As eCA_TransactionMessage.MSG_T5F1S1_WOClose = Nothing
          If eCA_TransactionMessage.ParseMessage_T5F1S1_WOClose(New_Msg, objMSG, RetMsg) = True Then
            If Module_T5F1S1_WOClose.O_T5F1S1_WOClose(objMSG, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.PO_Release
          Dim strXML As String = ""
          If ERP_REPORT_Release(Nothing, RetMsg) = False Then
            SendMessageToLog("單據放行失敗，錯誤內容:" & RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

          End If
        Case enuMessageName.StockTaking_Report
          Dim strXML As String = ""
          If ERP_REPORT_Stocktaking(Nothing, RetMsg) = False Then
            SendMessageToLog("盤點單回報失敗，錯誤內容:" & RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)

          End If
        Case enuMessageName.Test
          Dim strXML As String = New_Msg

          Dim tmp_objMSG As MSG_Stocktaking_Respones = Nothing
          tmp_objMSG = ParseXmlStringToClass(Of MSG_Stocktaking_Respones)(strXML.ToString)
          SendMessageToLog("盤點單回報失敗，錯誤內容:" & RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
#Region "BOM"
        Case enuMessageName.T5F1U90_WOExcuting
          Dim obj As MSG_T5F1S90_WOExcuting = ParseXmlStringToClass(Of MSG_T5F1S90_WOExcuting)(New_Msg)
          If obj IsNot Nothing Then
            If Module_T5F1U90_WOExcuting.O_T5F1U90_WOExcuting(obj, RetMsg) = True Then
              txt_ResultMessage.Text = RetMsg
            Else
              txt_ResultMessage.Text = RetMsg
            End If
          Else
            txt_ResultMessage.Text = RetMsg
          End If
        Case enuMessageName.TEST_HOST_T_COMMAND
          Dim strXmlMessage As String = "JUST TEST"

          Dim header = New clsHeader()
          header.UUID = "2to1"
          header.EventID = "JUST TEST"
          header.Direction = "Primary"

          header.ClientInfo = New clsHeader.clsClientInfo
          header.ClientInfo.ClientID = "Handler"
          header.ClientInfo.UserID = "TEST"
          header.ClientInfo.IP = ""
          header.ClientInfo.MachineID = ""

          'If O_Send_ToWMSCommand_N(strXmlMessage, header) = False Then
          '  gMain.SendMessageToLog("測試功能失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          'End If

          'header.UUID = "2to3"
          'If O_Send_ToMCSCommand(strXmlMessage, header) = False Then
          '  gMain.SendMessageToLog("測試功能失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          'End If

          header.UUID = "2to4"
          If O_Send_ToGUICommand(strXmlMessage, header) = False Then
            gMain.SendMessageToLog("測試功能失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If

          'header.UUID = "2to9"
          'If O_Send_ToNSCommand(strXmlMessage, header) = False Then
          '  gMain.SendMessageToLog("測試功能失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          'End If
        Case enuMessageName.MQ_PRIMARY
          gMain.RabbitMQ.SendMessage("TEST_UUID", "TEST_FUNCTION_ID", "BRBRBRBR...", Msg_Direction_Primary, "Handler", 1)
        Case enuMessageName.MQ_SECONDARY
          gMain.RabbitMQ.ResultSecondaryMessage("TEST_RESULT", "TEST_RESULT_MSSAGE", enuRabbitMQ.HOST_TO_GUI.ToString, "", "", gMain.dicGUI_TO_HANDLING_Queue.First.Value, strLog)
#End Region
          'Try
          '  Dim ret_T = New XmlSerializer(GetType(MSG_Header)).Deserialize(New StringReader(_XmlString))
          '  Return ret_T
          'Catch ex As Exception
          '  SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
          '  Return Nothing
          'End Try

      End Select
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Enum enuMessageName
    T10F4U1_MainFileImport
    T5F1U6_ReceiptManagement
    T5F3U1_POManagement
    T5F3U3_POToWO
    T2F1U1_CompanyInfoManagement
    T2F2U1_CustomerManagement
    T2F3U1_SKUManagement
    T2F4U1_ClientManagement
    T5F2U1_CarrierInstall
    T5F2U2_CarrierRemove
    T5F2U3_CarrierLocationChange
    T3F1U1_VehicleMaintenance
    T3F1U2_VehicleManagement
    T3F3U1_ZoneMaintenance
    T3F3U2_ZoneManagement
    T3F3U4_ShelfManagement
    T3F3U3_ShelfMaintenance
    T3F2U3_PortManagement
    T3F2U1_PortMaintenance
    T3F2U2_PortDirectionChange
    T4F1U1_SendTransferCommand
    T4F1U2_DeleteTransferCommand
    T4F1U5_TransferCommandPriorityChange
    T4F1U6_TransferCommandCompleted
    T5F1U2_WorkOrderExecution
    T5F1U5_CarrierOutForPicking
    T5F1U9_WorkOrderClose
    T5F1U7_CarrierInForReceipt

    T10F1R4_CarrierMoveComplete
    T10F2R1_DeviceStatusChangeReport
    T10F2R2_DeviceAlarmReport
    T10F2R3_DeviceInventoryReport
    T10F2R4_ZoneSettingChange
    T10F3R1_CarrierWaitIn
    T10F3R2_CarrierMoveStart
    T10F3R3_CarrierMovePosition
    T7F1U2_POAccounting
    T3F4R2_DeviceAlarmReport
    T3F5R1LineStatusChangeReport
    T3F5R2_LineInfoReport
    T3F5R3_LineInProductionInfoReport
    T3F5R4LineInProductionInfoReset


    T10F2S1_StocktakingReport

    T5F1S1WOClose
    T5F1U90_WOExcuting
    '~~~~~~~~~~~~~~~~~~~~~~~~~~GUI~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

    T3F5U1_MaintenanceSet
    T3F5U2_Maintenance
    T3F5U3_LineBigDataAlarmSet
    T3F5U4_ProductionCountSet
    T3F5U5_ClassProductionSet
    T5F1U11_POExecution
    T11F1U11_ProducePOExecution
    T6F5U1_ItemLabelManagement
    T6F5U2_ItemLabelPrint
    T11F1U1_PODownload
    T11F1U2_POExecution
    '~~~~~~~~~~~~~~~~~~~~~~~~~~ERP~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    PO_Release

    StockTaking_Report
    Test

#Region "BOM"
    TEST_HOST_T_COMMAND
    MQ_PRIMARY
    MQ_SECONDARY
#End Region
  End Enum
  Public Shared Function CreateForm(ByVal MessageName As enuMessageName) As FrmMessageTest
    Try
      '當該MessageName的畫面被開啟時把該MessageName記錄起來，防止被重覆開啟
      If dicFormType.ContainsKey(MessageName) Then
        Dim newForm As FrmMessageTest = CType(dicFormType.Item(MessageName), FrmMessageTest)
        newForm.Focus()
        Return newForm
      Else
        Dim newForm As FrmMessageTest = New FrmMessageTest(MessageName)
        dicFormType.Add(MessageName, newForm)
        Return newForm
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Private Sub FrmMessageTest_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
    Try
      If dicFormType.ContainsKey(MessageName) Then
        dicFormType.Remove(MessageName)
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
End Class
