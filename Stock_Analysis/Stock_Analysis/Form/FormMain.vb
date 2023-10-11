Imports System.Threading
Imports System.Text
Imports eCA_TransactionMessage
Imports eCA_HostObject
Imports System.Net
Imports System.IO
Imports System.Net.Http

''' <summary>
''' 20181117
''' V1.0.0
''' Mark
''' 
''' </summary>
Public Class FormMain
  Friend WithEvents NotifyIcon As System.Windows.Forms.NotifyIcon

  Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    Try
      With gMain
        .LogTool.trcListView = lswtrc
        .LogTool.ViewColor = True
        .LogTool.InitialEnd = True
      End With
      Me.Text = String.Format("eHOST (V{0}) ", VersionInformation.Version)
      TSCBViewLogLevel.SelectedIndex = gMain.LogTool.ViewLV - 1


      NotifyIcon = New System.Windows.Forms.NotifyIcon()
      'NotifyIcon1
      '
      Me.NotifyIcon.Icon = CType(My.Resources.ResourceManager.GetObject("eHOST"), System.Drawing.Icon)
      Me.NotifyIcon.Text = String.Format("eHOST (V{0}) ", VersionInformation.Version)
      Me.NotifyIcon.Visible = True
      'NotifyIconWatchDog.ContextMenuStrip = MainMenuContextMenuStrip


      Me.WindowState = FormWindowState.Minimized
      Me.ShowInTaskbar = False

      SendMessageToLog("[HOST Start]", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      'Dim _SuccessOpenWebService = WCFHost() '-開啟WebService
      'WCFHostOpen() '-開啟WebService

      If gMain.RabbitMQ IsNot Nothing Then
        gMain.RabbitMQ.I_Init_RabbitMQService()
      End If

#Region "HostHandler收WS系統的訊息"
      If (ModuleDeclaration.WMSToHandlingInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB Then
        '接收WMS的Message
        ThreadReceiveWMSMessage = New Thread(New ThreadStart(AddressOf O_thr_WMSDBHandling))                                 'Vito19b18
        ThreadReceiveWMSMessage.IsBackground = True                                                                       'Vito19b18
        ThreadReceiveWMSMessage.Start()                                                                                   'Vito19b18
      ElseIf (ModuleDeclaration.WMSToHandlingInterfaceType And enuHandlingInterfaceType.MQ) = enuHandlingInterfaceType.MQ Then
        ThreadReceiveWMSMessage = New Thread(New ThreadStart(AddressOf O_thr_WMSMQHandling))
        ThreadReceiveWMSMessage.IsBackground = True
        ThreadReceiveWMSMessage.Start()
      ElseIf (ModuleDeclaration.WMSToHandlingInterfaceType And enuHandlingInterfaceType.WebAPI) = enuHandlingInterfaceType.WebAPI Then
        ThreadReceiveWMSMessage = New Thread(New ThreadStart(AddressOf Create_HttpListener_Noneblocking_WMS))
        ThreadReceiveWMSMessage.IsBackground = True
        ThreadReceiveWMSMessage.Start()
      End If
      ThreadReceiveWMSMessage = New Thread(New ThreadStart(AddressOf Create_HttpListener_Noneblocking_WMS))
      ThreadReceiveWMSMessage.IsBackground = True
      ThreadReceiveWMSMessage.Start()
#End Region

      '改用 O_thr_ToOtherDBHandling_Result 回覆其他系統的Message 因為回覆其他系統都是讀取HOST_T_COMMAND
      'If (ModuleDeclaration.WMSToHandlingInterfaceType_Result And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB Then 'Vito19b18
      '  '回覆WMS的Message 
      '  ThreadTransmitWMSMessage = New Thread(New ThreadStart(AddressOf thrWMSDBHandling_Result))                                 'Vito19b18
      '  ThreadTransmitWMSMessage.IsBackground = True                                                                              'Vito19b18
      '  ThreadTransmitWMSMessage.Start()                                                                                          'Vito19b18
      'End If   

#Region "HostHandler收GUI系統的訊息"
      If (ModuleDeclaration.GUIToHandlingInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB Then
        '接收GUI的Message
        ThreadReceiveGUIMessage = New Thread(New ThreadStart(AddressOf O_thr_GUIDBHandling))
        ThreadReceiveGUIMessage.IsBackground = True
        ThreadReceiveGUIMessage.Start()
      ElseIf (ModuleDeclaration.GUIToHandlingInterfaceType And enuHandlingInterfaceType.MQ) = enuHandlingInterfaceType.MQ Then
        ThreadReceiveGUIMessage = New Thread(New ThreadStart(AddressOf O_thr_GUIMQHandling))
        ThreadReceiveGUIMessage.IsBackground = True
        ThreadReceiveGUIMessage.Start()
      ElseIf (ModuleDeclaration.GUIToHandlingInterfaceType And enuHandlingInterfaceType.WebAPI) = enuHandlingInterfaceType.WebAPI Then
        ThreadReceiveGUIMessage = New Thread(New ThreadStart(AddressOf Create_HttpListener_Noneblocking_GUI))
        ThreadReceiveGUIMessage.IsBackground = True
        ThreadReceiveGUIMessage.Start()
      End If
#End Region

#Region "HostHandler收MCS系統的訊息"
      If (ModuleDeclaration.MCSToHandlingInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB Then
        '接收MCS的Message
        ThreadReceiveMCSMessage = New Thread(New ThreadStart(AddressOf O_thr_MCSDBHandling))
        ThreadReceiveMCSMessage.IsBackground = True
        'Vito_20421 ThreadReceiveMCSMessage.Start()
      ElseIf (ModuleDeclaration.MCSToHandlingInterfaceType And enuHandlingInterfaceType.MQ) = enuHandlingInterfaceType.MQ Then
        ThreadReceiveMCSMessage = New Thread(New ThreadStart(AddressOf O_thr_MCSMQHandling))
        ThreadReceiveMCSMessage.IsBackground = True
        ThreadReceiveMCSMessage.Start()
      ElseIf (ModuleDeclaration.MCSToHandlingInterfaceType And enuHandlingInterfaceType.WebAPI) = enuHandlingInterfaceType.WebAPI Then
        ThreadReceiveMCSMessage = New Thread(New ThreadStart(AddressOf Create_HttpListener_Noneblocking_MCS))
        ThreadReceiveMCSMessage.IsBackground = True
        ThreadReceiveMCSMessage.Start()
      End If
#End Region

#Region "HostHandler收NS系統的訊息"
      If (ModuleDeclaration.NSToHandlingInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB Then
        '接收MCS的Message
        ThreadReceiveNSMessage = New Thread(New ThreadStart(AddressOf O_thr_NSDBHandling))
        ThreadReceiveNSMessage.IsBackground = True
        ThreadReceiveNSMessage.Start()
      ElseIf (ModuleDeclaration.NSToHandlingInterfaceType And enuHandlingInterfaceType.MQ) = enuHandlingInterfaceType.MQ Then
        ThreadReceiveNSMessage = New Thread(New ThreadStart(AddressOf O_thr_NSMQHandling))
        ThreadReceiveNSMessage.IsBackground = True
        ThreadReceiveNSMessage.Start()
      ElseIf (ModuleDeclaration.NSToHandlingInterfaceType And enuHandlingInterfaceType.WebAPI) = enuHandlingInterfaceType.WebAPI Then
        ThreadReceiveNSMessage = New Thread(New ThreadStart(AddressOf Create_HttpListener_Noneblocking_NS))
        ThreadReceiveNSMessage.IsBackground = True
        ThreadReceiveNSMessage.Start()
      End If
#End Region

#Region "HostHandler送給其他內部系統"
      '三種系統有其中一種用DB通訊，就要開啟THREAD去讀取HOST_T_COMMAND處理
      If (ModuleDeclaration.HandlingToWMSInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB OrElse
         (ModuleDeclaration.HandlingToGUIInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB OrElse
         (ModuleDeclaration.HandlingToMCSInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB OrElse
         (ModuleDeclaration.HandlingToNSInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB Then
        ThreadTransmitWMSMessage = New Thread(New ThreadStart(AddressOf O_thr_ToOtherDBHandling_Result))
        ThreadTransmitWMSMessage.IsBackground = True
        ThreadTransmitWMSMessage.Start()
      End If

      If (ModuleDeclaration.HandlingToWMSInterfaceType And enuHandlingInterfaceType.MQ) = enuHandlingInterfaceType.MQ Then
        ThreadReceiveMCSMessage = New Thread(New ThreadStart(AddressOf O_thr_ToWMSMQHandling_Result))
        ThreadReceiveMCSMessage.IsBackground = True
        ThreadReceiveMCSMessage.Start()
      End If

      If (ModuleDeclaration.HandlingToGUIInterfaceType And enuHandlingInterfaceType.MQ) = enuHandlingInterfaceType.MQ Then
        ThreadReceiveMCSMessage = New Thread(New ThreadStart(AddressOf O_thr_ToGUIMQHandling_Result))
        ThreadReceiveMCSMessage.IsBackground = True
        ThreadReceiveMCSMessage.Start()
      End If

      If (ModuleDeclaration.HandlingToMCSInterfaceType And enuHandlingInterfaceType.MQ) = enuHandlingInterfaceType.MQ Then
        ThreadReceiveMCSMessage = New Thread(New ThreadStart(AddressOf O_thr_ToMCSMQHandling_Result))
        ThreadReceiveMCSMessage.IsBackground = True
        ThreadReceiveMCSMessage.Start()
      End If
#End Region

#Region "HostHandler收外部系統"
      If (ModuleDeclaration.HostToHandlingInterfaceType And enuHandlingInterfaceType.DB) = enuHandlingInterfaceType.DB Then
        'DB的部分放在AUTO EXCUTE做
        f_AutoDownLoad = True
      End If

      If (ModuleDeclaration.HostToHandlingInterfaceType And enuHandlingInterfaceType.WebService) = enuHandlingInterfaceType.WebService Then
        '對外部系統會有自己的多個WCF接口
      End If

      If (ModuleDeclaration.HostToHandlingInterfaceType And enuHandlingInterfaceType.WebAPI) = enuHandlingInterfaceType.WebAPI Then
        gHttpListenerThread = New Thread(New ThreadStart(AddressOf Create_HttpListener))
        gHttpListenerThread.IsBackground = True
        gHttpListenerThread.Start()
      End If
#End Region

      '如果有系統是使用WebService和WMS進行對接，就必須把WMS的WCF Service開啟
      If (ModuleDeclaration.GUIToHandlingInterfaceType And enuHandlingInterfaceType.WebService) = enuHandlingInterfaceType.WebService OrElse
         (ModuleDeclaration.MCSToHandlingInterfaceType And enuHandlingInterfaceType.WebService) = enuHandlingInterfaceType.WebService OrElse
         (ModuleDeclaration.WMSToHandlingInterfaceType And enuHandlingInterfaceType.WebService) = enuHandlingInterfaceType.WebService OrElse
         (ModuleDeclaration.NSToHandlingInterfaceType And enuHandlingInterfaceType.WebService) = enuHandlingInterfaceType.WebService Then
        WCFHostOpen()
      End If

      '处理自动执行的Thread
      'ThreadAutoExcute = New Thread(New ThreadStart(AddressOf O_thr_Auto_Posting))
      'ThreadAutoExcute.IsBackground = True
      'ThreadAutoExcute.Start()

      '处理自动执行的Thread
      ThreadAutoCheck = New Thread(New ThreadStart(AddressOf O_thr_Auto_Excute))
      ThreadAutoCheck.IsBackground = True
      ThreadAutoCheck.Start()

      '處理下載主檔的Thread
      ThreadDownLoadMainFile = New Thread(New ThreadStart(AddressOf O_thr_DownLoad_MainFile))
      ThreadDownLoadMainFile.IsBackground = True
      ThreadDownLoadMainFile.Start()

      '处理自动执行的Thread
      ThreadWriteQueueLog = New Thread(New ThreadStart(AddressOf Common_DBManagement.O_thr_Write_QueueLog))
      ThreadWriteQueueLog.IsBackground = True
      'Vito_20421 ThreadWriteQueueLog.Start()

      ''
      'ThreadAutoReport = New Thread(New ThreadStart(AddressOf O_thr_Auto_Report))
      'ThreadAutoReport.IsBackground = True
      'ThreadAutoReport.Start()


    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  'HeartBeat
  Private Sub mHBTimer_Tick() Handles HBTimer.Tick
    Dim ret_Msg As String = ""                                                                  'Vito_12b30
    Try
      'mHB.UpdateHeartBeat()
      'If HeartBeat.SetCurrentTime(ModuleHelpFunc.GetNewTime_DBFormat(), ret_Msg) = False Then   'Vito_12b30
      '  SendMessageToLog(ret_Msg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)                       'Vito_12b30
      'End If                                                                                    'Vito_12b30
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '雙擊Log，跳出視窗顯示該行Log的資料
  Private Sub lswtrc_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lswtrc.DoubleClick
    Try
      FormInfo.TextBox1.Text = ""
      FormInfo.TextBox1.Text = lswtrc.SelectedItems(0).Text
      FormInfo.ShowDialog()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub RefreshDBToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshDBToolStripMenuItem1.Click
    Try
      Refreshflag = True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Private Sub NotifyIcon_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon.MouseDoubleClick
    Try
      Me.ShowInTaskbar = True
      Me.Show()
      Me.WindowState = FormWindowState.Normal
      NotifyIcon.Visible = False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    Try
      If Me.WindowState = FormWindowState.Minimized Then
        Me.Hide()
        NotifyIcon.Visible = True
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Private Sub TSCBViewLogLevel_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles TSCBViewLogLevel.SelectedIndexChanged
    Try
      gMain.LogTool.ViewLV = TSCBViewLogLevel.SelectedIndex + 1
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Private Sub LCSInitialToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LCSInitialToolStripMenuItem.Click
    '  OnlineWithLCS()
  End Sub
  Private Sub ForEcatchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ForEcatchToolStripMenuItem.Click
    Dim _test As New Form1
    _test.Show()
  End Sub

  Private Sub JJGtestToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ' Dim _test As New FrmForJJG
    'FrmForJJG.Show()
  End Sub
  Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_Stop.CheckedChanged
    Try
      If CheckBox_Stop.Checked() Then
        Handling_CycleStop = 1
      Else
        Handling_CycleStop = 0
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub






  Private Sub T3F5R1LineStatusChangeReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5R1LineStatusChangeReportToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5R1LineStatusChangeReport)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F5R2LineInfoReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5R2LineInfoReportToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5R2_LineInfoReport)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F5R3LineInProductionInfoReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5R3LineInProductionInfoReportToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5R3_LineInProductionInfoReport)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F5R4LineInProductionInfoResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5R4LineInProductionInfoResetToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5R4LineInProductionInfoReset)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F4R2DeviceAlarmReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F4R2DeviceAlarmReportToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F4R2_DeviceAlarmReport)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F5U1MaintenanceSetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5U1MaintenanceSetToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5U1_MaintenanceSet)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F5U2MaintenanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5U2MaintenanceToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5U2_Maintenance)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F5U3LineBigDataAlarmSetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5U3LineBigDataAlarmSetToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5U3_LineBigDataAlarmSet)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F5U4ProductionCountSetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5U4ProductionCountSetToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5U4_ProductionCountSet)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T3F5U5ClassProductionSetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T3F5U5ClassProductionSetToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T3F5U5_ClassProductionSet)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T11F1U11ProducePOExecutionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T11F1U11ProducePOExecutionToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T11F1U11_ProducePOExecution)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '環鴻 用from to 日期提單
  'Private Sub ASRSstockCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASRSstockCheckToolStripMenuItem.Click
  '  Dim FromData = InputBox("ASRS_stockCheck", "FromData", "YYYYMMDD")
  '  Dim ToData = InputBox("ASRS_stockCheck", "ToData", "YYYYMMDD")
  '  Dim ret_msg = ""
  '  If Mod_WCFHost.ASRS_stockCheck(FromData, ToData, ret_msg) = False Then
  '    MsgBox(ret_msg)
  '  End If

  'End Sub

  'Private Sub ASRSsingleMatDocToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASRSsingleMatDocToolStripMenuItem.Click
  '  Dim PO_ID = InputBox("ASRS_singleMatDoc", "PO_ID", "14碼")
  '  'Dim Download_TYPE = InputBox("ASRS_singleMatDoc", "Download_TYPE", "InBound = 0, OutBound = 1,  Transcation = 2,  Stocktaking = 3")
  '  Dim ret_msg = ""
  '  If Mod_WCFHost.ASRS_singleMatDoc(PO_ID, ret_msg, "") = False Then
  '    MsgBox(ret_msg)
  '  End If
  'End Sub

  'Private Sub ASRSupdateBINToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASRSupdateBINToolStripMenuItem.Click
  '  Dim WERKS = InputBox("ASRS_update_BIN", "WERKS (OWNER)", "")
  '  Dim MATNR = InputBox("ASRS_update_BIN", "MATNR (SKU_NO)", "")
  '  Dim CHARG = InputBox("ASRS_update_BIN", "CHARG (BATCH)", "")
  '  Dim LGORT = InputBox("ASRS_update_BIN", "LGORT (SubOWNER)", "")

  '  Dim Result = False
  '  If Mod_WCFHost.ASRS_update_BIN(WERKS, MATNR, CHARG, LGORT, Result) = False Then

  '  End If

  '  'Dim str = ""
  '  'Dim clsresponse = Tools.ParseXmlStringToClass(Of MSG_ASRS_update_BIN)(str)

  'End Sub

  Private Sub T10F4U1MainFileImportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T10F4U1MainFileImportToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T10F4U1_MainFileImport)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T6F5U1ItemLabelManagementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T6F5U1_ItemLabelManagement)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T6F5U2ItemLabelPrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T6F5U2ItemLabelPrintToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T6F5U2_ItemLabelPrint)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub



  Private Sub T7F1U2POAccountingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T7F1U2POAccountingToolStripMenuItem.Click

  End Sub

  Private Sub T10F2S1StocktakingReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T10F2S1StocktakingReportToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T10F2S1_StocktakingReport)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '  Private Sub 採購單回報ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 採購單回報ToolStripMenuItem.Click
  '    Dim Input As String = ""
  '    Dim ret_strResultMsg As String = ""
  '    Input = InputBox("輸入PO_ID", "單據回報")
  '    If Input = "" Then
  '      Exit Sub
  '    End If
  '    Dim PO_ID As String = Input.Trim
  '    Dim dicPO_id As New Dictionary(Of String, String)
  '    dicPO_id.Add(PO_ID, PO_ID)
  '    Dim dicPO As New Dictionary(Of String, clsPO)
  '    Dim dicPO_Line As New Dictionary(Of String, clsPO_LINE)
  '    Dim dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
  '    Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)
  '    If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(dicPO_id, dicPO) = False Then
  '      MsgBox("GET WMS_T_PO Fail")
  '      Exit Sub
  '    End If
  '    If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(dicPO_id, dicPO_Line) = False Then
  '      MsgBox("GET WMS_T_PO_LINE Fail")
  '    End If
  '    If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(dicPO_id, dicPO_DTL) = False Then
  '      MsgBox("GET WMS_T_PO_DTL Fail")
  '    End If
  '    If gMain.objHandling.O_GetDB_dicPO_MergeBydicPO_ID(dicPO_id, dicPO_Merge) = False Then
  '      MsgBox("GET WMS_T_PO_Merge Fail")
  '    End If
  '    Dim StrXML As String = ""
  '    For Each objPO In dicPO.Values
  '      Dim PO_TYPE1 = objPO.PO_Type1
  '      Dim PO_TYPE2 = objPO.PO_Type2

  '      '回報單據

  '      If CLng(PO_TYPE1) = enuPOType_1.Combination_in Then
  '        Select Case PO_TYPE2
  '#Region "採購單"
  '          Case "101"
  '            If SendBuyData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '#Region "採購入庫單"
  '          Case "102"  '採購入庫單
  '            If InboundData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '#Region "生產入庫單"
  '          Case "103" '生產入庫單
  '            If ProduceInData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '#Region "雜收單"
  '          Case "104" '雜收單
  '            If OtherInData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '#Region "調撥入庫單"
  '          Case "144" '調撥入庫單
  '            Dim dicInbound_DTL As New Dictionary(Of String, clsINBOUND_DTL)
  '            Dim dicOutbound_DTL As New Dictionary(Of String, clsOUTBOUND_DTL)
  '            If gMain.objHandling.O_GetDB_dicInboundDTLByPOID(PO_ID, dicInbound_DTL) = False Then
  '              ret_strResultMsg = String.Format("Get WMS_T_INBOUND_DTL Failed PO_ID={0}", PO_ID)
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '            If TransferData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg, dicInbound_DTL, dicOutbound_DTL) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '#Region "銷退單"
  '          Case enuPOType_2.SellReturn '調撥入庫單
  '            If SellReturnData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '        End Select
  '      ElseIf CLng(PO_TYPE1) = enuPOType_1.Picking_out Then
  '        Select Case PO_TYPE2
  '#Region "雜發單"
  '          Case "303" '雜發單
  '            If OtherOutData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '#Region "銷貨單"
  '          Case "304" '銷貨單
  '            If SellData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '#Region "調撥出庫"
  '          Case "344" '調撥出庫
  '            Dim dicOutbound_DTL As New Dictionary(Of String, clsOUTBOUND_DTL)
  '            Dim dicInbound_DTL As New Dictionary(Of String, clsINBOUND_DTL)
  '            If gMain.objHandling.O_GetDB_dicOutboundDTLByPOID(objPO.PO_ID, dicOutbound_DTL) = False Then
  '              ret_strResultMsg = String.Format("Get WMS_T_OUTBOUND_DTL Failed PO_ID={0}", objPO.PO_ID)
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If

  '            If TransferData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg, dicInbound_DTL, dicOutbound_DTL) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '#Region "退供應商"
  '          Case enuPOType_2.InboundReturn_Data '退供應商
  '            If InboundReturnData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '        End Select
  '      ElseIf CLng(PO_TYPE1) = enuPOType_1.Transaction Then
  '        Select Case PO_TYPE2
  '#Region "貨主調撥"
  '          Case "631" '貨主調撥
  '            If TransferOwnerData_Report(objPO, dicPO_DTL, dicPO_Merge, StrXML, ret_strResultMsg) = False Then
  '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '              Exit Sub
  '            End If
  '#End Region
  '        End Select
  '      End If
  '    Next
  '    Dim Result As String = ""
  '    If STD_IN(StrXML, Result, ret_strResultMsg) = False Then
  '      If ret_strResultMsg = "" Then
  '        ret_strResultMsg = "回報ERP失敗"
  '      End If
  '      SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  '      Mod_Command_Record.O_Handle_Host_To_HS_Command_Hist("T5F1S1_WOClose", enuConnectionType.WebService, enuSystemType.ERP, "", GetNewTime_DBFormat, StrXML, "1", Result, "", "", "", ret_strResultMsg)
  '      MsgBox(ret_strResultMsg)

  '    Else
  '      Mod_Command_Record.O_Handle_Host_To_HS_Command_Hist("T5F1S1_WOClose", enuConnectionType.WebService, enuSystemType.ERP, "", GetNewTime_DBFormat, StrXML, "0", Result, "", "", "", ret_strResultMsg)
  '      MsgBox(PO_ID & "回報成功")
  '    End If

  '  End Sub

  Private Sub T5F1U11POExecutionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T5F1U11POExecutionToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T5F1U11_POExecution)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T11F1U2POExecutionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T11F1U2POExecutionToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T11F1U2_POExecution)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T11F1U14SwitchOnLocationLightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T11F1U14SwitchOnLocationLightToolStripMenuItem.Click
    Try
      Dim dicLocation_No As New Dictionary(Of String, String)
      Dim Location_No = ""
      Dim ret_strResultMsg As String = ""
      Dim Host_Command As New Dictionary(Of String, clsFromHostCommand)
      Dim lstSql As New List(Of String)
      Location_No = InputBox("輸入要點亮的儲位編號")
      If Location_No = "" Then
        Return
      End If
      dicLocation_No.Add(Location_No, Location_No)
      If Send_T11F1U14_SwitchOnLocationLight_to_WMS(ret_strResultMsg, Host_Command, dicLocation_No) = False Then
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return
      End If

      For Each obj As clsFromHostCommand In Host_Command.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert Host_Command Info SQL Failed"
          Return
        End If
      Next
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        ret_strResultMsg = "WMS Update DB Failed"
        Return
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T5F1S1WOCloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T5F1S1WOCloseToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T5F1S1WOClose)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub 盤點單回報ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 盤點單回報ToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.StockTaking_Report)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub


  Private Sub 單據放行ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 單據放行ToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.PO_Release)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub 測試ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 測試ToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.Test)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub HOSTTCOMMAND測試ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HOSTTCOMMAND測試ToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.TEST_HOST_T_COMMAND)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub MQTESTPRIMARYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MQTESTPRIMARYToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.MQ_PRIMARY)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub MQTESTSECONDARYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MQTESTSECONDARYToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.MQ_SECONDARY)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T11F1U1PODownloadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T11F1U1PODownloadToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T11F1U1_PODownload)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T5F1U90WOExcutingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T5F1U90WOExcutingToolStripMenuItem.Click
    Try
      Dim _form = FrmMessageTest.CreateForm(FrmMessageTest.enuMessageName.T5F1U90_WOExcuting)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub 讀取ERP中介檔ToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles 讀取ERP中介檔ToolStripMenuItem.Click
    Try
      '單據主檔 每一分鐘處理一次
      Dim strResultMsg As String = ""
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
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub 讀取料品中介ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 讀取料品中介ToolStripMenuItem.Click
    Dim strResultMsg As String = ""
    Dim dicINVXB As New Dictionary(Of String, clsINVXB)
    If gMain.objHandling.O_GetDB_dicINVXBByXB008_IS_ZERO(dicINVXB) = True Then

      Module_SKUManagement_INVXB.O_SKUManagement(dicINVXB, strResultMsg)
    End If
  End Sub

  Private Sub 特定料品下載ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 特定料品下載ToolStripMenuItem.Click
    Dim strResultMsg As String = ""
    Dim dicINVXB As New Dictionary(Of String, clsINVXB)
    Dim SKU_NO = InputBox("輸入料號")
    If gMain.objHandling.O_GetDB_dicINVXBByXB008_IS_ZEROBY_SKU(SKU_NO, dicINVXB) = True Then

      Module_SKUManagement_INVXB.O_SKUManagement(dicINVXB, strResultMsg)
    End If
  End Sub

  Private Sub 基本資料ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 基本資料ToolStripMenuItem.Click
    Using httpClient As New HttpClient()
      ' 设置要访问的URL
      Dim url As String = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_peratio_analysis"

      ' 发送GET请求并获取响应
      Dim response As HttpResponseMessage = httpClient.GetAsync(url).Result

      ' 检查响应是否成功
      If response.IsSuccessStatusCode Then
        ' 读取响应内容
        Dim responseContent As String = response.Content.ReadAsStringAsync().Result

        Dim data As List(Of MSG_Peratio) = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of MSG_Peratio))(responseContent)

        ' 输出响应内容
        Console.WriteLine("响应内容:")
        Console.WriteLine(responseContent)
        SendMessageToLog("[Request]：" & responseContent, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      Else
        ' 处理请求失败的情况
        Console.WriteLine("请求失败. HTTP状态码: " & response.StatusCode)
      End If
    End Using
  End Sub







  'Private Sub 對帳ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 對帳ToolStripMenuItem.Click
  '  Dim str_log = ""
  '  If Module_T11F2U1_InventoryComparison.O_Process(Nothing, str_log, "") = False Then
  '    MsgBox(str_log)
  '  End If
  'End Sub

  'Private Sub 回報入庫短缺ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 回報入庫短缺ToolStripMenuItem.Click
  '  Try
  '    Module_Auto_Excute.I_Auto_Report_Inbound_Lack()
  '  Catch ex As Exception

  '  End Try
  'End Sub
End Class
