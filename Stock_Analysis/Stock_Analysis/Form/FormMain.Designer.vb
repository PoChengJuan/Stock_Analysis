<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMain))
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lswtrc = New System.Windows.Forms.ListView()
        Me.WMSMenuStrip = New System.Windows.Forms.MenuStrip()
        Me.RefreshDBToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RefreshDBToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.LCSIntitialToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LCSInitialToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ForEcatchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogToolToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChangeLogLevelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSCBViewLogLevel = New System.Windows.Forms.ToolStripComboBox()
        Me.SendMessageToMESToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.採購單回報ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.盤點單回報ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.單據放行ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.測試ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HostHandlerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FromWMSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T7F1U2POAccountingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T10F2S1StocktakingReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T5F1S1WOCloseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T5F1U90WOExcutingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FromMCSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F4R2DeviceAlarmReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5R1LineStatusChangeReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5R2LineInfoReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5R3LineInProductionInfoReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5R4LineInProductionInfoResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FromGUIToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5U1MaintenanceSetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5U2MaintenanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5U3LineBigDataAlarmSetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5U4ProductionCountSetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T3F5U5ClassProductionSetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T5F1U11POExecutionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T6F5U2ItemLabelPrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T11F1U11ProducePOExecutionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T10F4U1MainFileImportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T11F1U1PODownloadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T11F1U2POExecutionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToWMSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.T11F1U14SwitchOnLocationLightToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TESTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HOSTTCOMMAND測試ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MQTESTPRIMARYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MQTESTSECONDARYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ERP單據測試ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.讀取ERP中介檔ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.讀取料品中介ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.特定料品下載ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HBTimer = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckBox_Stop = New System.Windows.Forms.CheckBox()
        Me.功能ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.基本資料ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.WMSMenuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(18, 53)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1206, 539)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lswtrc)
        Me.TabPage1.Location = New System.Drawing.Point(4, 28)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabPage1.Size = New System.Drawing.Size(1198, 507)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "TraceLog"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lswtrc
        '
        Me.lswtrc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lswtrc.HideSelection = False
        Me.lswtrc.Location = New System.Drawing.Point(4, 5)
        Me.lswtrc.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lswtrc.Name = "lswtrc"
        Me.lswtrc.Size = New System.Drawing.Size(1190, 497)
        Me.lswtrc.TabIndex = 0
        Me.lswtrc.UseCompatibleStateImageBehavior = False
        '
        'WMSMenuStrip
        '
        Me.WMSMenuStrip.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
        Me.WMSMenuStrip.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.WMSMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RefreshDBToolStripMenuItem, Me.LCSIntitialToolStripMenuItem, Me.LogToolToolStripMenuItem, Me.SendMessageToMESToolStripMenuItem, Me.HostHandlerToolStripMenuItem, Me.TESTToolStripMenuItem, Me.ERP單據測試ToolStripMenuItem, Me.功能ToolStripMenuItem})
        Me.WMSMenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.WMSMenuStrip.Name = "WMSMenuStrip"
        Me.WMSMenuStrip.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.WMSMenuStrip.Size = New System.Drawing.Size(1274, 32)
        Me.WMSMenuStrip.TabIndex = 3
        Me.WMSMenuStrip.Text = "MenuStrip1"
        '
        'RefreshDBToolStripMenuItem
        '
        Me.RefreshDBToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RefreshDBToolStripMenuItem1})
        Me.RefreshDBToolStripMenuItem.Name = "RefreshDBToolStripMenuItem"
        Me.RefreshDBToolStripMenuItem.Size = New System.Drawing.Size(115, 28)
        Me.RefreshDBToolStripMenuItem.Text = "RefreshDB"
        '
        'RefreshDBToolStripMenuItem1
        '
        Me.RefreshDBToolStripMenuItem1.Name = "RefreshDBToolStripMenuItem1"
        Me.RefreshDBToolStripMenuItem1.Size = New System.Drawing.Size(199, 34)
        Me.RefreshDBToolStripMenuItem1.Text = "RefreshDB"
        '
        'LCSIntitialToolStripMenuItem
        '
        Me.LCSIntitialToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LCSInitialToolStripMenuItem, Me.ForEcatchToolStripMenuItem})
        Me.LCSIntitialToolStripMenuItem.Name = "LCSIntitialToolStripMenuItem"
        Me.LCSIntitialToolStripMenuItem.Size = New System.Drawing.Size(81, 28)
        Me.LCSIntitialToolStripMenuItem.Text = "Intitial"
        '
        'LCSInitialToolStripMenuItem
        '
        Me.LCSInitialToolStripMenuItem.Name = "LCSInitialToolStripMenuItem"
        Me.LCSInitialToolStripMenuItem.Size = New System.Drawing.Size(193, 34)
        Me.LCSInitialToolStripMenuItem.Text = "LCSInitial"
        '
        'ForEcatchToolStripMenuItem
        '
        Me.ForEcatchToolStripMenuItem.Name = "ForEcatchToolStripMenuItem"
        Me.ForEcatchToolStripMenuItem.Size = New System.Drawing.Size(193, 34)
        Me.ForEcatchToolStripMenuItem.Text = "ForEcatch"
        '
        'LogToolToolStripMenuItem
        '
        Me.LogToolToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ChangeLogLevelToolStripMenuItem})
        Me.LogToolToolStripMenuItem.Name = "LogToolToolStripMenuItem"
        Me.LogToolToolStripMenuItem.Size = New System.Drawing.Size(94, 28)
        Me.LogToolToolStripMenuItem.Text = "LogTool"
        '
        'ChangeLogLevelToolStripMenuItem
        '
        Me.ChangeLogLevelToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSCBViewLogLevel})
        Me.ChangeLogLevelToolStripMenuItem.Name = "ChangeLogLevelToolStripMenuItem"
        Me.ChangeLogLevelToolStripMenuItem.Size = New System.Drawing.Size(305, 34)
        Me.ChangeLogLevelToolStripMenuItem.Text = "Change View Log Level"
        '
        'TSCBViewLogLevel
        '
        Me.TSCBViewLogLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.TSCBViewLogLevel.Items.AddRange(New Object() {"Error", "WARN", "TRACE", "DEBUG", "ALL"})
        Me.TSCBViewLogLevel.MaxDropDownItems = 5
        Me.TSCBViewLogLevel.Name = "TSCBViewLogLevel"
        Me.TSCBViewLogLevel.Size = New System.Drawing.Size(121, 31)
        '
        'SendMessageToMESToolStripMenuItem
        '
        Me.SendMessageToMESToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.採購單回報ToolStripMenuItem, Me.盤點單回報ToolStripMenuItem, Me.單據放行ToolStripMenuItem, Me.測試ToolStripMenuItem})
        Me.SendMessageToMESToolStripMenuItem.Name = "SendMessageToMESToolStripMenuItem"
        Me.SendMessageToMESToolStripMenuItem.Size = New System.Drawing.Size(102, 28)
        Me.SendMessageToMESToolStripMenuItem.Text = "ERP_Test"
        '
        '採購單回報ToolStripMenuItem
        '
        Me.採購單回報ToolStripMenuItem.Name = "採購單回報ToolStripMenuItem"
        Me.採購單回報ToolStripMenuItem.Size = New System.Drawing.Size(200, 34)
        Me.採購單回報ToolStripMenuItem.Text = "單據回報"
        '
        '盤點單回報ToolStripMenuItem
        '
        Me.盤點單回報ToolStripMenuItem.Name = "盤點單回報ToolStripMenuItem"
        Me.盤點單回報ToolStripMenuItem.Size = New System.Drawing.Size(200, 34)
        Me.盤點單回報ToolStripMenuItem.Text = "盤點單回報"
        '
        '單據放行ToolStripMenuItem
        '
        Me.單據放行ToolStripMenuItem.Name = "單據放行ToolStripMenuItem"
        Me.單據放行ToolStripMenuItem.Size = New System.Drawing.Size(200, 34)
        Me.單據放行ToolStripMenuItem.Text = "單據放行"
        '
        '測試ToolStripMenuItem
        '
        Me.測試ToolStripMenuItem.Name = "測試ToolStripMenuItem"
        Me.測試ToolStripMenuItem.Size = New System.Drawing.Size(200, 34)
        Me.測試ToolStripMenuItem.Text = "測試"
        '
        'HostHandlerToolStripMenuItem
        '
        Me.HostHandlerToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FromWMSToolStripMenuItem, Me.FromMCSToolStripMenuItem, Me.FromGUIToolStripMenuItem, Me.ToWMSToolStripMenuItem})
        Me.HostHandlerToolStripMenuItem.Name = "HostHandlerToolStripMenuItem"
        Me.HostHandlerToolStripMenuItem.Size = New System.Drawing.Size(139, 28)
        Me.HostHandlerToolStripMenuItem.Text = "Host Handler"
        '
        'FromWMSToolStripMenuItem
        '
        Me.FromWMSToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.T7F1U2POAccountingToolStripMenuItem, Me.T10F2S1StocktakingReportToolStripMenuItem, Me.T5F1S1WOCloseToolStripMenuItem, Me.T5F1U90WOExcutingToolStripMenuItem})
        Me.FromWMSToolStripMenuItem.Name = "FromWMSToolStripMenuItem"
        Me.FromWMSToolStripMenuItem.Size = New System.Drawing.Size(204, 34)
        Me.FromWMSToolStripMenuItem.Text = "From WMS"
        '
        'T7F1U2POAccountingToolStripMenuItem
        '
        Me.T7F1U2POAccountingToolStripMenuItem.Name = "T7F1U2POAccountingToolStripMenuItem"
        Me.T7F1U2POAccountingToolStripMenuItem.Size = New System.Drawing.Size(346, 34)
        Me.T7F1U2POAccountingToolStripMenuItem.Text = "T7F1U2_POAccounting"
        '
        'T10F2S1StocktakingReportToolStripMenuItem
        '
        Me.T10F2S1StocktakingReportToolStripMenuItem.Name = "T10F2S1StocktakingReportToolStripMenuItem"
        Me.T10F2S1StocktakingReportToolStripMenuItem.Size = New System.Drawing.Size(346, 34)
        Me.T10F2S1StocktakingReportToolStripMenuItem.Text = "T10F2S1_StocktakingReport"
        '
        'T5F1S1WOCloseToolStripMenuItem
        '
        Me.T5F1S1WOCloseToolStripMenuItem.Name = "T5F1S1WOCloseToolStripMenuItem"
        Me.T5F1S1WOCloseToolStripMenuItem.Size = New System.Drawing.Size(346, 34)
        Me.T5F1S1WOCloseToolStripMenuItem.Text = "T5F1S1_WOClose"
        '
        'T5F1U90WOExcutingToolStripMenuItem
        '
        Me.T5F1U90WOExcutingToolStripMenuItem.Name = "T5F1U90WOExcutingToolStripMenuItem"
        Me.T5F1U90WOExcutingToolStripMenuItem.Size = New System.Drawing.Size(346, 34)
        Me.T5F1U90WOExcutingToolStripMenuItem.Text = "T5F1U90_WOExcuting"
        '
        'FromMCSToolStripMenuItem
        '
        Me.FromMCSToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.T3F4R2DeviceAlarmReportToolStripMenuItem, Me.T3F5R1LineStatusChangeReportToolStripMenuItem, Me.T3F5R2LineInfoReportToolStripMenuItem, Me.T3F5R3LineInProductionInfoReportToolStripMenuItem, Me.T3F5R4LineInProductionInfoResetToolStripMenuItem})
        Me.FromMCSToolStripMenuItem.Name = "FromMCSToolStripMenuItem"
        Me.FromMCSToolStripMenuItem.Size = New System.Drawing.Size(204, 34)
        Me.FromMCSToolStripMenuItem.Text = "From MCS"
        '
        'T3F4R2DeviceAlarmReportToolStripMenuItem
        '
        Me.T3F4R2DeviceAlarmReportToolStripMenuItem.Name = "T3F4R2DeviceAlarmReportToolStripMenuItem"
        Me.T3F4R2DeviceAlarmReportToolStripMenuItem.Size = New System.Drawing.Size(415, 34)
        Me.T3F4R2DeviceAlarmReportToolStripMenuItem.Text = "T3F4R2_DeviceAlarmReport"
        '
        'T3F5R1LineStatusChangeReportToolStripMenuItem
        '
        Me.T3F5R1LineStatusChangeReportToolStripMenuItem.Name = "T3F5R1LineStatusChangeReportToolStripMenuItem"
        Me.T3F5R1LineStatusChangeReportToolStripMenuItem.Size = New System.Drawing.Size(415, 34)
        Me.T3F5R1LineStatusChangeReportToolStripMenuItem.Text = "T3F5R1_LineStatusChangeReport"
        '
        'T3F5R2LineInfoReportToolStripMenuItem
        '
        Me.T3F5R2LineInfoReportToolStripMenuItem.Name = "T3F5R2LineInfoReportToolStripMenuItem"
        Me.T3F5R2LineInfoReportToolStripMenuItem.Size = New System.Drawing.Size(415, 34)
        Me.T3F5R2LineInfoReportToolStripMenuItem.Text = "T3F5R2_LineInfoReport"
        '
        'T3F5R3LineInProductionInfoReportToolStripMenuItem
        '
        Me.T3F5R3LineInProductionInfoReportToolStripMenuItem.Name = "T3F5R3LineInProductionInfoReportToolStripMenuItem"
        Me.T3F5R3LineInProductionInfoReportToolStripMenuItem.Size = New System.Drawing.Size(415, 34)
        Me.T3F5R3LineInProductionInfoReportToolStripMenuItem.Text = "T3F5R3_LineInProductionInfoReport"
        '
        'T3F5R4LineInProductionInfoResetToolStripMenuItem
        '
        Me.T3F5R4LineInProductionInfoResetToolStripMenuItem.Name = "T3F5R4LineInProductionInfoResetToolStripMenuItem"
        Me.T3F5R4LineInProductionInfoResetToolStripMenuItem.Size = New System.Drawing.Size(415, 34)
        Me.T3F5R4LineInProductionInfoResetToolStripMenuItem.Text = "T3F5R4_LineInProductionInfoReset"
        '
        'FromGUIToolStripMenuItem
        '
        Me.FromGUIToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.T3F5U1MaintenanceSetToolStripMenuItem, Me.T3F5U2MaintenanceToolStripMenuItem, Me.T3F5U3LineBigDataAlarmSetToolStripMenuItem, Me.T3F5U4ProductionCountSetToolStripMenuItem, Me.T3F5U5ClassProductionSetToolStripMenuItem, Me.T5F1U11POExecutionToolStripMenuItem, Me.T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem, Me.T6F5U2ItemLabelPrintToolStripMenuItem, Me.T11F1U11ProducePOExecutionToolStripMenuItem, Me.T10F4U1MainFileImportToolStripMenuItem, Me.T11F1U1PODownloadToolStripMenuItem, Me.T11F1U2POExecutionToolStripMenuItem})
        Me.FromGUIToolStripMenuItem.Name = "FromGUIToolStripMenuItem"
        Me.FromGUIToolStripMenuItem.Size = New System.Drawing.Size(204, 34)
        Me.FromGUIToolStripMenuItem.Text = "From GUI"
        '
        'T3F5U1MaintenanceSetToolStripMenuItem
        '
        Me.T3F5U1MaintenanceSetToolStripMenuItem.Name = "T3F5U1MaintenanceSetToolStripMenuItem"
        Me.T3F5U1MaintenanceSetToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T3F5U1MaintenanceSetToolStripMenuItem.Text = "T3F5U1_MaintenanceSet"
        '
        'T3F5U2MaintenanceToolStripMenuItem
        '
        Me.T3F5U2MaintenanceToolStripMenuItem.Name = "T3F5U2MaintenanceToolStripMenuItem"
        Me.T3F5U2MaintenanceToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T3F5U2MaintenanceToolStripMenuItem.Text = "T3F5U2_Maintenance"
        '
        'T3F5U3LineBigDataAlarmSetToolStripMenuItem
        '
        Me.T3F5U3LineBigDataAlarmSetToolStripMenuItem.Name = "T3F5U3LineBigDataAlarmSetToolStripMenuItem"
        Me.T3F5U3LineBigDataAlarmSetToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T3F5U3LineBigDataAlarmSetToolStripMenuItem.Text = "T3F5U3_LineBigDataAlarmSet"
        '
        'T3F5U4ProductionCountSetToolStripMenuItem
        '
        Me.T3F5U4ProductionCountSetToolStripMenuItem.Name = "T3F5U4ProductionCountSetToolStripMenuItem"
        Me.T3F5U4ProductionCountSetToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T3F5U4ProductionCountSetToolStripMenuItem.Text = "T3F5U4_ProductionCountSet"
        '
        'T3F5U5ClassProductionSetToolStripMenuItem
        '
        Me.T3F5U5ClassProductionSetToolStripMenuItem.Name = "T3F5U5ClassProductionSetToolStripMenuItem"
        Me.T3F5U5ClassProductionSetToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T3F5U5ClassProductionSetToolStripMenuItem.Text = "T3F5U5_ClassProductionSet"
        '
        'T5F1U11POExecutionToolStripMenuItem
        '
        Me.T5F1U11POExecutionToolStripMenuItem.Name = "T5F1U11POExecutionToolStripMenuItem"
        Me.T5F1U11POExecutionToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T5F1U11POExecutionToolStripMenuItem.Text = "T5F1U11_POExecution"
        '
        'T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem
        '
        Me.T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem.Name = "T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem"
        Me.T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem.Text = "T6F5U1_ItemLabelManagement"
        '
        'T6F5U2ItemLabelPrintToolStripMenuItem
        '
        Me.T6F5U2ItemLabelPrintToolStripMenuItem.Name = "T6F5U2ItemLabelPrintToolStripMenuItem"
        Me.T6F5U2ItemLabelPrintToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T6F5U2ItemLabelPrintToolStripMenuItem.Text = "T6F5U2_ItemLabelPrint"
        '
        'T11F1U11ProducePOExecutionToolStripMenuItem
        '
        Me.T11F1U11ProducePOExecutionToolStripMenuItem.Name = "T11F1U11ProducePOExecutionToolStripMenuItem"
        Me.T11F1U11ProducePOExecutionToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T11F1U11ProducePOExecutionToolStripMenuItem.Text = "T11F1U11_ProducePOExecution"
        '
        'T10F4U1MainFileImportToolStripMenuItem
        '
        Me.T10F4U1MainFileImportToolStripMenuItem.Name = "T10F4U1MainFileImportToolStripMenuItem"
        Me.T10F4U1MainFileImportToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T10F4U1MainFileImportToolStripMenuItem.Text = "T10F4U1_MainFileImport"
        '
        'T11F1U1PODownloadToolStripMenuItem
        '
        Me.T11F1U1PODownloadToolStripMenuItem.Name = "T11F1U1PODownloadToolStripMenuItem"
        Me.T11F1U1PODownloadToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T11F1U1PODownloadToolStripMenuItem.Text = "T11F1U1_PODownload"
        '
        'T11F1U2POExecutionToolStripMenuItem
        '
        Me.T11F1U2POExecutionToolStripMenuItem.Name = "T11F1U2POExecutionToolStripMenuItem"
        Me.T11F1U2POExecutionToolStripMenuItem.Size = New System.Drawing.Size(379, 34)
        Me.T11F1U2POExecutionToolStripMenuItem.Text = "T11F1U2_POExecution"
        '
        'ToWMSToolStripMenuItem
        '
        Me.ToWMSToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.T11F1U14SwitchOnLocationLightToolStripMenuItem})
        Me.ToWMSToolStripMenuItem.Name = "ToWMSToolStripMenuItem"
        Me.ToWMSToolStripMenuItem.Size = New System.Drawing.Size(204, 34)
        Me.ToWMSToolStripMenuItem.Text = "To WMS"
        '
        'T11F1U14SwitchOnLocationLightToolStripMenuItem
        '
        Me.T11F1U14SwitchOnLocationLightToolStripMenuItem.Name = "T11F1U14SwitchOnLocationLightToolStripMenuItem"
        Me.T11F1U14SwitchOnLocationLightToolStripMenuItem.Size = New System.Drawing.Size(398, 34)
        Me.T11F1U14SwitchOnLocationLightToolStripMenuItem.Text = "T11F1U14_SwitchOnLocationLight"
        '
        'TESTToolStripMenuItem
        '
        Me.TESTToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HOSTTCOMMAND測試ToolStripMenuItem, Me.MQTESTPRIMARYToolStripMenuItem, Me.MQTESTSECONDARYToolStripMenuItem})
        Me.TESTToolStripMenuItem.Name = "TESTToolStripMenuItem"
        Me.TESTToolStripMenuItem.Size = New System.Drawing.Size(66, 28)
        Me.TESTToolStripMenuItem.Text = "TEST"
        '
        'HOSTTCOMMAND測試ToolStripMenuItem
        '
        Me.HOSTTCOMMAND測試ToolStripMenuItem.Name = "HOSTTCOMMAND測試ToolStripMenuItem"
        Me.HOSTTCOMMAND測試ToolStripMenuItem.Size = New System.Drawing.Size(323, 34)
        Me.HOSTTCOMMAND測試ToolStripMenuItem.Text = "HOST_T_COMMAND測試"
        '
        'MQTESTPRIMARYToolStripMenuItem
        '
        Me.MQTESTPRIMARYToolStripMenuItem.Name = "MQTESTPRIMARYToolStripMenuItem"
        Me.MQTESTPRIMARYToolStripMenuItem.Size = New System.Drawing.Size(323, 34)
        Me.MQTESTPRIMARYToolStripMenuItem.Text = "MQ_TEST_PRIMARY"
        '
        'MQTESTSECONDARYToolStripMenuItem
        '
        Me.MQTESTSECONDARYToolStripMenuItem.Name = "MQTESTSECONDARYToolStripMenuItem"
        Me.MQTESTSECONDARYToolStripMenuItem.Size = New System.Drawing.Size(323, 34)
        Me.MQTESTSECONDARYToolStripMenuItem.Text = "MQ_TEST_SECONDARY"
        '
        'ERP單據測試ToolStripMenuItem
        '
        Me.ERP單據測試ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.讀取ERP中介檔ToolStripMenuItem, Me.讀取料品中介ToolStripMenuItem, Me.特定料品下載ToolStripMenuItem})
        Me.ERP單據測試ToolStripMenuItem.Name = "ERP單據測試ToolStripMenuItem"
        Me.ERP單據測試ToolStripMenuItem.Size = New System.Drawing.Size(131, 28)
        Me.ERP單據測試ToolStripMenuItem.Text = "ERP單據讀取"
        '
        '讀取ERP中介檔ToolStripMenuItem
        '
        Me.讀取ERP中介檔ToolStripMenuItem.Name = "讀取ERP中介檔ToolStripMenuItem"
        Me.讀取ERP中介檔ToolStripMenuItem.Size = New System.Drawing.Size(236, 34)
        Me.讀取ERP中介檔ToolStripMenuItem.Text = "讀取單據中介檔"
        '
        '讀取料品中介ToolStripMenuItem
        '
        Me.讀取料品中介ToolStripMenuItem.Name = "讀取料品中介ToolStripMenuItem"
        Me.讀取料品中介ToolStripMenuItem.Size = New System.Drawing.Size(236, 34)
        Me.讀取料品中介ToolStripMenuItem.Text = "讀取料品中介檔"
        '
        '特定料品下載ToolStripMenuItem
        '
        Me.特定料品下載ToolStripMenuItem.Name = "特定料品下載ToolStripMenuItem"
        Me.特定料品下載ToolStripMenuItem.Size = New System.Drawing.Size(236, 34)
        Me.特定料品下載ToolStripMenuItem.Text = "特定料品下載"
        '
        'HBTimer
        '
        Me.HBTimer.Enabled = True
        Me.HBTimer.Interval = 1000
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1140, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(16, 18)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "v"
        '
        'CheckBox_Stop
        '
        Me.CheckBox_Stop.AutoSize = True
        Me.CheckBox_Stop.Location = New System.Drawing.Point(687, 41)
        Me.CheckBox_Stop.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CheckBox_Stop.Name = "CheckBox_Stop"
        Me.CheckBox_Stop.Size = New System.Drawing.Size(142, 22)
        Me.CheckBox_Stop.TabIndex = 5
        Me.CheckBox_Stop.Text = "停止任務受理"
        Me.CheckBox_Stop.UseVisualStyleBackColor = True
        '
        '功能ToolStripMenuItem
        '
        Me.功能ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.基本資料ToolStripMenuItem})
        Me.功能ToolStripMenuItem.Name = "功能ToolStripMenuItem"
        Me.功能ToolStripMenuItem.Size = New System.Drawing.Size(62, 28)
        Me.功能ToolStripMenuItem.Text = "功能"
        '
        '基本資料ToolStripMenuItem
        '
        Me.基本資料ToolStripMenuItem.Name = "基本資料ToolStripMenuItem"
        Me.基本資料ToolStripMenuItem.Size = New System.Drawing.Size(270, 34)
        Me.基本資料ToolStripMenuItem.Text = "基本資料"
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1274, 589)
        Me.Controls.Add(Me.CheckBox_Stop)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.WMSMenuStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FormMain"
        Me.Text = "eHost v20181120"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.WMSMenuStrip.ResumeLayout(False)
        Me.WMSMenuStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents lswtrc As System.Windows.Forms.ListView
    Friend WithEvents WMSMenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents RefreshDBToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RefreshDBToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LogToolToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChangeLogLevelToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HBTimer As System.Windows.Forms.Timer
    Friend WithEvents TSCBViewLogLevel As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents LCSIntitialToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LCSInitialToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ForEcatchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CheckBox_Stop As CheckBox
    Friend WithEvents SendMessageToMESToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HostHandlerToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FromWMSToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T7F1U2POAccountingToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FromMCSToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5R1LineStatusChangeReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5R2LineInfoReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5R3LineInProductionInfoReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5R4LineInProductionInfoResetToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FromGUIToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5U1MaintenanceSetToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5U2MaintenanceToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5U3LineBigDataAlarmSetToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5U4ProductionCountSetToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F5U5ClassProductionSetToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T11F1U11ProducePOExecutionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T3F4R2DeviceAlarmReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T10F4U1MainFileImportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T6F5U1T6F5U1ItemLabelManagementToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T6F5U2ItemLabelPrintToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T10F2S1StocktakingReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 採購單回報ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 盤點單回報ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T5F1U11POExecutionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T11F1U2POExecutionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToWMSToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T11F1U14SwitchOnLocationLightToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T5F1S1WOCloseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 單據放行ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 測試ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TESTToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HOSTTCOMMAND測試ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MQTESTPRIMARYToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MQTESTSECONDARYToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T11F1U1PODownloadToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents T5F1U90WOExcutingToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ERP單據測試ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 讀取ERP中介檔ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 讀取料品中介ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 特定料品下載ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 功能ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 基本資料ToolStripMenuItem As ToolStripMenuItem
End Class
