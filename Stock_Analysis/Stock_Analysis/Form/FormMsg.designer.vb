<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FormMsg
#Region "Windows Form �]�p�u�㲣�ͪ��{���X "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        '���� Windows Form �]�p�u��һݪ��I�s�C
        InitializeComponent()
    End Sub
    'Form �мg Dispose �H�M������M��C
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    '�� Windows Form �]�p�u�㪺���n��
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents lblMsg As System.Windows.Forms.Label
    '�`�N: �H�U�� Windows Form �]�p�u��һݪ��{��
    '�i�H�ϥ� Windows Form �]�p�u��i��ק�C
    '�Ф��n�ϥε{���X�s�边�i��ק�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMsg))
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.lblMsg = New System.Windows.Forms.Label()
    Me.SuspendLayout()
    '
    'lblMsg
    '
    Me.lblMsg.BackColor = System.Drawing.SystemColors.Control
    Me.lblMsg.Cursor = System.Windows.Forms.Cursors.Default
    Me.lblMsg.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblMsg.ForeColor = System.Drawing.SystemColors.ControlText
    Me.lblMsg.Location = New System.Drawing.Point(6, 8)
    Me.lblMsg.Name = "lblMsg"
    Me.lblMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lblMsg.Size = New System.Drawing.Size(766, 40)
    Me.lblMsg.TabIndex = 0
    Me.lblMsg.Text = "Message"
    Me.lblMsg.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '
    'FormMsg
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackColor = System.Drawing.SystemColors.Control
    Me.ClientSize = New System.Drawing.Size(778, 56)
    Me.ControlBox = False
    Me.Controls.Add(Me.lblMsg)
    Me.Cursor = System.Windows.Forms.Cursors.Default
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Location = New System.Drawing.Point(3, 22)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "FormMsg"
    Me.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = " "
    Me.ResumeLayout(False)

  End Sub
#End Region
End Class