<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMessageTest
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
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
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.txt_SendMessage = New System.Windows.Forms.TextBox()
    Me.gb_SendMessage = New System.Windows.Forms.GroupBox()
    Me.gb_ResultMessage = New System.Windows.Forms.GroupBox()
    Me.txt_ResultMessage = New System.Windows.Forms.TextBox()
    Me.btn_SendMessage = New System.Windows.Forms.Button()
    Me.gb_SendMessage.SuspendLayout()
    Me.gb_ResultMessage.SuspendLayout()
    Me.SuspendLayout()
    '
    'txt_SendMessage
    '
    Me.txt_SendMessage.Location = New System.Drawing.Point(6, 24)
    Me.txt_SendMessage.Multiline = True
    Me.txt_SendMessage.Name = "txt_SendMessage"
    Me.txt_SendMessage.Size = New System.Drawing.Size(715, 445)
    Me.txt_SendMessage.TabIndex = 0
    '
    'gb_SendMessage
    '
    Me.gb_SendMessage.Controls.Add(Me.txt_SendMessage)
    Me.gb_SendMessage.Location = New System.Drawing.Point(12, 12)
    Me.gb_SendMessage.Name = "gb_SendMessage"
    Me.gb_SendMessage.Size = New System.Drawing.Size(727, 476)
    Me.gb_SendMessage.TabIndex = 1
    Me.gb_SendMessage.TabStop = False
    Me.gb_SendMessage.Text = "SendMessage"
    '
    'gb_ResultMessage
    '
    Me.gb_ResultMessage.Controls.Add(Me.txt_ResultMessage)
    Me.gb_ResultMessage.Location = New System.Drawing.Point(12, 494)
    Me.gb_ResultMessage.Name = "gb_ResultMessage"
    Me.gb_ResultMessage.Size = New System.Drawing.Size(727, 147)
    Me.gb_ResultMessage.TabIndex = 2
    Me.gb_ResultMessage.TabStop = False
    Me.gb_ResultMessage.Text = "ResultMessage"
    '
    'txt_ResultMessage
    '
    Me.txt_ResultMessage.Location = New System.Drawing.Point(6, 24)
    Me.txt_ResultMessage.Multiline = True
    Me.txt_ResultMessage.Name = "txt_ResultMessage"
    Me.txt_ResultMessage.Size = New System.Drawing.Size(715, 117)
    Me.txt_ResultMessage.TabIndex = 3
    '
    'btn_SendMessage
    '
    Me.btn_SendMessage.Location = New System.Drawing.Point(745, 405)
    Me.btn_SendMessage.Name = "btn_SendMessage"
    Me.btn_SendMessage.Size = New System.Drawing.Size(120, 83)
    Me.btn_SendMessage.TabIndex = 3
    Me.btn_SendMessage.Text = "SendMessage"
    Me.btn_SendMessage.UseVisualStyleBackColor = True
    '
    'FrmMessageTest
    '
    Me.ClientSize = New System.Drawing.Size(879, 647)
    Me.Controls.Add(Me.btn_SendMessage)
    Me.Controls.Add(Me.gb_ResultMessage)
    Me.Controls.Add(Me.gb_SendMessage)
    Me.Name = "FrmMessageTest"
    Me.gb_SendMessage.ResumeLayout(False)
    Me.gb_SendMessage.PerformLayout()
    Me.gb_ResultMessage.ResumeLayout(False)
    Me.gb_ResultMessage.PerformLayout()
    Me.ResumeLayout(False)

  End Sub

  Friend WithEvents txt_SendMessage As TextBox
  Friend WithEvents gb_SendMessage As GroupBox
  Friend WithEvents gb_ResultMessage As GroupBox
  Friend WithEvents txt_ResultMessage As TextBox
  Friend WithEvents btn_SendMessage As Button
End Class
