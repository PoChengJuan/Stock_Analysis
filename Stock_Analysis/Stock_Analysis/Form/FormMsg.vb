Option Strict Off
Option Explicit On

Public Class FormMsg
  Inherits System.Windows.Forms.Form

  Private Sub frmMsg_Load(ByVal eventSender As System.Object, _
                          ByVal eventArgs As System.EventArgs) _
          Handles MyBase.Load
    '�ت�:���J�T������
    TimeDelay(0.5)
  End Sub
  Public Sub SetFormMessage(ByVal msg As String)
    Try
      Me.lblMsg.Text = msg
    Catch ex As Exception

    End Try
  End Sub
End Class
