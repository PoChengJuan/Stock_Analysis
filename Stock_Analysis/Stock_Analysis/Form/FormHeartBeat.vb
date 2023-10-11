Public Class FormHeartBeat
  Public mHB As HeartBeatObject.CHeartBeatUse

  Dim _RtnMsg As String
  Property RtnMsg As String
    Get
      Return _RtnMsg
    End Get
    Set(value As String)
      If value <> _RtnMsg Then
        MsgBox(value)
        _RtnMsg = value
      End If

    End Set
  End Property



  Sub New(ByVal _portno As Integer, ByVal _interval As Integer)

    ' This call is required by the designer.
    InitializeComponent()
    If CreateHBServer(_portno, RtnMsg) = 0 Then
      Timer1.Interval = _interval
      Timer1.Start()
    End If
    
    ' Add any initialization after the InitializeComponent() call.

  End Sub
  Protected Overrides Sub Finalize() '解構子
    MyBase.Finalize()
    Timer1.Stop()
  End Sub




  Private Function CreateHBServer(ByVal port As String, ByRef RtnMsg As String) As Integer
    Dim msg As String = ""
    Try
      mHB = New HeartBeatObject.CHeartBeatUse

      If mHB.CreateHeartBeatServer(port, RtnMsg) < 0 Then
        'Create Heart Beat Server failed
        msg = RtnMsg
        Return -1
      Else
        'Successfully Create Heart Beat Server
        msg = "OK"
        Return 0
      End If

    Catch ex As Exception
      msg = ex.ToString()
      Return -1
    End Try
  End Function
 
  Private Sub Timer1_Tick_1(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
    Try
      mHB.UpdateHeartBeat()
      Return
    Catch ex As Exception
      Return
    End Try
  End Sub
End Class