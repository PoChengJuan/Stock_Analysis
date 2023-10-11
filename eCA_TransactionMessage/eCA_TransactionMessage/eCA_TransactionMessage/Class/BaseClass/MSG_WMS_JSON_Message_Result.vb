Public Class MSG_WMS_JSON_Message_Result
  Public Property Message As New clsMessage
  Public Class clsMessage
    Public Property Header As New clsHeader
    Public Property Body As New clsBody
    Public Property KeepData As String = ""
    Public Class clsHeader
      Public Property UUID As String = ""
      Public Property EventID As String = ""
      Public Property Direction As String = ""
      Public Property SystemID As String = ""
    End Class

    Public Class clsBody
      Public Property ResultInfo As New clsResultinfo
      Public Class clsResultinfo
        Public Property Result As String = ""
        Public Property ResultMessage As String = ""
      End Class
    End Class
  End Class
End Class
