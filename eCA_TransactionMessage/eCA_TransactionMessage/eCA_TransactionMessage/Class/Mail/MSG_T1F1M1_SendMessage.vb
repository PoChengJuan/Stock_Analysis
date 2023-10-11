Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T1F1M1_SendMessage
  Public Property Header As New HeaderDataInfo
  Public Property Body As New BodyInfo
  'Public Property SendList As SendDataList

  'Public Class SendDataList
  '  Inherits BodyInfo
  '  Public Property SendList As SendDataList
  '  Public Class SendDataList
  '    'Inherits List(Of SendDataInfo)
  '    <XmlElement(ElementName:="SendInfo")>
  '    Public Property SendInfo As New List(Of SendDataInfo)
  '  End Class
  '  Public Class SendDataInfo
  '    Public Property SendType As String
  '    Public Property RecipientList As String
  '  End Class
  'End Class

  Public Class HeaderDataInfo
    Public Property UUID As String = ""
    Public Property EventID As String = ""
    Public Property Direction As String = ""
    Public Property SystemID As String = "HostHandler"
  End Class

  Public Class BodyInfo

    Public Property MessageList As New MessageDataList
    Public Class MessageDataList
      'Inherits List(Of MessageDataInfo)
      <XmlElement(ElementName:="MessageInfo")>
      Public Property MessageInfo As New List(Of MessageDataInfo)
    End Class
    Public Class MessageDataInfo
      Public Property MESSAGE_TITLE As String = ""
      Public Property MESSAGE_TEXT As String = ""
      Public Property ATTACHMENT_PATH As String = ""
    End Class

    Public Property SendList As New SendDataList
    Public Class SendDataList
      'Inherits List(Of SendDataInfo)
      <XmlElement(ElementName:="SendInfo")>
      Public Property SendInfo As New List(Of SendDataInfo)
    End Class
    Public Class SendDataInfo
      Public Property SEND_TYPE As String
      Public Property SEND_RETRY_COUNT As String
      Public Property SEND_INFO As String '接收的人員 mail
    End Class
  End Class

End Class