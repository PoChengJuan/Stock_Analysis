'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F5R2_LineInfoReport
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody
    <XmlElement(ElementName:="MessageList")>
    Public Property MessageList As New clsMessageList
    Public Class clsMessageList
      <XmlElement(ElementName:="MessageInfo")>
      Public Property MessageInfo As New List(Of clsMessageInfo)
      Public Class clsMessageInfo
        Public Property FACTORY_NO As String
        Public Property DEVICE_NO As String
        Public Property AREA_NO As String
        Public Property UNIT_ID As String
        Public Property TIME As String
        Public Property MESSAGE As String
        Public Property SET_FLAG As String
      End Class
    End Class
  End Class
End Class

