'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F5R1_LineStatusChangeReport
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody
    <XmlElement(ElementName:="DeviceInfo")>
    Public Property DeviceInfo As New clsDeviceInfo
    <XmlElement(ElementName:="LineList")>
    Public Property LineList As New clsLineList
    Public Class clsDeviceInfo
      Public Property FACTORY_NO As String
      Public Property DEVICE_NO As String
    End Class

    Public Class clsLineList
      <XmlElement(ElementName:="LineInfo")>
      Public Property LineInfo As New List(Of clsLineInfo)
      Public Class clsLineInfo
        Public Property AREA_NO As String
        Public Property UNIT_ID As String
        Public Property STATUS As String
      End Class
    End Class
  End Class
End Class

