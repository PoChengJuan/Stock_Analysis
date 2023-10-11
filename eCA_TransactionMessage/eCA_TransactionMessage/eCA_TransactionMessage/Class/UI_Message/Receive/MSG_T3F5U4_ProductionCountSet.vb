'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F5U4_ProductionCountSet
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody

    <XmlElement(ElementName:="AreaList")>
    Public Property AreaList As New clsAreaList

    Public Class clsAreaList
      <XmlElement(ElementName:="AreaInfo")>
      Public Property AreaInfo As New List(Of clsAreaInfo)
      Public Class clsAreaInfo
        Public Property FACTORY_NO As String
        Public Property AREA_NO As String
        Public Property DEVICE_NO As String
        Public Property UNIT_ID As String
        Public Property QTY_MODIFY As String
        Public Property QTY_NG As String
      End Class
    End Class
  End Class

End Class

