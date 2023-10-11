'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F5S1_LineProductionResetRequest
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody
    <XmlElement(ElementName:="UnitList")>
    Public Property UnitList As New clsUnitList
    Public Class clsUnitList
      <XmlElement(ElementName:="UnitInfo")>
      Public Property UnitInfo As New List(Of clsUnitInfo)
      Public Class clsUnitInfo
        Public Property FACTORY_NO As String
        Public Property DEVICE_NO As String
        Public Property AREA_NO As String
        Public Property UNIT_ID As String
      End Class
    End Class
  End Class
End Class

