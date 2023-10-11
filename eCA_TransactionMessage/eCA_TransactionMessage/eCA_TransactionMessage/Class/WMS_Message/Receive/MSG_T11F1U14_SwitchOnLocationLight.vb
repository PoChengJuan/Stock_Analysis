Imports System.Xml.Serialization
Imports eCA_MessageConversion
<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F1U14_SwitchOnLocationLight
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String

  Public Class clsBody
    Public Property Mode As String = ""
    <XmlElement(ElementName:="LocationList")>
    Public Property LocationList As New clsLocationList
    Public Class clsLocationList
      <XmlElement(ElementName:="LocationInfo")>
      Public Property LocationInfo As New List(Of clsLocationInfo)

      Public Class clsLocationInfo
        Public Property LOCATION_NO As String = ""
        Public Property SKU_NO As String = ""
      End Class
    End Class
  End Class
End Class

'End Class
