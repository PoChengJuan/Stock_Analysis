Imports System.Xml.Serialization
<XmlRoot(ElementName:="Message")>
Public Class MSG_T2F3U11_PackeUnitManagement
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String
  Public Class clsBody
    Public Property Action As String
    <XmlElement(ElementName:="PackeUnitList")>
    Public Property PackeUnitList As clsPackeUnitList

    Public Class clsPackeUnitList
      <XmlElement(ElementName:="PackeUnitInfo")>
      Public Property PackeUnitInfo As New List(Of clsPackeUnitInfo)

      Public Class clsPackeUnitInfo
        Public Property PACKE_UNIT As String = ""
        Public Property PACKE_UNIT_NAME As String = ""
        Public Property PACKE_UNIT_COMMON1 As String = ""
        Public Property PACKE_UNIT_COMMON2 As String = ""
        Public Property PACKE_UNIT_COMMON3 As String = ""
        Public Property PACKE_UNIT_COMMON4 As String = ""
        Public Property PACKE_UNIT_COMMON5 As String = ""
        Public Property COMMENTS As String = ""
      End Class
    End Class
  End Class
End Class