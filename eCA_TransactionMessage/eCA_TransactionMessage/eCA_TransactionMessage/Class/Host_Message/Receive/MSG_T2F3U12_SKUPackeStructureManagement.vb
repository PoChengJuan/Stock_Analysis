Imports System.Xml.Serialization
<XmlRoot(ElementName:="Message")>
Public Class MSG_T2F3U12_SKUPackeStructureManagement
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String
  Public Class clsBody
    Public Property Action As String
    <XmlElement(ElementName:="PackeStructureList")>
    Public Property PackeStructureList As New clsPackeStructureList

    Public Class clsPackeStructureList
      <XmlElement(ElementName:="PackeStructureInfo")>
      Public Property PackeStructureInfo As New List(Of clsPackeStructureInfo)

      Public Class clsPackeStructureInfo
        Public Property SKU_NO As String = ""
        Public Property PACKE_LV As String = ""
        Public Property PACKE_UNIT As String = ""
        Public Property SUB_PACKE_UNIT As String = ""
        Public Property PACKE_WEIGHT As String = ""
        Public Property PACKE_VOLUME As String = ""
        Public Property PACKE_BCR As String = ""
        Public Property OUT_MAX_UNIT As String = ""
        Public Property IN_MAX_UNIT As String = ""
        Public Property QTY As String = ""
        Public Property COMMENTS As String = ""
      End Class
    End Class
  End Class
End Class