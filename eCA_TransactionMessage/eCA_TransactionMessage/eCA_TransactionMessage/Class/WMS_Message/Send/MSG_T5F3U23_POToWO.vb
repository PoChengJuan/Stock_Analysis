Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T5F3U23_POToWO
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody

    Public Property Action As String
    Public Property AutoFlag As String
    <XmlElement(ElementName:="POList")>
    Public Property POList As New clsPOList
    <XmlElement(ElementName:="WOInfo")>
    Public Property WOInfo As New clsWOInfo
    Public Class clsPOList
      <XmlElement(ElementName:="POInfo")>
      Public Property POInfo As New List(Of clsPOInfo)
      Public Class clsPOInfo
        Public Property PO_ID As String
        Public Property PO_SERIAL_NO As String
        Public Property QTY As String
        Public Property COMMENTS As String
        Public Property SOURCE_AREA_NO As String
        Public Property SOURCE_LOCATION_NO As String
      End Class
    End Class
    Public Class clsWOInfo
      Public Property WO_ID As String '工单号 送空
      Public Property SHIPPING_NO As String '出库区域 送空
    End Class
  End Class
End Class
