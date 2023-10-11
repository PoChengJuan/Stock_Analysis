'執行生產工單
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F1U11_ProducePOExecution
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody
    Public Property Action As String
    <XmlElement(ElementName:="POList")>
    Public Property POList As New clsPOList

    Public Class clsPOList
      <XmlElement(ElementName:="POInfo")>
      Public Property POInfo As New List(Of clsPOInfo)

      Public Class clsPOInfo
        Public Property PO_ID As String
        Public Property H_PO_ORDER_TYPE As String
        Public Property COMMENTS As String
        Public Property SOURCE_AREA_NO As String = ""
        Public Property SOURCE_LOCATION_NO As String = ""
      End Class
    End Class
  End Class
End Class

