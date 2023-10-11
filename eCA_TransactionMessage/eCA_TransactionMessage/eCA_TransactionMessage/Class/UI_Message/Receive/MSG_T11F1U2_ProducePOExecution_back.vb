'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F1U2_ProducePOExecution
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
        Public Property PO_TYPE1 As String
        Public Property PO_TYPE2 As String
        Public Property PO_TYPE3 As String
      End Class
    End Class
  End Class

End Class

