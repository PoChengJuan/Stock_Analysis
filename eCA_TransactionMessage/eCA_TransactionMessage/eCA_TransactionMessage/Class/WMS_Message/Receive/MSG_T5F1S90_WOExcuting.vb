Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T5F1S90_WOExcuting
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String
  Public Class clsBody
    Public Property POList As New clsPOList

    Public Class clsPOList
      <XmlElement(ElementName:="POInfo")>
      Public Property POInfo As New List(Of clsPOInfo)

      Public Class clsPOInfo
        Public Property PO_ID As String = ""
      End Class
    End Class
  End Class
End Class
