Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T5F1S13_TransationPOExecution
  Public Property Header As clsHeader
  Public Property Body As BodyData
  Public Property KeepData As String

  Public Class BodyData
    <XmlElement(ElementName:="POList")>
    Public Property POList As New clsPOList
    Public Class clsPOList
      <XmlElement(ElementName:="POInfo")>
      Public Property POInfo As New List(Of clsPOInfo)
      Public Class clsPOInfo
        Public Property PO_ID As String
      End Class
    End Class
  End Class
End Class