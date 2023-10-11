Imports System.Xml.Serialization

<XmlRoot(ElementName:="Response")>
Public Class MSG_Release_Respones

  <XmlArray("Execution")>
  <XmlArrayItem("status")>
  Public Property Execution As status()

  Public Class status
    <XmlAttribute("code")> Public Property code As String = ""
    <XmlAttribute("sqlcode")> Public Property sqlcode As String = ""
    <XmlAttribute("description")> Public Property description As String = ""
  End Class
  <XmlElement(ElementName:="ResponseContent")>
  Public Property ResponseContent As New clsResponseContent
  Public Class clsResponseContent
    <XmlElement(ElementName:="Parameter")>
    Public Property Parameter As New clsParameter
    Public Class clsParameter
      <XmlArray("Record")>
      <XmlArrayItem("Field")>
      Public Property Record As Field()
      Public Class Field
        <XmlAttribute("name")> Public Property name As String = ""
        <XmlAttribute("value")> Public Property value As String = ""

      End Class
    End Class
  End Class
End Class
