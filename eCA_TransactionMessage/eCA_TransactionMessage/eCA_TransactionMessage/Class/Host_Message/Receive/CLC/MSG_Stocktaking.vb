Imports System.Xml.Serialization

<XmlRoot(ElementName:="Request")>
Public Class MSG_Stocktaking
  <XmlElement(ElementName:="Access")>
  Public Property Access As New clsAccess
  Public Class clsAccess
    <XmlElement(ElementName:="Authentication")>
    Public Property Authentication As New clsAuthentication()
    Public Class clsAuthentication
      <XmlAttribute("user")> Public Property user As String = ""
      <XmlAttribute("password")> Public Property password As String = ""
    End Class

    Public Property Connection As New clsConnection()
    Public Class clsConnection
      <XmlAttribute("application")> Public Property application As String = ""
      <XmlAttribute("source")> Public Property source As String = ""
    End Class

    Public Property Organization As New clsOrganization()
    Public Class clsOrganization
      <XmlAttribute("name")> Public Property name As String = ""
    End Class

    Public Property Locale As New clsLocale()
    Public Class clsLocale
      <XmlAttribute("language")> Public Property language As String = ""
    End Class

  End Class

  Public Property SendTime As String = ""
  <XmlElement(ElementName:="RequestContent")>
  Public Property RequestContent As New clsRequestContent

  Public Class clsRequestContent
    <XmlElement(ElementName:="Parameter")>
    Public Property Parameter As New clsParameter
    Public Class clsParameter
      <XmlArray("Record")>
      <XmlArrayItem("Field")>
      Public Property Record As Field()
      'Public Class Field
      '  <XmlAttribute("name")> Public Property name As String = ""
      '  <XmlAttribute("value")> Public Property value As String = ""
      'End Class
    End Class

    <XmlElement(ElementName:="Document")>
    Public Property Document As New clsDocument
    Public Class clsDocument
      <XmlElement(ElementName:="RecordSet")>
      Public Property RecordSet As clsRecordSet
      Public Class clsRecordSet
        <XmlAttribute("id")> Public Property id As String = ""

        <XmlElement(ElementName:="Master")>
        Public Property Master As New clsMaster
        Public Property Detail As New clsDetail
        Public Class clsMaster
          <XmlAttribute("name")> Public Property name As String = ""

          <XmlArray("Record")>
          <XmlArrayItem("Field")>
          Public Property Record As Field()

        End Class
        '<XmlElement(ElementName:="Detail")>
        'Public Property Detail As New clsDetail
        Public Class clsDetail
          <XmlAttribute("name")> Public Property name As String = ""

          <XmlArray("Record")>
          <XmlArrayItem("Field")>
          Public Property Record As Field()

        End Class
      End Class
    End Class
  End Class
  Public Class Field
    <XmlAttribute("name")> Public Property name As String = ""
    <XmlAttribute("value")> Public Property value As String = ""
  End Class
End Class
