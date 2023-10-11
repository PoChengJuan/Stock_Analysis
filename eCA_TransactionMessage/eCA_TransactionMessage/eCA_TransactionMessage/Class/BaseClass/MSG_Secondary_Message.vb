Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_Secondary_Message
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String
  Public Class clsBody
    <XmlElement(ElementName:="ResultInfo")>
    Public Property ResultInfo As New clsResultInfo
    Public Class clsResultInfo
      Public Property Result As String = ""
      Public Property ResultMessage As String = ""
    End Class
    <XmlElement(ElementName:="Report")>
    Public Property Report As New clsReport
    Public Class clsReport
      <XmlElement(ElementName:="LabelList")>
      Public Property LabelList As New clsLabelList
      Public Class clsLabelList
        <XmlElement(ElementName:="LabelInfo")>
        Public Property LabelInfo As New List(Of clsLabelInfo)
        Public Class clsLabelInfo
          Public Property SEQ_NO As String = ""
          Public Property LABEL_ID As String = ""
        End Class
      End Class
    End Class
  End Class
End Class