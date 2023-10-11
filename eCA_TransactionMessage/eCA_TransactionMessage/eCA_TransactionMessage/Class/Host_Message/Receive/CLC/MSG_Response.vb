Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_Response
  <XmlElement(ElementName:="ReportData")>
  Public Property ReportData As New List(Of clsReportData)
  Public Class clsReportData
    Public Property Result As String = ""
    Public Property ResultMessage As String = ""
  End Class

End Class