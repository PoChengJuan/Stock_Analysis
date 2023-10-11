Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_WMS_XML_Message_Result
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Class clsBody
    <XmlElement(ElementName:="ResultInfo")>
    Public Property ResultInfo As New clsResultInfo

    Public Class clsResultInfo
      Public Property Result As String = enuResultInfo.OK
      Public Property ResultMessage As String = ""

    End Class
  End Class
End Class