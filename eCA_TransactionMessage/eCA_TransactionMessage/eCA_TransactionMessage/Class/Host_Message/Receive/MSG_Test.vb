Imports System.Xml.Serialization
<XmlRoot(ElementName:="soapenv:Envelope")>
Public Class MSG_Test
  <XmlElement(ElementName:="soapenv:Body")>
  Public Property Body As New clsBody
  Public Class clsBody
    <XmlElement(ElementName:="out:MT_ZRFCWMS021_Out_Res")>
    Public Property out As New clsout

    Public Class clsout
      <XmlElement(ElementName:="item")>
      Public Property item As New List(Of clsitem)

      Public Class clsitem
        Public Property EBELN As String = ""
        Public Property BSART As String = ""
        Public Property AEDAT As String = ""
      End Class
    End Class
  End Class
End Class