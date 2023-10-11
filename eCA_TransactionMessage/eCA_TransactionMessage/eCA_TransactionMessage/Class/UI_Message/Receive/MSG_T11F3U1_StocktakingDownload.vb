
'MSG_T11F1U1_PODownload

Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F3U1_StocktakingDownload
  Public Property Header As clsHeader
  Public Property Body As BodyData
  Public Property KeepData As String

  Public Class BodyData

    Public Property Action As String
    <XmlElement(ElementName:="DataInfo")>
    Public Property DataInfo As New clsDataInfo

    Public Class clsDataInfo
      Public Property FromDate As String
      Public Property ToDate As String
    End Class
  End Class
End Class

