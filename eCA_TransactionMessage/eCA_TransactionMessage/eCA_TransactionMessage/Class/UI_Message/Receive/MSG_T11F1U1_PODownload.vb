
'MSG_T11F1U1_PODownload

Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F1U1_PODownload
  Public Property Header As clsHeader
  Public Property Body As BodyData
  Public Property KeepData As String

  Public Class BodyData

    Public Property Action As String
    <XmlElement(ElementName:="POList")>
    Public Property POList As New PODataList

    Public Class PODataList
      <XmlElement(ElementName:="POInfo")>
      Public Property POInfo As New List(Of PODataInfo)
    End Class
  End Class

  Public Class PODataInfo
    Public Property PO_ID As String
    Public Property H_PO_ORDER_TYPE As String
    Public Property LGORT As String
    Public Property FORCED_UPDATE As String
    Public Property ERP_ORDER_TYPE As String
    Public Property COMMON01 As String
    Public Property COMMON02 As String
    Public Property COMMON03 As String
    Public Property COMMON04 As String
    Public Property COMMON05 As String



  End Class

End Class

