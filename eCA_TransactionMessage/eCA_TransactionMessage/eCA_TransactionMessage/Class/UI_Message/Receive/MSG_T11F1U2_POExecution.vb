'MSG_T11F1U2_POExecution
'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F1U2_POExecution
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
    Public Property COMMENTS As String


  End Class

End Class

