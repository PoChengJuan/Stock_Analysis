Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F2U1_InventoryComparison
  Public Property Header As clsHeader
  Public Property Body As BodyData
  Public Property KeepData As String

  Public Class BodyData

    Public Property Action As String
    <XmlElement(ElementName:="AccountInfo")>
    Public Property AccountInfo As New clsAccountInfo

    Public Class clsAccountInfo
      Public Property FROM_DATE As String
      Public Property TO_DATE As String
    End Class
  End Class


End Class

