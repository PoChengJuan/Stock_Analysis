'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F5R3_LineInProductionInfoReport
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody
    <XmlElement(ElementName:="ProductionList")>
    Public Property ProductionList As New clsProductionList
    Public Class clsProductionList
      <XmlElement(ElementName:="ProductionInfo")>
      Public Property ProductionInfo As New List(Of clsProductionInfo)
      Public Class clsProductionInfo
        Public Property FACTORY_NO As String
        Public Property DEVICE_NO As String
        Public Property AREA_NO As String
        Public Property UNIT_ID As String
        Public Property QTY_TOTAL As String
        Public Property QTY_PROCESS As String
        Public Property QTY_MODIFY As String
        Public Property QTY_NG As String
      End Class
    End Class
  End Class
End Class

