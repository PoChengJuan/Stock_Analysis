Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_InventoryData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="InventoryDataList")>
  Public Property InventoryDataList As New clsInventoryDataList
  Public Class clsInventoryDataList
    <XmlElement(ElementName:="InventoryDataInfo")>
    Public Property InventoryDataInfo As New List(Of clsInventoryDataInfo)
    Public Class clsInventoryDataInfo
      Public Property Inventory_ID As String = ""
      <XmlElement(ElementName:="InventoryDetailDataList")>
      Public Property InventoryDetailDataList As New clsInventoryDetailDataList
      Public Class clsInventoryDetailDataList
        <XmlElement(ElementName:="InventoryDetailDataInfo")>
        Public Property InventoryDetailDataInfo As New List(Of clsInventoryDetailDataInfo)
        Public Class clsInventoryDetailDataInfo
          Public Property SerialId As String = ""
          Public Property SKU As String = ""
          Public Property LotId As String = ""
          Public Property InventoryQty As String = ""
        End Class
      End Class
    End Class
  End Class
End Class