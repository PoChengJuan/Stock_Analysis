Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendWarehouseData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="WarehouseDataList")>
  Public Property WarehouseDataList As New clsWarehouseDataList
  Public Class clsWarehouseDataList
    <XmlElement(ElementName:="WarehouseDataInfo")>
    Public Property WarehouseDataInfo As New List(Of clsWarehouseDataInfo)
    Public Class clsWarehouseDataInfo
      Public Property Owner As String = ""
      Public Property Warehouse As String = ""
      Public Property WarehouseName As String = ""
      Public Property Specification As String = ""
    End Class
  End Class
End Class