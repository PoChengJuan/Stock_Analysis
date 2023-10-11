Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendOutsourceData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="OutsourcePurchaseDataList")>
  Public Property OutsourcePurchaseDataList As New clsOutsourcePurchaseDataList

  Public Class clsOutsourcePurchaseDataList
    <XmlElement(ElementName:="OutsourcePurchaseDataInfo")>
    Public Property OutsourcePurchaseDataInfo As New List(Of clsOutsourcePurchaseDataInfo)

    Public Class clsOutsourcePurchaseDataInfo
      Public Property POType As String
      Public Property POId As String
      Public Property OutsourcePurchaseDateTime As String
      Public Property FactoryId As String
      Public Property ConfirmCode As String
      <XmlElement(ElementName:="OutsourcePurchaseDetailDataList")>
      Public Property OutsourcePurchaseDetailDataList As clsOutsourcePurchaseDetailDataList

      Public Class clsOutsourcePurchaseDetailDataList
        <XmlElement(ElementName:="OutsourcePurchaseDetailDataInfo")>
        Public Property OutsourcePurchaseDetailDataInfo As List(Of clsOutsourcePurchaseDetailDataInfo)
        Public Class clsOutsourcePurchaseDetailDataInfo
          Public Property SerialId As String
          Public Property SKU As String
          Public Property Unit As String
          Public Property WarehouseId As String
          Public Property LotId As String
          Public Property CheckQty As String
          Public Property ShelfId As String
        End Class
      End Class
    End Class
  End Class
End Class