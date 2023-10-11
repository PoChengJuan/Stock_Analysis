Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendOtherOutData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="PickingDataList")>
  Public Property PickingDataList As New clsPickingDataList
  Public Class clsPickingDataList
    <XmlElement(ElementName:="PickingDataInfo")>
    Public Property PickingDataInfo As New List(Of clsPickingDataInfo)

    Public Class clsPickingDataInfo
      Public Property POType As String
      Public Property POId As String
      Public Property PickingDateTime As String
      Public Property FactoryId As String
      Public Property ProductionLine As String
      Public Property Kind As String
      Public Property ConfirmCode As String
      <XmlElement(ElementName:="PickingDetailDataList")>
      Public Property PickingDetailDataList As clsPickingDetailDataList

      Public Class clsPickingDetailDataList
        <XmlElement(ElementName:="PickingDetailDataInfo")>
        Public Property PickingDetailDataInfo As List(Of clsPickingDetailDataInfo)
        Public Class clsPickingDetailDataInfo
          Public Property SerialId As String
          Public Property SKU As String
          Public Property PickingQty As String
          Public Property Unit As String
          Public Property WarehouseId As String
          Public Property LotId As String
          Public Property ShelfId As String
        End Class
      End Class
    End Class
  End Class
End Class