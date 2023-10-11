Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendSellReturnData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="SellReturnDataList")>
  Public Property SellReturnDataList As New clsSellReturnDataList
  Public Class clsSellReturnDataList
    <XmlElement(ElementName:="SellReturnDataInfo")>
    Public Property SellReturnDataInfo As New List(Of clsSellReturnDataInfo)

    Public Class clsSellReturnDataInfo
      Public Property POType As String = ""
      Public Property POId As String = ""
      Public Property SellReturnDateTime As String = ""
      Public Property FactoryId As String = ""
      Public Property Warehouse As String = ""
      <XmlElement(ElementName:="SellReturnDetailDataList")>
      Public Property SellReturnDetailDataList As clsSellReturnDetailDataList

      Public Class clsSellReturnDetailDataList
        <XmlElement(ElementName:="SellReturnDetailDataInfo")>
        Public Property SellReturnDetailDataInfo As List(Of clsSellReturnDetailDataInfo)
        Public Class clsSellReturnDetailDataInfo
          Public Property SerialId As String = ""
          Public Property SKU As String = ""
          Public Property CheckQty As String = ""
          Public Property LotId As String = ""
        End Class
      End Class
    End Class
  End Class
End Class