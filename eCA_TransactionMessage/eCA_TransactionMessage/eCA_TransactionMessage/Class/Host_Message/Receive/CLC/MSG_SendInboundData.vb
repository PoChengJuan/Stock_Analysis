Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendInboundData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="InboundDataList")>
  Public Property InboundDataList As New clsInboundDataList
  Public Class clsInboundDataList
    <XmlElement(ElementName:="InboundDataInfo")>
    Public Property InboundDataInfo As New List(Of clsInboundDataInfo)
    Public Class clsInboundDataInfo
      Public Property POType As String = ""
      Public Property POId As String = ""
      Public Property InboundDateTime As String = ""
      Public Property FactoryId As String = ""
      'Public Property Warehouse As String = ""
      <XmlElement(ElementName:="InboundDetailDataList")>
      Public Property InboundDetailDataList As clsInboundDetailDataList

      Public Class clsInboundDetailDataList
        <XmlElement(ElementName:="InboundDetailDataInfo")>
        Public Property InboundDetailDataInfo As List(Of clsInboundDetailDataInfo)
        Public Class clsInboundDetailDataInfo
          Public Property SerialId As String = ""
          Public Property SKU As String = ""
          Public Property LotId As String = ""
          Public Property CheckQty As String = ""
          Public Property Item_Common1 As String = ""
          Public Property Item_Common2 As String = ""
        End Class
      End Class
    End Class
  End Class
End Class