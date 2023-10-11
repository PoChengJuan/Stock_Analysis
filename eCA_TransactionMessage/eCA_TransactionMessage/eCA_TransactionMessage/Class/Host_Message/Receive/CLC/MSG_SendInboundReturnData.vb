Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendInboundReturnData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="InboundReturnDataList")>
  Public Property InboundReturnDataList As New clsInboundReturnDataList
  Public Class clsInboundReturnDataList
    <XmlElement(ElementName:="InboundReturnDataInfo")>
    Public Property InboundReturnDataInfo As New List(Of clsInboundReturnDataInfo)
    Public Class clsInboundReturnDataInfo
      Public Property POType As String = ""
      Public Property POId As String = ""
      Public Property InboundReturnDateTime As String = ""
      Public Property FactoryId As String = ""
      Public Property Warehouse As String = ""
      <XmlElement(ElementName:="InboundReturnDetailDataList")>
      Public Property InboundReturnDetailDataList As clsInboundReturnDetailDataList

      Public Class clsInboundReturnDetailDataList
        <XmlElement(ElementName:="InboundReturnDetailDataInfo")>
        Public Property InboundReturnDetailDataInfo As List(Of clsInboundReturnDetailDataInfo)
        Public Class clsInboundReturnDetailDataInfo
          Public Property SerialId As String = ""
          Public Property SKU As String = ""
          Public Property SN As String = ""
          Public Property CheckQty As String = ""
        End Class
      End Class
    End Class
  End Class
End Class