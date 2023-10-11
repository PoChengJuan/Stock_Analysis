Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendProduceInData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="ProduceInDataList")>
  Public Property ProduceInDataList As New clsProduceInDataList

  Public Class clsProduceInDataList
    <XmlElement(ElementName:="ProduceInDataInfo")>
    Public Property ProduceInDataInfo As New List(Of clsProduceInDataInfo)

    Public Class clsProduceInDataInfo
      Public Property POType As String = ""
      Public Property POId As String = ""
      Public Property ProduceInDateTime As String = ""
      Public Property FactoryId As String = ""
      ' Public Property Warehouse As String = ""
      <XmlElement(ElementName:="ProduceInDetailDataList")>
      Public Property ProduceInDetailDataList As clsProduceInDetailDataList

      Public Class clsProduceInDetailDataList
        <XmlElement(ElementName:="ProduceInDetailDataInfo")>
        Public Property ProduceInDetailDataInfo As List(Of clsProduceInDetailDataInfo)
        Public Class clsProduceInDetailDataInfo
          Public Property SerialId As String = ""
          Public Property SKU As String = ""
          Public Property LotId As String = ""
          'Public Property SN As String = ""
          Public Property CheckQty As String = ""
          Public Property Item_Common3 As String = ""
        End Class
      End Class
    End Class
  End Class
End Class