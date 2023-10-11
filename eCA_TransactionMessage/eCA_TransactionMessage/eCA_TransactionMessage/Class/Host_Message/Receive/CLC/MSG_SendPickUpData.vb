Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendPickUpData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="PickUpDataList")>
  Public Property PickUpDataList As New clsPickUpDataList
  Public Class clsPickUpDataList
    <XmlElement(ElementName:="PickUpDataInfo")>
    Public Property PickUpDataInfo As New List(Of clsPickUpDataInfo)

    Public Class clsPickUpDataInfo
      Public Property POType As String
      Public Property POId As String
      Public Property PickUpTime As String
      Public Property FactoryId As String
      <XmlElement(ElementName:="PickUpDetailDataList")>
      Public Property PickUpDetailDataList As clsPickUpDetailDataList

      Public Class clsPickUpDetailDataList
        <XmlElement(ElementName:="PickUpDetailDataInfo")>
        Public Property PickUpDetailDataInfo As List(Of clsPickUpDetailDataInfo)
        Public Class clsPickUpDetailDataInfo
          Public Property SerialId As String
          Public Property SKU As String
          Public Property CheckQty As String
          Public Property LotId As String
          Public Property Item_Common3 As String
        End Class
      End Class
    End Class
  End Class
End Class