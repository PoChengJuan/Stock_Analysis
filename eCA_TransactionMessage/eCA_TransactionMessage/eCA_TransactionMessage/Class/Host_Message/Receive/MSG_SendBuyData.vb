Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendInData
  Public Property identity As clsidentity
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String = ""       'EventID
  <XmlElement(ElementName:="InDataList")>
  Public Property InDataList As New clsInDataList
  Public Class clsInDataList
    <XmlElement(ElementName:="InDataInfo")>
    Public Property InDataInfo As New List(Of clsInDataInfo)
    Public Class clsInDataInfo
      Public Property POType As String = ""
      Public Property POID As String = ""
      Public Property InDateTime As String = ""
      Public Property FactoryID As String = ""
      Public Property Owner As String = ""
      <XmlElement(ElementName:="InDetailDataList")>
      Public Property InDetailDataList As clsInDetailDataList

      Public Class clsInDetailDataList
        <XmlElement(ElementName:="InDetailDataInfo")>
        Public Property InDetailDataInfo As List(Of clsInDetailDataInfo)
        Public Class clsInDetailDataInfo
          Public Property SerialID As String = ""
          Public Property SKU As String = ""
          Public Property LotId As String = ""
          Public Property CheckQty As String = ""
          Public Property Item_Common1 As String = ""
          Public Property Item_Common2 As String = ""
          Public Property Item_Common3 As String = ""
        End Class
      End Class
    End Class
  End Class
End Class