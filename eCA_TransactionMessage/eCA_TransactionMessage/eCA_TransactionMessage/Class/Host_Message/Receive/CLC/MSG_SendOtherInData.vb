Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendOtherInData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="OtherInDataList")>
  Public Property OtherInDataList As New clsOtherInDataList

  Public Class clsOtherInDataList
    <XmlElement(ElementName:="OtherInDataInfo")>
    Public Property OtherInDataInfo As New List(Of clsOtherInDataInfo)

    Public Class clsOtherInDataInfo
      Public Property POType As String = ""
      Public Property POId As String = ""
      Public Property OtherInDateTime As String = ""
      Public Property FactoryId As String = ""
      Public Property Warehouse As String = ""
      <XmlElement(ElementName:="OtherInDetailDataList")>
      Public Property OtherInDetailDataList As clsOtherInDetailDataList

      Public Class clsOtherInDetailDataList
        <XmlElement(ElementName:="OtherInDetailDataInfo")>
        Public Property OtherInDetailDataInfo As List(Of clsOtherInDetailDataInfo)
        Public Class clsOtherInDetailDataInfo
          Public Property SerialId As String = ""
          Public Property SKU As String = ""
          Public Property LotId As String = ""
          Public Property SN As String = ""
          Public Property CheckQty As String = ""
        End Class
      End Class
    End Class
  End Class
End Class