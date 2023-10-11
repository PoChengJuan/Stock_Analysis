Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendOtherOutData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="OtherOutDataList")>
  Public Property OtherOutDataList As New clsOtherOutDataList
  Public Class clsOtherOutDataList
    <XmlElement(ElementName:="OtherOutDataInfo")>
    Public Property OtherOutDataInfo As New List(Of clsOtherOutDataInfo)

    Public Class clsOtherOutDataInfo
      Public Property POType As String = ""
      Public Property POId As String = ""
      Public Property OtherOutDateTime As String = ""
      Public Property FactoryId As String = ""
      Public Property Warehouse As String = ""
      <XmlElement(ElementName:="OtherOutDetailDataList")>
      Public Property OtherOutDetailDataList As clsOtherOutDetailDataList

      Public Class clsOtherOutDetailDataList
        <XmlElement(ElementName:="OtherOutDetailDataInfo")>
        Public Property OtherOutDetailDataInfo As List(Of clsOtherOutDetailDataInfo)
        Public Class clsOtherOutDetailDataInfo
          Public Property SerialId As String = ""
          Public Property SKU As String = ""
          Public Property CheckQty As String = ""
          Public Property SN As String = ""
          Public Property LotId As String = ""
        End Class
      End Class
    End Class
  End Class
End Class