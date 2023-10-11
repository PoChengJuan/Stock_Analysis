Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendReturnData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="ReturnDataList")>
  Public Property ReturnDataList As New clsReturnDataList
  Public Class clsReturnDataList
    <XmlElement(ElementName:="ReturnDataInfo")>
    Public Property ReturnDataInfo As New List(Of clsReturnDataInfo)

    Public Class clsReturnDataInfo
      Public Property POType As String
      Public Property POId As String
      Public Property ReturnDateTime As String
      Public Property FactoryId As String
      <XmlElement(ElementName:="ReturnDetailDataList")>
      Public Property ReturnDetailDataList As clsReturnDetailDataList

      Public Class clsReturnDetailDataList
        <XmlElement(ElementName:="ReturnDetailDataInfo")>
        Public Property ReturnDetailDataInfo As List(Of clsReturnDetailDataInfo)
        Public Class clsReturnDetailDataInfo
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