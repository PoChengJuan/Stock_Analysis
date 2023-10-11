Imports System.Xml.Serialization


<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendTransferOwnerData
  Public Property WebService_ID As String = "" 'WebService_ID
  Public Property EventID As String = ""    'EventID
  <XmlElement(ElementName:="TransferOwnerDataList")>
  Public Property TransferOwnerDataList As New clsTransferOwnerDataList
  Public Class clsTransferOwnerDataList
    <XmlElement(ElementName:="TransferOwnerDataInfo")>
    Public Property TransferOwnerDataInfo As New List(Of clsTransferOwnerDataInfo)

    Public Class clsTransferOwnerDataInfo
      Public Property POType As String = ""
      Public Property POId As String = ""
      Public Property TransferOwnerDateTime As String = ""
      Public Property FactoryId As String = ""
      Public Property TransferOutOwner As String = ""
      Public Property TransferInOwner As String = ""
      <XmlElement(ElementName:="TransferOwnerDetailDataList")>
      Public Property TransferOwnerDetailDataList As clsTransferOwnerDetailDataList

      Public Class clsTransferOwnerDetailDataList
        <XmlElement(ElementName:="TransferOwnerDetailDataInfo")>
        Public Property TransferOwnerDetailDataInfo As List(Of clsTransferOwnerDetailDataInfo)
        Public Class clsTransferOwnerDetailDataInfo
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