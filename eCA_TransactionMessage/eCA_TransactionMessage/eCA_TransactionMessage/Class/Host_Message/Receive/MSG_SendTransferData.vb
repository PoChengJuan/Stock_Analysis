Imports System.Xml.Serialization


<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendTransferData
  Public Property WebService_ID As String = "" 'WebService_ID
  Public Property EventID As String = ""    'EventID
  <XmlElement(ElementName:="TransferDataList")>
  Public Property TransferDataList As New clsTransferDataList
  Public Class clsTransferDataList
    <XmlElement(ElementName:="TransferDataInfo")>
    Public Property TransferDataInfo As New List(Of clsTransferDataInfo)

    Public Class clsTransferDataInfo
      Public Property POType As String = ""
      Public Property POId As String = ""
      Public Property TransferDateTime As String = ""
      Public Property FactoryId As String = ""
      Public Property TransferOutWarehouse As String = ""
      Public Property TransferInWarehouse As String = ""
      Public Property Owner As String = ""
      <XmlElement(ElementName:="TransferDetailDataList")>
      Public Property TransferDetailDataList As clsTransferDetailDataList

      Public Class clsTransferDetailDataList
        <XmlElement(ElementName:="TransferDetailDataInfo")>
        Public Property TransferDetailDataInfo As List(Of clsTransferDetailDataInfo)
        Public Class clsTransferDetailDataInfo
          Public Property SerialId As String = ""
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