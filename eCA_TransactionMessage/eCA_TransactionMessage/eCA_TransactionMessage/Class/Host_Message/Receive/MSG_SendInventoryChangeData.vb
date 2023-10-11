Imports System.Xml.Serialization


<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendInventoryChangeData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="TransactionDataList")>
  Public Property TransactionDataList As New clsTransactionDataList
  Public Class clsTransactionDataList
    <XmlElement(ElementName:="TransactionDataInfo")>
    Public Property TransactionDataInfo As New List(Of clsTransactionDataInfo)

    Public Class clsTransactionDataInfo
      Public Property POType As String
      Public Property POId As String
      Public Property TransactionDateTime As String
      Public Property FactoryId As String
      Public Property ConfirmCode As String
      Public Property DocTypeCode As String
      <XmlElement(ElementName:="TransactionDetailDataList")>
      Public Property TransactionDetailDataList As clsTransactionDetailDataList

      Public Class clsTransactionDetailDataList
        <XmlElement(ElementName:="TransactionDetailDataInfo")>
        Public Property TransactionDetailDataInfo As List(Of clsTransactionDetailDataInfo)
        Public Class clsTransactionDetailDataInfo
          Public Property SerialId As String
          Public Property SKU As String
          Public Property Qty As String
          Public Property Unit As String
          Public Property TransferOutWarehouse As String
          Public Property TransferInWarehouse As String
          Public Property LotId As String
          Public Property TransferOutShelfId As String
          Public Property TransferInShelfId As String
        End Class
      End Class
    End Class
  End Class
End Class