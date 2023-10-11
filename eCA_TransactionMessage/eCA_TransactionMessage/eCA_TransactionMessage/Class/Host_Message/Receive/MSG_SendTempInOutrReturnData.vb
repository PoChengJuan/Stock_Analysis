Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendTempInOutrReturnData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="TempInOutDataList")>
  Public Property TempInOutDataList As New clsTempInOutDataList
  Public Class clsTempInOutDataList
    <XmlElement(ElementName:="TempInOutDataInfo")>
    Public Property TempInOutDataInfo As New List(Of clsTempInOutDataInfo)
    Public Class clsTempInOutDataInfo
      Public Property POType As String
      Public Property POId As String
      Public Property TempInOutDateTime As String
      Public Property FactoryId As String
      Public Property ConfirmCode As String
      <XmlElement(ElementName:="TempInOutDetailDataList")>
      Public Property TempInOutDetailDataList As clsTempInOutDetailDataList

      Public Class clsTempInOutDetailDataList
        <XmlElement(ElementName:="TempInOutDetailDataInfo")>
        Public Property TempInOutDetailDataInfo As List(Of clsTempInOutDetailDataInfo)
        Public Class clsTempInOutDetailDataInfo
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