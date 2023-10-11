Imports System.Xml.Serialization
'Imports eCA_MessageConversion

<XmlRoot(ElementName:="Message")>
Public Class MSG_T5F2U5_BatchCreateReceiptByPO
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String

  Public Class clsBody
    <XmlElement(ElementName:="ReceiptCarrierList")>
    Public Property ReceiptCarrierList As New clsReceiptCarrierList
    Public Class clsReceiptCarrierList
      <XmlElement(ElementName:="ReceiptCarrierInfo")>
      Public Property ReceiptCarrierInfo As New List(Of clsReceiptCarrierInfo)
      Public Class clsReceiptCarrierInfo
        Public Property CARRIER_ID As String = ""
        Public Property CARRIER_TYPE As String = ""
        Public Property CARRIER_MODE As String = ""
        Public Property COMMENTS As String = ""
        Public Property LOCATION_NO As String = ""
        Public Property FACTORY_NO As String = ""
        Public Property DEST_AREA_NO As String = ""
        <XmlElement(ElementName:="ReceiptList")>
        Public Property ReceiptList As New clsReceiptList
        Public Class clsReceiptList
          <XmlElement(ElementName:="ReceiptInfo")>
          Public Property ReceiptInfo As New List(Of clsReceiptInfo)
          Public Class clsReceiptInfo
            Public Property PO_ID As String = ""
            Public Property PO_SERIAL_NO As String = ""
            Public Property ITEM_KEY_NO As String = ""
            Public Property SKU_NO As String = ""
            Public Property PACKAGE_ID As String = ""
            Public Property QTY As String = ""
            Public Property LOT_NO As String = ""
            Public Property ITEM_COMMON1 As String = ""
            Public Property ITEM_COMMON2 As String = ""
            Public Property ITEM_COMMON3 As String = ""
            Public Property ITEM_COMMON4 As String = ""
            Public Property ITEM_COMMON5 As String = ""
            Public Property ITEM_COMMON6 As String = ""
            Public Property ITEM_COMMON7 As String = ""
            Public Property ITEM_COMMON8 As String = ""
            Public Property ITEM_COMMON9 As String = ""
            Public Property ITEM_COMMON10 As String = ""
            Public Property SORT_ITEM_COMMON1 As String = ""
            Public Property SORT_ITEM_COMMON2 As String = ""
            Public Property SORT_ITEM_COMMON3 As String = ""
            Public Property SORT_ITEM_COMMON4 As String = ""
            Public Property SORT_ITEM_COMMON5 As String = ""
            Public Property LENGTH As String = ""
            Public Property WIDTH As String = ""
            Public Property HEIGHT As String = ""
            Public Property WEIGHT As String = ""
            Public Property ITEM_VALUE As String = ""
            Public Property CONTRACT_NO As String = ""
            Public Property CONTRACT_SERIAL_NO As String = ""
            Public Property STORAGE_TYPE As String = ""
            Public Property BND As String = ""
            Public Property QC_STATUS As String = ""
            Public Property RECEIPT_DATE As String = ""
            Public Property MANUFACETURE_DATE As String = ""
            Public Property EXPIRED_DATE As String = ""
            Public Property EFFECTIVE_DATE As String = ""
            Public Property COMMENTS As String = ""
          End Class
        End Class
      End Class
    End Class
  End Class
End Class