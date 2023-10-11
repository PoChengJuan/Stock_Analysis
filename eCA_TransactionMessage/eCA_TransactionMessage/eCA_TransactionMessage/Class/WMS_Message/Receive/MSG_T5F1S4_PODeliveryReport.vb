'Vito_20210
'T5F1S3_PODeliveryReport

Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T5F1S4_PODeliveryReport
  Public Property Header As clsHeader
  Public Property Body As BodyData
  Public Property KeepData As String

  Public Class BodyData
    <XmlElement(ElementName:="DeliveryInfo")>
    Public Property DeliveryInfo As New clsDeliveryInfo
    Public Class clsDeliveryInfo
      Public Property WO_ID As String = ""
      Public Property OWNER_NO As String = ""
      Public Property SUB_OWNER_NO As String = ""
      Public Property SHIPPING_COMPANY_NO As String = ""
      Public Property SHIPPING_UID As String = ""
      Public Property COMMON1 As String = ""
      Public Property COMMON2 As String = ""
      Public Property COMMON3 As String = ""
      Public Property COMMON4 As String = ""
      Public Property COMMON5 As String = ""
      Public Property COMMON6 As String = ""
      Public Property COMMON7 As String = ""
      Public Property COMMON8 As String = ""
      Public Property COMMON9 As String = ""
      Public Property COMMON10 As String = ""
      Public Property COMMENTS As String = ""
      Public Property ESTIMATE_CARRIER_QTY As String = ""
      Public Property CARRIER_QTY As String = ""
      <XmlElement(ElementName:="DeliveryDTLList")>
      Public Property DeliveryDTLList As New clsDeliveryDTLList
      Public Class clsDeliveryDTLList
        <XmlElement(ElementName:="DeliveryDTLInfo")>
        Public Property DeliveryDTLInfo As New List(Of clsDeliveryDTLInfo)
        Public Class clsDeliveryDTLInfo
          Public Property CARRIER_ID As String = ""
          Public Property CARRIER_LABEL_ID As String = ""
          Public Property CARRIER_KIND As String = ""
          Public Property PO_ID As String = ""
          Public Property DELIVERY_TIME As String = ""
          <XmlElement(ElementName:="ItemList")>
          Public Property ItemList As New clsItemList
          Public Class clsItemList
            <XmlElement(ElementName:="ItemInfo")>
            Public Property ItemInfo As New List(Of clsItemInfo)
            Public Class clsItemInfo
              Public Property SKU_KIND As String = ""
              Public Property SKU_NO As String = ""
              Public Property SKU_ID1 As String = ""
              Public Property QTY As String = ""
              Public Property LOT_NO As String = ""
              Public Property PO_SERIAL_NO As String = ""
              Public Property ITEM_COMMON1 As String = ""
              Public Property ITEM_COMMON2 As String = ""
              Public Property ITEM_COMMON3 As String = ""
              Public Property ITEM_COMMON4 As String = ""
            End Class
          End Class
        End Class
      End Class
    End Class
  End Class
End Class