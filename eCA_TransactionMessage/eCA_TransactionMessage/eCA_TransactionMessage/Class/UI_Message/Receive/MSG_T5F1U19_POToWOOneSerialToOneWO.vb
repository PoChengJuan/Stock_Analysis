Imports System.Xml.Serialization
'Imports eCA_MessageConversion

<XmlRoot(ElementName:="Message")>
Public Class MSG_T5F1U19_POToWOOneSerialToOneWO
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String

  Public Class clsBody
    <XmlElement(ElementName:="POList")>
    Public Property POList As New clsPOList
    <XmlElement(ElementName:="WOInfo")>
    Public Property WOInfo As New clsWOInfo

    Public Class clsPOList
      <XmlElement(ElementName:="POInfo")>
      Public Property POInfo As New List(Of clsPOInfo)
      Public Class clsPOInfo
        Public Property PO_ID As String
        Public Property PO_SERIAL_NO As String
        Public Property SKU_NO As String
        Public Property QTY As String
        Public Property LOT_NO As String
        Public Property PACKAGE_ID As String
        Public Property ITEM_COMMON1 As String
        Public Property ITEM_COMMON2 As String
        Public Property ITEM_COMMON3 As String
        Public Property ITEM_COMMON4 As String
        Public Property ITEM_COMMON5 As String
        Public Property ITEM_COMMON6 As String
        Public Property ITEM_COMMON7 As String
        Public Property ITEM_COMMON8 As String
        Public Property ITEM_COMMON9 As String
        Public Property ITEM_COMMON10 As String
        Public Property SORT_ITEM_COMMON1 As String
        Public Property SORT_ITEM_COMMON2 As String
        Public Property SORT_ITEM_COMMON3 As String
        Public Property SORT_ITEM_COMMON4 As String
        Public Property SORT_ITEM_COMMON5 As String
        Public Property FROM_OWNER_NO As String
        Public Property FROM_SUB_OWNER_NO As String
        Public Property TO_OWNER_NO As String
        Public Property TO_SUB_OWNER_NO As String
        Public Property FACTORY_NO As String
        Public Property DEST_AREA_NO As String
        Public Property DEST_LOCATION_NO As String
      End Class
    End Class
    Public Class clsWOInfo
      Public Property RECEIPT_ENABLE As String = ""
      Public Property COMMENTS As String = ""
      Public Property SOURCE_AREA_NO As String = ""
      Public Property SOURCE_LOCATION_NO As String = ""
      Public Property SHIPPING_NO As String = ""
      Public Property SHIPPING_PRIORITY As String = ""
    End Class
  End Class
End Class
