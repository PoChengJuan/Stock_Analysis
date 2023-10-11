'MSG_T5F1U1_PO_Management

Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T5F1U1_PO_Management
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String

  Public Class clsBody
    Public Property Action As String = ""
    Public Property POInfo As New clsPOInfo
    Public Class clsPOInfo
      Public Property PO_ID As String = ""
      Public Property WRITE_OFF_NO As String = ""
      Public Property PO_TYPE1 As String = ""
      Public Property PO_TYPE2 As String = ""
      Public Property PO_TYPE3 As String = ""
      Public Property WO_TYPE As String = ""
      Public Property PRIORITY As String = ""
      Public Property CUSTOMER_NO As String = ""
      Public Property SUPPLIER_NO As String = ""
      Public Property CLASS_NO As String = ""
      Public Property SHIPPING_NO As String = ""
      Public Property H_PO_ORDER_TYPE As String = ""
      Public Property H_PO1 As String = ""
      Public Property H_PO2 As String = ""
      Public Property H_PO3 As String = ""
      Public Property H_PO4 As String = ""
      Public Property H_PO5 As String = ""
      Public Property H_PO6 As String = ""
      Public Property H_PO7 As String = ""
      Public Property H_PO8 As String = ""
      Public Property H_PO9 As String = ""
      Public Property H_PO10 As String = ""
      Public Property H_PO11 As String = ""
      Public Property H_PO12 As String = ""
      Public Property H_PO13 As String = ""
      Public Property H_PO14 As String = ""
      Public Property H_PO15 As String = ""
      Public Property H_PO16 As String = ""
      Public Property H_PO17 As String = ""
      Public Property H_PO18 As String = ""
      Public Property H_PO19 As String = ""
      Public Property H_PO20 As String = ""
      Public Property PO_KEY1 As String = ""
      Public Property PO_KEY2 As String = ""
      Public Property PO_KEY3 As String = ""
      Public Property PO_KEY4 As String = ""
      Public Property PO_KEY5 As String = ""
      <XmlElement(ElementName:="PODetailList")>
      Public Property PODetailList As New clsPODetailList

      Public Class clsPODetailList
        <XmlElement(ElementName:="PODetailInfo")>
        Public Property PODetailInfo As New List(Of clsPODetailInfo)

        Public Class clsPODetailInfo
          Public Property PO_SERIAL_NO As String = ""
          Public Property PO_LINE_NO As String = ""
          Public Property SKU_NO As String = ""
          Public Property LOT_NO As String = ""
          Public Property QTY As String = ""
          Public Property COMMENTS As String = ""
          Public Property PACKAGE_ID As String = ""
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
          Public Property STORAGE_TYPE As String = ""
          Public Property BND As String = ""
          Public Property QC_STATUS As String = ""
          Public Property FROM_OWNER_ID As String = ""
          Public Property FROM_SUB_OWNER_ID As String = ""
          Public Property TO_OWNER_ID As String = ""
          Public Property TO_SUB_OWNER_ID As String = ""
          Public Property FACTORY_ID As String = ""
          Public Property DEST_AREA_ID As String = ""
          Public Property DEST_LOCATION_ID As String = ""
          Public Property H_POD1 As String = ""
          Public Property H_POD2 As String = ""
          Public Property H_POD3 As String = ""
          Public Property H_POD4 As String = ""
          Public Property H_POD5 As String = ""
          Public Property H_POD6 As String = ""
          Public Property H_POD7 As String = ""
          Public Property H_POD8 As String = ""
          Public Property H_POD9 As String = ""
          Public Property H_POD10 As String = ""
          Public Property H_POD11 As String = ""
          Public Property H_POD12 As String = ""
          Public Property H_POD13 As String = ""
          Public Property H_POD14 As String = ""
          Public Property H_POD15 As String = ""
          Public Property H_POD16 As String = ""
          Public Property H_POD17 As String = ""
          Public Property H_POD18 As String = ""
          Public Property H_POD19 As String = ""
          Public Property H_POD20 As String = ""
          Public Property H_POD21 As String = ""
          Public Property H_POD22 As String = ""
          Public Property H_POD23 As String = ""
          Public Property H_POD24 As String = ""
          Public Property H_POD25 As String = ""
          Public Property EXPIRED_DATE As String = ""
        End Class
      End Class
    End Class
  End Class
End Class