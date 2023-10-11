Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T10F2U1_StocktakingManagement
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody
    Public Property Action As String
    <XmlElement(ElementName:="StocktakingInfo")>
    Public Property StocktakingInfo As New clsStocktakingInfo

    Public Class clsStocktakingInfo
      Public Property STOCKTAKING_ID As String
      Public Property LOCATION_GROUP_NO As String
      Public Property PRIORITY As String
      Public Property STOCKTAKING_TYPE1 As String
      Public Property STOCKTAKING_TYPE2 As String
      Public Property STOCKTAKING_TYPE3 As String
      Public Property SEND_TO_HOST As String
      Public Property CHANGE_INVENTORY As String

      <XmlElement(ElementName:="StocktakingDTLList")>
      Public Property StocktakingDTLList As New clsStocktakingDTLList

      Public Class clsStocktakingDTLList
        <XmlElement(ElementName:="StocktakingDTLInfo")>
        Public Property StocktakingDTLInfo As New List(Of clsStocktakingDTLInfo)

        Public Class clsStocktakingDTLInfo
          Public Property STOCKTAKING_SERIAL_NO As String
          Public Property AREA_NO As String
          Public Property BLOCK_NO As String
          Public Property SKU_NO As String
          Public Property BND As String
          Public Property SL_NO As String
          Public Property CARRIER_ID As String
          Public Property PERCENTAGE As String
          Public Property LOT_NO As String
          Public Property STORAGE_TYPE As String
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
          Public Property OWNER_NO As String
          Public Property SUB_OWNER_NO As String
          Public Property SUPPLIER_NO As String
          Public Property CUSTOMER_NO As String
          Public Property RECEIPT_DATE As String
          Public Property MANUFACETURE_DATE As String
          Public Property EXPIRED_DATE As String
          Public Property ERP_QTY As Double
        End Class
      End Class
    End Class
  End Class
End Class