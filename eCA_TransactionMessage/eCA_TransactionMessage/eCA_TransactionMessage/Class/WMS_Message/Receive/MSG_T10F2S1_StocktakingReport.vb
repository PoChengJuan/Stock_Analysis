'Vito_20210
'T5F1S3_PODeliveryReport

Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T10F2S1_StocktakingReport
  Public Property Header As clsHeader
  Public Property Body As BodyData
  Public Property KeepData As String

  Public Class BodyData

    <XmlElement(ElementName:="StocktakingList")>
    Public Property StocktakingList As New clsStocktakingList
    Public Class clsStocktakingList
      <XmlElement(ElementName:="StocktakingInfo")>
      Public Property StocktakingInfo As New List(Of clsStocktakingInfo)
      Public Class clsStocktakingInfo
        Public Property STOCKTAKING_ID As String = ""
        'Public Property STOCKTAKING_SERIAL_NO As String = ""
        'Public Property CARRIER_ID As String = ""
        'Public Property LOCATION_NO As String = ""
        'Public Property LOCATION_ID As String = ""
        'Public Property OWNER_NO As String = ""
        'Public Property SUB_OWNER_NO As String = ""
        'Public Property SKU_NO As String = ""
        'Public Property SKU_ID1 As String = ""
        'Public Property SKU_UNIT As String = ""
        'Public Property LOT_NO As String = ""
        'Public Property EXPIRED_DATE As String = ""
        'Public Property QTY As String = ""
        'Public Property REPORT_LOT_NO As String = ""
        'Public Property REPORT_EXPIRED_DATE As String = ""
        'Public Property REPORT_QTY As String = ""
        'Public Property STATUS As String = ""
      End Class
    End Class
  End Class
End Class
