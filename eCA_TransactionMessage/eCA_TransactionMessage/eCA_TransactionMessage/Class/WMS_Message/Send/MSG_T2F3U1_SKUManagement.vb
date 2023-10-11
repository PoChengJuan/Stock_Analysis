Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T2F3U1_SKUManagement
  Public Property Header As clsHeader
  Public Property Body As SKUDataList
  Public Property KeepData As String

  Public Class SKUDataList

    Public Property Action As String

    Public Property SKUList As SKUDataInfoList

    Public Class SKUDataInfoList
      <XmlElement(ElementName:="SKUInfo")>
      Public Property SKUInfo As New List(Of SKUDataInfo)

      Public Class SKUDataInfo
        Public Property SKU_NO As String
        Public Property SKU_ID1 As String
        Public Property SKU_ID2 As String
        Public Property SKU_ID3 As String
        Public Property SKU_ALIS1 As String
        Public Property SKU_ALIS2 As String
        Public Property SKU_DESC As String
        Public Property SKU_CATALOG As String
        Public Property SKU_TYPE1 As String
        Public Property SKU_TYPE2 As String
        Public Property SKU_TYPE3 As String
        Public Property SKU_COMMON1 As String
        Public Property SKU_COMMON2 As String
        Public Property SKU_COMMON3 As String
        Public Property SKU_COMMON4 As String
        Public Property SKU_COMMON5 As String
        Public Property SKU_COMMON6 As String
        Public Property SKU_COMMON7 As String
        Public Property SKU_COMMON8 As String
        Public Property SKU_COMMON9 As String
        Public Property SKU_COMMON10 As String
        Public Property SKU_L As String
        Public Property SKU_W As String
        Public Property SKU_H As String
        Public Property SKU_WEIGHT As String
        Public Property SKU_VALUE As String
        Public Property SKU_UNIT As String
        Public Property HIGH_WATER As String
        Public Property LOW_WATER As String
        Public Property AVAILABLE_DAYS As String
        Public Property SAVE_DAYS As String
        'Public Property CREATE_TIME As String
        'Public Property UPDATE_TIME As String
        Public Property WEIGHT_DIFFERENCE As String
        Public Property ENABLE As String
        Public Property COMMENTS As String
      End Class
    End Class
  End Class
End Class