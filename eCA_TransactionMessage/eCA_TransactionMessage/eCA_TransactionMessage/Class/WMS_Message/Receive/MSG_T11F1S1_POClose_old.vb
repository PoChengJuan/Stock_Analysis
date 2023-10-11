'MSG_T11F1U1_PODownload

Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F1S1_POClose_old
  Public Property Header As clsHeader
  Public Property Body As BodyData
  Public Property KeepData As String

  Public Class BodyData
    Public Property Forced_Close As String
    Public Property WO_ID As String
    <XmlElement(ElementName:="POList")>
    Public Property POList As New clsPOList
    Public Class clsPOList
      <XmlElement(ElementName:="POInfo")>
      Public Property POInfo As New List(Of clsPOInfo)
      Public Class clsPOInfo
        Public Property PO_ID As String
        Public Property H_PO_ORDER_TYPE As String
        <XmlElement(ElementName:="PO_DTLList")>
        Public Property PO_DTLList As New clsPO_DTLList
        Public Class clsPO_DTLList
          <XmlElement(ElementName:="PO_DTLInfo")>
          Public Property PO_DTLInfo As New List(Of clsPO_DTLInfo)
          Public Class clsPO_DTLInfo
            Public Property PO_LINE_NO As String
            Public Property PO_SERIAL_NO As String
            Public Property SKU_NO As String
            Public Property SORT_ITEM_COMMON1 As String
            Public Property SORT_ITEM_COMMON2 As String
            Public Property SORT_ITEM_COMMON3 As String
            Public Property SORT_ITEM_COMMON4 As String
            Public Property SORT_ITEM_COMMON5 As String
            Public Property QTY As String
            <XmlElement(ElementName:="TextList")>
            Public Property TextList As New clsTextList
            Public Class clsTextList
              <XmlElement(ElementName:="TextInfo")>
              Public Property TextInfo As New List(Of clsTextInfo)
              Public Class clsTextInfo
                Public Property Name As String
                Public Property Value As String
              End Class
            End Class
          End Class
        End Class
      End Class
    End Class
  End Class
End Class