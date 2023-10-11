Imports System.Xml.Serialization

<XmlRoot(ElementName:="lstOUTBOUND")>
Public Class MSG_ASRS_getCompStock

  <XmlElement(ElementName:="OUT_TAB")>
  Public Property output As New List(Of clsoutput)

  Public Class clsoutput
    Public Property MATNR As String 'Material Number

    Public Property CPUDT As String 'TransactionDate

    Public Property CLABS As String 'Quantity

    Public Property LGORT As String 'Storage Location

    Public Property MANDT As String '無說明

    Public Property CHARG As String 'Batch Number

    Public Property WERKS As String 'Plant

    Public Property LABST As String 'Plant

    Public Property MENGE As String 'Plant

  End Class
End Class