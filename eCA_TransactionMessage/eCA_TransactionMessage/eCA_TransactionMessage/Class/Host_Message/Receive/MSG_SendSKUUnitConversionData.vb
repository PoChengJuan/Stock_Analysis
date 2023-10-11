Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendSKUUnitConversionData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="SKUUnitConversionDataList")>
  Public Property SKUUnitConversionDataList As New clsSKUUnitConversionDataList
  Public Class clsSKUUnitConversionDataList
    <XmlElement(ElementName:="SKUUnitConversionDataInfo")>
    Public Property SKUUnitConversionDataInfo As New List(Of clsSKUUnitConversionDataInfo)
    Public Class clsSKUUnitConversionDataInfo
      Public Property SKU As String = ""
      Public Property InventoryUnit As String = ""
      Public Property ConversionUnit As String = ""
      Public Property Molecule As String = ""
      Public Property Denominator As String = ""
    End Class
  End Class
End Class