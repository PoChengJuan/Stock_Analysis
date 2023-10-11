Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendSKUData
  Public Property identity As clsidentity
  Public Property WebService_ID As String = "" 'WebService_ID
  Public Property EventID As String = ""       'EventID
  <XmlElement(ElementName:="SKUDataList")>
  Public Property SKUDataList As New clsSKUDataList
  Public Class clsSKUDataList
    <XmlElement(ElementName:="SKUDataInfo")>
    Public Property SKUDataInfo As New List(Of clsSKUDataInfo)
    Public Class clsSKUDataInfo
      Public Property SKU As String = ""
      Public Property SKUName As String = ""
      Public Property Specification As String = ""
      Public Property InventoryUnit As String = ""
      Public Property ASRSPart As String = ""
      Public Property SKUType1 As String = ""
      Public Property SKUType2 As String = ""
      Public Property SKUType3 As String = ""
      Public Property SKUType4 As String = ""
      Public Property ProductDescription As String = ""
      Public Property EffectiveDateTime As String = ""
      Public Property FailureDateTime As String = ""
      Public Property Comment As String = ""

    End Class
  End Class
End Class
