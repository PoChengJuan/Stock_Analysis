Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendSKUChangeData
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="SKUChangeDataList")>
  Public Property SKUChangeDataList As New clsSKUChangeDataList
  Public Class clsSKUChangeDataList
    <XmlElement(ElementName:="SKUChangeDataInfo")>
    Public Property SKUChangeDataInfo As New List(Of clsSKUChangeDataInfo)
    Public Class clsSKUChangeDataInfo
      Public Property SKU As String
      Public Property ChangeEdition As String
      Public Property ChangeDateTime As String
      Public Property NewSKU As String
      Public Property NewSpecification As String
      Public Property ConfirmCode As String
      Public Property FieldID As String
      Public Property NewStringFieldValue As String
      Public Property NewIntFieldValue As String
    End Class
  End Class
End Class
