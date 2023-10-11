Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_SendOutData
  Public Property WebService_ID As String = "" 'WebService_ID
  Public Property EventID As String = ""       'EventID
  <XmlElement(ElementName:="OutDataList")>
  Public Property OutDataList As New clsOutDataList
  Public Class clsOutDataList
    <XmlElement(ElementName:="OutDataInfo")>
    Public Property OutDataInfo As New List(Of clsOutDataInfo)

    Public Class clsOutDataInfo
      Public Property POType As String = ""
      Public Property PoID As String = ""
      Public Property OutDateTime As String = ""
      Public Property FactoryID As String = ""
      Public Property Owner As String = ""
      <XmlElement(ElementName:="OutDetailDataList")>
      Public Property OutDetailDataList As clsOutDetailDataList

      Public Class clsOutDetailDataList
        <XmlElement(ElementName:="OutDetailDataInfo")>
        Public Property OutDetailDataInfo As List(Of clsOutDetailDataInfo)
        Public Class clsOutDetailDataInfo
          Public Property SerialID As String = ""
          Public Property SKU As String = ""
          Public Property LotId As String = ""
          Public Property CheckQty As String = ""
          Public Property Item_Common1 As String = ""
          Public Property Item_Common2 As String = ""
          Public Property Item_Common3 As String = ""
        End Class
      End Class
    End Class
  End Class
End Class