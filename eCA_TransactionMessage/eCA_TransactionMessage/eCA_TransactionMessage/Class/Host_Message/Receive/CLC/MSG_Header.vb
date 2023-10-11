Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_Header
  Public Property identity As clsidentity
  Public Property WebService_ID As String = "" 'WebService_ID
  Public Property EventID As String = "" 'EventID
End Class

Public Class clsidentity
  Public Property TransactionID As String = ""
  Public Property PlantID As String = ""
  Public Property ProgramID As String = ""
  Public Property TableName As String = ""
  Public Property Hostname As String = ""
  Public Property UserID As String = ""
  Public Property SendTime As String = ""
End Class