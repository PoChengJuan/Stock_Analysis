Imports System.Xml.Serialization

<XmlRoot(ElementName:="eWMSMessage")>
Public Class MSG_eWMSMessage
  Public Property WebService_ID As String 'WebService_ID
  Public Property EventID As String       'EventID
  <XmlElement(ElementName:="NoticeDataList")>
  Public Property NoticeDataList As New clsNoticeDataList

  Public Class clsNoticeDataList
    <XmlElement(ElementName:="NoticeDataInfo")>
    Public Property NoticeDataInfo As New List(Of clsNoticeDatainfo)

    Public Class clsNoticeDatainfo
      Public Property NoticeType As String
      Public Property NoticeId As String
      Public Property NoticeDate As String
      Public Property NoticeTime As String
      Public Property Spare1 As String
      Public Property Spare2 As String
      Public Property Spare3 As String
      Public Property Spare4 As String
      Public Property Spare5 As String
      <XmlElement(ElementName:="NoticeDetailDataList")>
      Public Property NoticeDetailDataList As clsNoticeDetailDataList

      Public Class clsNoticeDetailDataList
        <XmlElement(ElementName:="NoticeDetailDataInfo")>
        Public Property NoticeDetailDataInfo As List(Of clsNoticeDetailDataInfo)
        Public Class clsNoticeDetailDataInfo
          Public Property NoticeSerialId As String
          Public Property TransferType As String
          Public Property SKU As String
          Public Property QTY As String
          Public Property Unit As String
          Public Property WH As String
          Public Property Location As String
          Public Property LotId As String
          Public Property SKUName As String
          Public Property WEIGHT As String
          Public Property LENGTH As String
          Public Property WIDTH As String
          Public Property WO As String
          Public Property GRADE As String
          Public Property Spare1 As String
          Public Property Spare2 As String
          Public Property Spare3 As String
          Public Property Spare4 As String
          Public Property Spare5 As String
        End Class
      End Class
    End Class
  End Class
End Class