'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T10F4U1_MainFileImport
    Public Property Header As clsHeader
    Public Property Body As BodyData
  Public Property KeepData As String

  Public Class BodyData
    Public Property Excute As String

    <XmlElement(ElementName:="FileList")>
    Public Property FileList As New FileDataList

    Public Class FileDataList
      <XmlElement(ElementName:="FileInfo")>
      Public Property FileInfo As New List(Of FileDataInfo)
    End Class
  End Class

  Public Class FileDataInfo
    Public Property MainFileType As String
    Public Property FileType As String
    Public Property FilePath As String
  End Class

End Class

