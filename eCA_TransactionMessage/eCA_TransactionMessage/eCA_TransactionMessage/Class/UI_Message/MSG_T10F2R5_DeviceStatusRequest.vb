Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T10F2R5_DeviceStatusRequest
  Public Property Header As HeaderDataInfo
  Public Property Body As DeviceDataList
  Public Property KeepData As String

  Public Class DeviceDataList
    Inherits BodyInfo

    Public Property DeviceList As DeviceDataInfoList


    Public Class DeviceDataInfoList
      Inherits List(Of DeviceInfo)


      Public Class DeviceInfo
        Public Property DeviceID As String
      End Class
    End Class
  End Class


End Class
