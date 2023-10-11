Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F4R2_DeviceAlarmReport
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String
  Public Class clsBody
    Public Property AlarmList As clsAlarmList
    Public Class clsAlarmList
      <XmlElement(ElementName:="AlarmInfo")>
      Public Property AlarmInfo As New List(Of clsAlarmInfo)
      Public Class clsAlarmInfo
        Public Property ALARM_CODE As String
        Public Property ALARM_DESC As String
        Public Property COMMAND_ID As String
        Public Property FACTORY_NO As String
        Public Property DEVICE_NO As String
        Public Property UNIT_ID As String
        Public Property TIME As String
        Public Property SET_FLAG As String
      End Class
    End Class
  End Class
End Class