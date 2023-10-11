'拉料完成確認後的回傳資訊
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F5U3_LineBigDataAlarmSet
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody

    Public Property Action As String
    <XmlElement(ElementName:="RoleList")>
    Public Property RoleList As New clsRoleList

    Public Class clsRoleList
      <XmlElement(ElementName:="RoleInfo")>
      Public Property RoleInfo As New List(Of clsRoleInfo)
      Public Class clsRoleInfo
        Public Property ROLE_ID As String
        Public Property ROLE_TYPE As String
        Public Property FUNCTION_ID As String
        Public Property DEVICE_NO As String
        Public Property AREA_NO As String
        Public Property UNIT_ID As String
        Public Property HIGH_WATER_VALUE As String
        Public Property LOW_WATER_VALUE As String
        Public Property STANDARD_VALUE As String
        Public Property VALUE_RANGE As String
        Public Property NOTICE_TYPE As String
        Public Property CONTINUE_SEND As String
        Public Property SEND_INTERVAL As String
        Public Property ENABLE As String
      End Class
    End Class
  End Class

End Class

