'保養參數設定
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F5U1_MaintenanceSet
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody

    Public Property Action As String
    <XmlElement(ElementName:="MaintenanceList")>
    Public Property MaintenanceList As New clsMaintenanceList

    Public Class clsMaintenanceList
      <XmlElement(ElementName:="MaintenanceInfo")>
      Public Property MaintenanceInfo As New List(Of clsMaintenanceInfo)

      Public Class clsMaintenanceInfo
        Public Property FACTORY_NO As String
        Public Property DEVICE_NO As String
        Public Property AREA_NO As String
        Public Property UNIT_ID As String
        Public Property MAINTENANCE_ID As String
        Public Property MAINTENANCE_NAME As String
        Public Property CONTINUE_SEND As String
        Public Property SEND_INTERVAL As String
        Public Property SEND_TYPE As String
        Public Property ENABLE As String
        <XmlElement(ElementName:="MaintenanceDTLList")>
        Public Property MaintenanceDTLList As New clsMaintenanceDTLList

        Public Class clsMaintenanceDTLList
          <XmlElement(ElementName:="MaintenanceDTLInfo")>
          Public Property MaintenanceDTLInfo As New List(Of clsMaintenanceDTLInfo)

          Public Class clsMaintenanceDTLInfo
            Public Property FUNCTION_ID As String
            Public Property VALUE_TYPE As String
            Public Property NOTICE_TYPE As String
            Public Property HIGH_WATER_VALUE As String
            Public Property LOW_WATER_VALUE As String
            Public Property STANDARD_VALUE As String
            Public Property VALUE_RANGE As String
            Public Property MAINTENANCE_MESSAGE As String
            Public Property VALUE_SOURCE As String
            Public Property VALUE_UPDATE_TYPE As String
          End Class
        End Class
      End Class
    End Class
  End Class
End Class
