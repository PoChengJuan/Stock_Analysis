'保養參數設定
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T3F5U5_ClassProductionSet
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody

    Public Property Action As String
    <XmlElement(ElementName:="ClassList")>
    Public Property ClassList As New clsClassList

    Public Class clsClassList
      <XmlElement(ElementName:="ClassInfo")>
      Public Property ClassInfo As New List(Of clsClassInfo)

      Public Class clsClassInfo
        Public Property CLASS_NO As String
        Public Property ATTENDANCE_COUNT As String
        <XmlElement(ElementName:="AssignationList")>
        Public Property AssignationList As New clsAssignationList

        Public Class clsAssignationList
          <XmlElement(ElementName:="AssignationInfo")>
          Public Property AssignationInfo As New List(Of clsAssignationInfo)

          Public Class clsAssignationInfo
            Public Property FACTORY_NO As String
            Public Property AREA_NO As String
            Public Property ASSIGNATION_RATE As String
          End Class
        End Class
      End Class
    End Class
  End Class
End Class
