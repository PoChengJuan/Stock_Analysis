﻿Imports System.Xml.Serialization

Public Class clsHeader
  Public Property UUID As String = ""
  Public Property EventID As String = ""
  Public Property Direction As String = ""
  Public Property SystemID As String = ""

  <XmlElement(ElementName:="ClientInfo")>
  Public Property ClientInfo As clsClientInfo

  Public Class clsClientInfo
    Public Property ClientID As String
    Public Property UserID As String
    Public Property IP As String
    Public Property MachineID As String
  End Class

End Class

Public Class clsPOHeaderInfo
  Public Property WebService_ID As String = ""
  Public Property EventID As String = ""
End Class

Public Class KeepData
  Public Property KeepData As String
End Class