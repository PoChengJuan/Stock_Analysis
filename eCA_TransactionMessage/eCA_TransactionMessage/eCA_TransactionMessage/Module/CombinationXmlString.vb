'20180814
'V1.0.0
'Mark
'把要接收的Message，從Object轉成Xml String

Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Newtonsoft.Json

Public Module CombinationXmlString

  Public Function PrepareMessage_SendTransferDataToERP(ByRef strXML As String, ByRef objMSG_SendTransferDataToERP As MSG_SendTransferDataToERP, ByRef RetMsg As String) As Boolean
    Try
      strXML = ParseClassToXmlString(Of MSG_SendTransferDataToERP)(objMSG_SendTransferDataToERP)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function PrepareMessage_InventoryToERP(ByRef strXML As String, ByRef objMsg_InventoryData As MSG_InventoryData, ByRef RetMsg As String) As Boolean
    Try
      strXML = ParseClassToXmlString(Of MSG_InventoryData)(objMsg_InventoryData)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

  Public Function PrepareMessage_ERPReport(ByRef strXML As String, ByRef objMSG_ERPReport As MSG_ERPReport, ByRef RetMsg As String) As Boolean
    Try
      strXML = ParseClassToXmlString(Of MSG_ERPReport)(objMSG_ERPReport)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

  Public Function PrepareMessage_T3F5S1_LineProductionResetRequest(ByRef strXML As String, ByRef objMSG As MSG_T3F5S1_LineProductionResetRequest, ByRef RetMsg As String) As Boolean
    Try
      strXML = ParseClassToXmlString(Of MSG_T3F5S1_LineProductionResetRequest)(objMSG)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function PrepareMessage_T_objMSG(Of T)(ByRef strXML As String, ByRef objMSG As Object, ByRef RetMsg As String) As Boolean
    Try
      strXML = ParseClassToXmlString(Of T)(objMSG)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

  Private Function ParseClassToXmlString(Of T)(Clsaa As Object) As String
    'Dim xs As New System.Xml.Serialization.XmlSerializer(Clsaa.GetType)
    'Dim w As New IO.StringWriter
    'xs.Serialize(w, Clsaa)

    'Return w.ToString
    '去掉xml声明 '20180925 新增
    Dim settings As XmlWriterSettings = New XmlWriterSettings()
    settings.OmitXmlDeclaration = True
    settings.Encoding = Encoding.UTF8
    Dim mem As System.IO.MemoryStream = New MemoryStream()
    Using writer As XmlWriter = XmlWriter.Create(mem, settings)
      '去除默认命名空间xmlns: xsd和xmlns : xsi
      Dim nss As XmlSerializerNamespaces = New XmlSerializerNamespaces()
      nss.Add("", "")
      Dim formatter As XmlSerializer = New XmlSerializer(Clsaa.GetType())
      formatter.Serialize(writer, Clsaa, nss)
    End Using
    Return Encoding.UTF8.GetString(mem.ToArray()).Replace("﻿", "")
  End Function

End Module
