Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Newtonsoft.Json

''' <summary>
''' 20180814
''' V1.0.0
''' Mark
''' 把要傳送的Message，從Xml String轉成Object
''' </summary>
Public Module ParseXmlString

  Public Function ParseMessage_MSG_Outbound_Info(ByVal strJSON As String,
                                                 ByRef objMSG_Outbound_Info As List(Of MSG_Outbound_Info),
                                                 ByRef RetMsg As String) As Boolean
    Try
      objMSG_Outbound_Info = JsonConvert.DeserializeObject(Of List(Of MSG_Outbound_Info))(strJSON)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function PrepareMessage_MSG_PostingCheck(ByRef strJSON As String,
                                                  ByRef objMSG_PostingCheck As MSG_PostingCheck,
                                                  ByRef RetMsg As String) As Boolean
    Try
      strJSON = JsonConvert.SerializeObject(objMSG_PostingCheck)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

#Region "GUI PrimaryIn"
  Public Function ParseMessage_T11F1U11_ProducePOExecution(ByVal strXML As String,
                                                            ByRef objMSG As MSG_T11F1U11_ProducePOExecution,
                                                            ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T11F1U11_ProducePOExecution)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F1U4_PODownload(ByVal strXML As String,
                                                            ByRef objMSG As MSG_T5F1U4_PODownload,
                                                            ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T5F1U4_PODownload)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T11F1U1_PODownload(ByVal strXML As String,
                                                            ByRef objMSG As MSG_T11F1U1_PODownload,
                                                            ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T11F1U1_PODownload)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T11F2U1_InventoryComparison(ByVal strXML As String,
                                                            ByRef objMSG As MSG_T11F2U1_InventoryComparison,
                                                            ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T11F2U1_InventoryComparison)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T11F1U12_StocktakingExecution(ByVal strXML As String,
                                                            ByRef objMSG As MSG_T11F1U12_StocktakingExecution,
                                                            ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T11F1U12_StocktakingExecution)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T11F1U2_POExecution(ByVal strXML As String,
                                                        ByRef objMSG As MSG_T11F1U2_POExecution,
                                                        ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T11F1U2_POExecution)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F1U11_POExecution(ByVal strXML As String,
                                                        ByRef objMSG As MSG_T5F1U11_POExecution,
                                                        ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T5F1U11_POExecution)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F1U18_POToWOOneToOne(ByVal strXML As String,
                                                        ByRef objMSG As MSG_T5F1U18_POToWOOneToOne,
                                                        ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T5F1U18_POToWOOneToOne)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F1U19_POToWOOneSerialToOneWO(ByVal strXML As String,
                                                        ByRef objMSG As MSG_T5F1U19_POToWOOneSerialToOneWO,
                                                        ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T5F1U19_POToWOOneSerialToOneWO)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

  Public Function ParseMessage_T10F4U1_MainFileImport(ByVal strXML As String, ByRef obj As MSG_T10F4U1_MainFileImport, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T10F4U1_MainFileImport)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T6F5U1_ItemLabelManagement(ByVal strXML As String, ByRef obj As MSG_T6F5U1_ItemLabelManagement, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T6F5U1_ItemLabelManagement)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T6F5U2_ItemLabelPrint(ByVal strXML As String, ByRef obj As MSG_T6F5U2_ItemLabelPrint, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T6F5U2_ItemLabelPrint)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T10F2S1_StocktakingReport(ByVal strXML As String, ByRef obj As MSG_T10F2S1_StocktakingReport, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T10F2S1_StocktakingReport)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T11F3U1_StocktakingDownload(ByVal strXML As String, ByRef obj As MSG_T11F3U1_StocktakingDownload, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T11F3U1_StocktakingDownload)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5U1_MaintenanceSet(ByVal strXML As String,
                                                                                                         ByRef objMSG As MSG_T3F5U1_MaintenanceSet,
                                                                                                         ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5U1_MaintenanceSet)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5U2_Maintenance(ByVal strXML As String, ByRef objMSG As MSG_T3F5U2_Maintenance, ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5U2_Maintenance)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5U3_LineBigDataAlarmSet(ByVal strXML As String, ByRef objMSG As MSG_T3F5U3_LineBigDataAlarmSet, ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5U3_LineBigDataAlarmSet)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5U4_ProductionCountSet(ByVal strXML As String, ByRef objMSG As MSG_T3F5U4_ProductionCountSet, ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5U4_ProductionCountSet)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5U5_ClassProductionSet(ByVal strXML As String, ByRef objMSG As MSG_T3F5U5_ClassProductionSet, ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5U5_ClassProductionSet)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function




#End Region
#Region "MCS PrimaryIn"
  Public Function ParseMessage_T3F4R2_DeviceAlarmReport(ByVal strXML As String,
                                                        ByRef objMSG As MSG_T3F4R2_DeviceAlarmReport,
                                                        ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F4R2_DeviceAlarmReport)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5R1_LineStatusChangeReport(ByVal strXML As String,
                                                             ByRef objMSG As MSG_T3F5R1_LineStatusChangeReport,
                                                             ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5R1_LineStatusChangeReport)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5R2_LineInfoReport(ByVal strXML As String,
                                                     ByRef objMSG As MSG_T3F5R2_LineInfoReport,
                                                     ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5R2_LineInfoReport)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5R3_LineInProductionInfoReport(ByVal strXML As String,
                                                                 ByRef objMSG As MSG_T3F5R3_LineInProductionInfoReport,
                                                                 ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5R3_LineInProductionInfoReport)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T3F5R4_LineInProductionInfoReset(ByVal strXML As String,
                                                                ByRef objMSG As MSG_T3F5R4_LineInProductionInfoReset,
                                                                ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T3F5R4_LineInProductionInfoReset)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
#End Region

#Region "WMS Secondary In"
  Public Function ParseMessage_T5F3U23_POToWO(ByVal strXML As String,
                                              ByRef objMSG As MSG_T5F3U23_POToWO,
                                              ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T5F3U23_POToWO)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F1U1_POManagement(ByVal strXML As String,
                                              ByRef objMSG As MSG_T5F1U1_PO_Management,
                                              ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T5F1U1_PO_Management)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F5U1_TransactionOederManagement(ByVal strXML As String,
                                              ByRef objMSG As MSG_T5F5U1_TransactionOederManagement,
                                              ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T5F5U1_TransactionOederManagement)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T2F3U1_SKUManagement(ByVal strXML As String,
                                              ByRef objMSG As MSG_T2F3U1_SKUManagement,
                                              ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T2F3U1_SKUManagement)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F2U62_AutoInbound(ByVal strXML As String,
                                              ByRef objMSG As MSG_T5F2U62_AutoInbound,
                                              ByRef RetMsg As String) As Boolean
    Try
      objMSG = ParseXmlStringToClass(Of MSG_T5F2U62_AutoInbound)(strXML.ToString, RetMsg)
      If objMSG Is Nothing Then Return False
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
#End Region



  Public Function ParseMessage_T11F1S1_POClose(ByVal strXML As String, ByRef obj As MSG_T11F1S1_POClose, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T11F1S1_POClose)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F1S1_WOClose(ByVal strXML As String, ByRef obj As MSG_T5F1S1_WOClose, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T5F1S1_WOClose)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ParseMessage_T5F1S13_TransationPOExecution(ByVal strXML As String, ByRef obj As MSG_T5F1S13_TransationPOExecution, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T5F1S13_TransationPOExecution)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

  Public Function ParseMessage_T3F5S1_LineProductionResetRequest(ByVal strXML As String, ByRef obj As MSG_T3F5S1_LineProductionResetRequest, ByRef RetMsg As String) As Boolean
    Try
      obj = ParseXmlStringToClass(Of MSG_T3F5S1_LineProductionResetRequest)(strXML.ToString, RetMsg)
      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

  Public Function PrepareMessage_MSG(Of T)(ByRef strXML As String, ByRef obj As Object, ByRef RetMsg As String) As Boolean
    Try
      Dim XML = ParseClassToXmlString(Of T)(obj)

      strXML = XML.Replace(vbCrLf, "")

      strXML = strXML.Replace("  ", " ")


      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

  '轉XML給WMS
  Public Function PrepareMessage_T5F3U3_POToWO(ByRef strXML As String, ByRef objMSG_T5F3U3_POToWO As MSG_T5F3U23_POToWO, ByRef RetMsg As String) As Boolean
    Try
      Dim XML = ParseClassToXmlString(Of MSG_T5F3U23_POToWO)(objMSG_T5F3U3_POToWO)

      strXML = XML.Replace(vbCrLf, "")

      strXML = strXML.Replace("  ", " ")


      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function PrepareMessage_T5F1U12_POToWO(ByRef strXML As String, ByRef objMSG_T5F3U3_POToWO As MSG_T5F1U12_POToWO, ByRef RetMsg As String) As Boolean
    Try
      Dim XML = ParseClassToXmlString(Of MSG_T5F1U12_POToWO)(objMSG_T5F3U3_POToWO)

      strXML = XML.Replace(vbCrLf, "")

      strXML = strXML.Replace("  ", " ")


      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function PrepareMessage_T10F2U2_StocktakingExecute(ByRef strXML As String, ByRef objMSG As MSG_T10F2U2_StocktakingExecute, ByRef RetMsg As String) As Boolean
    Try
      Dim XML = ParseClassToXmlString(Of MSG_T10F2U2_StocktakingExecute)(objMSG)

      strXML = XML.Replace(vbCrLf, "")

      strXML = strXML.Replace("  ", " ")


      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
  '轉XML給WMS
  Public Function PrepareMessage_T10F2U1_StocktakingManagement(ByRef strXML As String, ByRef obj As MSG_T10F2U1_StocktakingManagement, ByRef RetMsg As String) As Boolean
    Try
      Dim XML = ParseClassToXmlString(Of MSG_T10F2U1_StocktakingManagement)(obj)

      strXML = XML.Replace(vbCrLf, "")

      strXML = strXML.Replace("  ", " ")


      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function

  Private Function ParseXmlStringToClass(Of T)(_XmlString As String, ByRef RetMsg As String) As T
    Try
      Return New XmlSerializer(GetType(T)).Deserialize(New MemoryStream(Encoding.Unicode.GetBytes(_XmlString)))
      'Return New XmlSerializer(GetType(T)).Deserialize(New MemoryStream(Encoding.UTF8.GetBytes(_XmlString)))
    Catch ex As Exception
      RetMsg = ex.ToString
      Return Nothing
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
      '去除默认命名空间xmlns:xsd和xmlns:xsi
      Dim nss As XmlSerializerNamespaces = New XmlSerializerNamespaces()
      nss.Add("", "")
      Dim formatter As XmlSerializer = New XmlSerializer(Clsaa.GetType())
      formatter.Serialize(writer, Clsaa, nss)
    End Using
    Return Encoding.UTF8.GetString(mem.ToArray()).Replace("﻿", "")
  End Function

  Public Function PrepareMessage_Secondary_Message(ByRef strXML As String, ByRef obj As MSG_Secondary_Message, ByRef RetMsg As String) As Boolean
    Try
      Dim XML = ParseClassToXmlString(Of MSG_Secondary_Message)(obj)

      strXML = XML.Replace(vbCrLf, "")

      strXML = strXML.Replace("  ", " ")

      Return True
    Catch ex As Exception
      RetMsg = ex.ToString()
      Return False
    End Try
  End Function
End Module
