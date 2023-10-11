'20180901
'V1.0.0
'Jerry
Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_Send_WMSMessage

  Public Function Send_T5F1U25_POUpdate_to_WMS(ByRef Result_Message As String, ByVal dicPO_ID As Dictionary(Of String, String),
                            ByRef Host_Command As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try

      Dim EventID = "T5F1U16_POUpdate"

      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)
      Dim UUID = objUUID.Get_NewUUID


      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim msg_POUpdate As New MSG_T5F1U16_POUpdate
      msg_POUpdate.Header = New clsHeader
      msg_POUpdate.Header.UUID = UUID
      msg_POUpdate.Header.EventID = EventID
      msg_POUpdate.Header.Direction = "Primary"
      msg_POUpdate.Header.ClientInfo = New clsHeader.clsClientInfo
      msg_POUpdate.Header.ClientInfo.ClientID = "Handler"
      msg_POUpdate.Header.ClientInfo.UserID = ""
      msg_POUpdate.Header.ClientInfo.IP = ""
      msg_POUpdate.Header.ClientInfo.MachineID = ""
      msg_POUpdate.Body = New MSG_T5F1U16_POUpdate.clsBody


      Dim PO_ID = ""
      Dim POList As New MSG_T5F1U16_POUpdate.clsBody.clsPOList
      POList.POInfo = New List(Of MSG_T5F1U16_POUpdate.clsBody.clsPOList.clsPOInfo)
      For Each objPO_ID In dicPO_ID.Values
        Dim lstPOInfo As New MSG_T5F1U16_POUpdate.clsBody.clsPOList.clsPOInfo
        lstPOInfo.PO_ID = objPO_ID
        POList.POInfo.Add(lstPOInfo)
      Next

      msg_POUpdate.Body.POList = POList '資料填寫完成

      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_MSG(Of MSG_T5F1U16_POUpdate)(strXML, msg_POUpdate, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(MSG_T5F1U25_POUpdate)"
        End If
        Return False
      End If


      '写Command给WMS 超过四千字判定
      O_Send_MessageToWMS(strXML, msg_POUpdate.Header, Host_Command)
      'O_Send_ToWMSCommand_N(strXML, msg_POUpdate.Header)
      '寫Command 
      'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")



      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function Send_T2F3U1_SKUManagement_to_WMS(ByRef Result_Message As String, ByVal dicSKU As Dictionary(Of String, clsSKU),
                           ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand), ByVal Action As String) As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      'For Each objSKU_NO In dicSKU.Values
      Dim UUID = objUUID.Get_NewUUID
      Dim EventID = "T2F3U1_SKUManagement"
      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim msg_SKUManagement As New MSG_T2F3U1_SKUManagement
      msg_SKUManagement.Header = New clsHeader
      msg_SKUManagement.Header.UUID = UUID
      msg_SKUManagement.Header.EventID = EventID
      msg_SKUManagement.Header.Direction = "Primary"
      msg_SKUManagement.Header.ClientInfo = New clsHeader.clsClientInfo
      msg_SKUManagement.Header.ClientInfo.ClientID = "Handler"
      msg_SKUManagement.Header.ClientInfo.UserID = ""
      msg_SKUManagement.Header.ClientInfo.IP = ""
      msg_SKUManagement.Header.ClientInfo.MachineID = ""
      msg_SKUManagement.Body = New MSG_T2F3U1_SKUManagement.SKUDataList

      Dim SKU_NO = ""
      Dim SKUList As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList
      SKUList.SKUInfo = New List(Of MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo)
      For Each objSKU_NO In dicSKU.Values
        Dim lstPOInfo As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo
        lstPOInfo.SKU_NO = objSKU_NO.SKU_NO
        lstPOInfo.SKU_ID1 = objSKU_NO.SKU_ID1
        lstPOInfo.SKU_ID2 = objSKU_NO.SKU_ID2
        lstPOInfo.SKU_ID3 = objSKU_NO.SKU_ID3
        lstPOInfo.SKU_ALIS1 = objSKU_NO.SKU_ALIS1
        lstPOInfo.SKU_ALIS2 = objSKU_NO.SKU_ALIS2
        lstPOInfo.SKU_DESC = objSKU_NO.SKU_DESC
        lstPOInfo.SKU_CATALOG = objSKU_NO.SKU_CATALOG
        lstPOInfo.SKU_TYPE1 = objSKU_NO.SKU_TYPE1
        lstPOInfo.SKU_TYPE2 = objSKU_NO.SKU_TYPE2
        lstPOInfo.SKU_TYPE3 = objSKU_NO.SKU_TYPE3
        lstPOInfo.SKU_COMMON1 = objSKU_NO.SKU_COMMON1
        lstPOInfo.SKU_COMMON2 = objSKU_NO.SKU_COMMON2
        lstPOInfo.SKU_COMMON3 = objSKU_NO.SKU_COMMON3
        lstPOInfo.SKU_COMMON4 = objSKU_NO.SKU_COMMON4
        lstPOInfo.SKU_COMMON5 = objSKU_NO.SKU_COMMON5
        lstPOInfo.SKU_COMMON6 = objSKU_NO.SKU_COMMON6
        lstPOInfo.SKU_COMMON7 = objSKU_NO.SKU_COMMON7
        lstPOInfo.SKU_COMMON8 = objSKU_NO.SKU_COMMON8
        lstPOInfo.SKU_COMMON9 = objSKU_NO.SKU_COMMON9
        lstPOInfo.SKU_COMMON10 = objSKU_NO.SKU_COMMON10
        lstPOInfo.SKU_L = objSKU_NO.SKU_L
        lstPOInfo.SKU_W = objSKU_NO.SKU_W
        lstPOInfo.SKU_H = objSKU_NO.SKU_H
        lstPOInfo.SKU_WEIGHT = objSKU_NO.SKU_WEIGHT
        lstPOInfo.SKU_VALUE = objSKU_NO.SKU_VALUE
        lstPOInfo.SKU_UNIT = objSKU_NO.SKU_UNIT
        lstPOInfo.HIGH_WATER = objSKU_NO.HIGH_WATER
        lstPOInfo.LOW_WATER = objSKU_NO.LOW_WATER
        lstPOInfo.AVAILABLE_DAYS = objSKU_NO.AVAILABLE_DAYS
        lstPOInfo.SAVE_DAYS = objSKU_NO.SAVE_DAYS
        lstPOInfo.WEIGHT_DIFFERENCE = objSKU_NO.WEIGHT_DIFFERENCE
        lstPOInfo.ENABLE = BooleanConvertToInteger(objSKU_NO.ENABLE)
        lstPOInfo.COMMENTS = objSKU_NO.COMMENTS
        SKUList.SKUInfo.Add(lstPOInfo)
      Next
      msg_SKUManagement.Body.Action = Action
      msg_SKUManagement.Body.SKUList = SKUList '資料填寫完成

      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_MSG(Of MSG_T2F3U1_SKUManagement)(strXML, msg_SKUManagement, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(T2F3U1_SKUManagement)"
        End If
        Return False
      End If
      strXML = strXML.Replace("'", "''")

      '写Command给WMS 超过四千字判定
      O_Send_MessageToWMS(strXML, msg_SKUManagement.Header, dicHost_Command)
      'O_Send_ToWMSCommand(strXML, msg_SKUManagement.Header)
      ''寫Command 
      'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")



      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Send_T2F3U11_PackeUnitManagement_to_WMS(ByRef Result_Message As String, ByVal dicPackeUnit As Dictionary(Of String, clsMPackeUnit),
                           ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand), ByVal Action As String) As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      For Each objPackeUnit In dicPackeUnit.Values
        Dim UUID = objUUID.Get_NewUUID
        Dim EventID = "T2F3U11_PackeUnitManagement"
        '將單據宜並送給WMS 取得回復為OK後才將單據更新
        Dim msg_PackeUnitManagement As New MSG_T2F3U11_PackeUnitManagement
        msg_PackeUnitManagement.Header = New clsHeader
        msg_PackeUnitManagement.Header.UUID = UUID
        msg_PackeUnitManagement.Header.EventID = EventID
        msg_PackeUnitManagement.Header.Direction = "Primary"
        msg_PackeUnitManagement.Header.ClientInfo = New clsHeader.clsClientInfo
        msg_PackeUnitManagement.Header.ClientInfo.ClientID = "Handler"
        msg_PackeUnitManagement.Header.ClientInfo.UserID = ""
        msg_PackeUnitManagement.Header.ClientInfo.IP = ""
        msg_PackeUnitManagement.Header.ClientInfo.MachineID = ""
        msg_PackeUnitManagement.Body = New MSG_T2F3U11_PackeUnitManagement.clsBody

        Dim SKU_NO = ""
        Dim PACKEUNITList As New MSG_T2F3U11_PackeUnitManagement.clsBody.clsPackeUnitList
        PACKEUNITList.PackeUnitInfo = New List(Of MSG_T2F3U11_PackeUnitManagement.clsBody.clsPackeUnitList.clsPackeUnitInfo)
        Dim lstPACKEUNITInfo As New MSG_T2F3U11_PackeUnitManagement.clsBody.clsPackeUnitList.clsPackeUnitInfo
        lstPACKEUNITInfo.PACKE_UNIT = objPackeUnit.PACKE_UNIT
        lstPACKEUNITInfo.PACKE_UNIT_NAME = objPackeUnit.PACKE_UNIT_NAME
        lstPACKEUNITInfo.PACKE_UNIT_COMMON1 = objPackeUnit.PACKE_UNIT_COMMON1
        lstPACKEUNITInfo.PACKE_UNIT_COMMON2 = objPackeUnit.PACKE_UNIT_COMMON2
        lstPACKEUNITInfo.PACKE_UNIT_COMMON3 = objPackeUnit.PACKE_UNIT_COMMON3
        lstPACKEUNITInfo.PACKE_UNIT_COMMON4 = objPackeUnit.PACKE_UNIT_COMMON4
        lstPACKEUNITInfo.PACKE_UNIT_COMMON5 = objPackeUnit.PACKE_UNIT_COMMON5
        lstPACKEUNITInfo.COMMENTS = objPackeUnit.COMMENTS

        PACKEUNITList.PackeUnitInfo.Add(lstPACKEUNITInfo)
        msg_PackeUnitManagement.Body.Action = Action
        msg_PackeUnitManagement.Body.PackeUnitList = PACKEUNITList  '資料填寫完成


        '將物件轉成xml
        Dim strXML = ""
        If PrepareMessage_MSG(Of MSG_T2F3U11_PackeUnitManagement)(strXML, msg_PackeUnitManagement, Result_Message) = False Then
          If Result_Message = "" Then
            Result_Message = "轉XML錯誤(T2F3U11_PackeUnitManagement)"
          End If
          Return False
        End If
        strXML = strXML.Replace("'", "''")

        '写Command给WMS 超过四千字判定
        O_Send_MessageToWMS(strXML, msg_PackeUnitManagement.Header, dicHost_Command)
        ''寫Command 
        'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")

      Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Send_T2F3U12_SKUPackeStructureManagement_to_WMS(ByRef Result_Message As String, ByVal dicSKUPackeStructure As Dictionary(Of String, clsMSKUPackeStructure),
                           ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand), ByVal Action As String) As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      For Each objSKUPackeStructure In dicSKUPackeStructure.Values
        Dim UUID = objUUID.Get_NewUUID
        Dim EventID = "T2F3U12_SKUPackeStructureManagement"
        '將單據宜並送給WMS 取得回復為OK後才將單據更新
        'Dim msg_PackeUnitManagement As New MSG_T2F3U12_SKUPackeStructureManagement
        Dim msg_SKUPackeStructurManagemente As New MSG_T2F3U12_SKUPackeStructureManagement
        msg_SKUPackeStructurManagemente.Header = New clsHeader
        msg_SKUPackeStructurManagemente.Header.UUID = UUID
        msg_SKUPackeStructurManagemente.Header.EventID = EventID
        msg_SKUPackeStructurManagemente.Header.Direction = "Primary"
        msg_SKUPackeStructurManagemente.Header.ClientInfo = New clsHeader.clsClientInfo
        msg_SKUPackeStructurManagemente.Header.ClientInfo.ClientID = "Handler"
        msg_SKUPackeStructurManagemente.Header.ClientInfo.UserID = ""
        msg_SKUPackeStructurManagemente.Header.ClientInfo.IP = ""
        msg_SKUPackeStructurManagemente.Header.ClientInfo.MachineID = ""
        msg_SKUPackeStructurManagemente.Body = New MSG_T2F3U12_SKUPackeStructureManagement.clsBody

        Dim SKU_NO = ""
        Dim SKUPACKESTRUCTUREList As New MSG_T2F3U12_SKUPackeStructureManagement.clsBody.clsPackeStructureList
        SKUPACKESTRUCTUREList.PackeStructureInfo = New List(Of MSG_T2F3U12_SKUPackeStructureManagement.clsBody.clsPackeStructureList.clsPackeStructureInfo)
        Dim lstSKUPACKESTRUCTUREInfo As New MSG_T2F3U12_SKUPackeStructureManagement.clsBody.clsPackeStructureList.clsPackeStructureInfo
        lstSKUPACKESTRUCTUREInfo.SKU_NO = objSKUPackeStructure.SKU_NO
        lstSKUPACKESTRUCTUREInfo.PACKE_LV = objSKUPackeStructure.PACKE_LV
        lstSKUPACKESTRUCTUREInfo.PACKE_UNIT = objSKUPackeStructure.PACKE_UNIT
        lstSKUPACKESTRUCTUREInfo.SUB_PACKE_UNIT = objSKUPackeStructure.SUB_PACKE_UNIT
        lstSKUPACKESTRUCTUREInfo.PACKE_WEIGHT = objSKUPackeStructure.PACKE_WEIGHT
        lstSKUPACKESTRUCTUREInfo.PACKE_VOLUME = objSKUPackeStructure.PACKE_VOLUME
        lstSKUPACKESTRUCTUREInfo.PACKE_BCR = objSKUPackeStructure.PACKE_BCR
        lstSKUPACKESTRUCTUREInfo.OUT_MAX_UNIT = objSKUPackeStructure.OUT_MAX_UNIT
        lstSKUPACKESTRUCTUREInfo.IN_MAX_UNIT = objSKUPackeStructure.IN_MAX_UNIT
        lstSKUPACKESTRUCTUREInfo.QTY = objSKUPackeStructure.QTY
        lstSKUPACKESTRUCTUREInfo.COMMENTS = objSKUPackeStructure.COMMENTS
        SKUPACKESTRUCTUREList.PackeStructureInfo.Add(lstSKUPACKESTRUCTUREInfo)

        msg_SKUPackeStructurManagemente.Body.Action = Action
        msg_SKUPackeStructurManagemente.Body.PackeStructureList = SKUPACKESTRUCTUREList '資料執寫完成


        '將物件轉成xml
        Dim strXML = ""
        If PrepareMessage_MSG(Of MSG_T2F3U12_SKUPackeStructureManagement)(strXML, msg_SKUPackeStructurManagemente, Result_Message) = False Then
          If Result_Message = "" Then
            Result_Message = "轉XML錯誤(T2F3U12_SKUPackeStructureManagement)"
          End If
          Return False
        End If
        strXML = strXML.Replace("'", "''")

        '写Command给WMS 超过四千字判定
        O_Send_MessageToWMS(strXML, msg_SKUPackeStructurManagemente.Header, dicHost_Command)
        ''寫Command 
        'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")

      Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Send_T6F5U1_ItemLabelManagement_to_WMS(ByRef Result_Message As String, ByVal dicItemLabel As Dictionary(Of String, clsItemLabel),
                           ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand), ByVal Action As String, Optional ByRef ret_Wait_UUID As String = "") As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      'For Each objItemLabel In dicItemLabel.Values
      Dim UUID = objUUID.Get_NewUUID
      Dim EventID = "T6F5U1_ItemLabelManagement"
      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim msg_ItemLabelManagement As New MSG_T6F5U1_ItemLabelManagement
      msg_ItemLabelManagement.Header = New clsHeader
      ret_Wait_UUID = UUID
      msg_ItemLabelManagement.Header.UUID = UUID
      msg_ItemLabelManagement.Header.EventID = EventID
      msg_ItemLabelManagement.Header.Direction = "Primary"
      msg_ItemLabelManagement.Header.ClientInfo = New clsHeader.clsClientInfo
      msg_ItemLabelManagement.Header.ClientInfo.ClientID = "Hangler"
      msg_ItemLabelManagement.Header.ClientInfo.UserID = ""
      msg_ItemLabelManagement.Header.ClientInfo.IP = ""
      msg_ItemLabelManagement.Header.ClientInfo.MachineID = ""
      msg_ItemLabelManagement.Body = New MSG_T6F5U1_ItemLabelManagement.ItemLabelDataList
      Dim ItemLabelList As New MSG_T6F5U1_ItemLabelManagement.ItemLabelDataList.ItemLabelDataInfoList
      ItemLabelList.ItemLabelInfo = New List(Of MSG_T6F5U1_ItemLabelManagement.ItemLabelDataList.ItemLabelDataInfoList.ItemLabelDataInfo)


      For Each objItemLabel In dicItemLabel.Values

        Dim lstPOInfo As New MSG_T6F5U1_ItemLabelManagement.ItemLabelDataList.ItemLabelDataInfoList.ItemLabelDataInfo
        lstPOInfo.ITEM_LABEL_ID = objItemLabel.ITEM_LABEL_ID
        lstPOInfo.ITEM_LABEL_TYPE = objItemLabel.ITEM_LABEL_TYPE
        lstPOInfo.PO_ID = objItemLabel.PO_ID
        lstPOInfo.TAG1 = objItemLabel.TAG1
        lstPOInfo.TAG2 = objItemLabel.TAG2
        lstPOInfo.TAG3 = objItemLabel.TAG3
        lstPOInfo.TAG4 = objItemLabel.TAG4
        lstPOInfo.TAG5 = objItemLabel.TAG5
        lstPOInfo.TAG6 = objItemLabel.TAG6
        lstPOInfo.TAG7 = objItemLabel.TAG7
        lstPOInfo.TAG8 = objItemLabel.TAG8
        lstPOInfo.TAG9 = objItemLabel.TAG9
        lstPOInfo.TAG10 = objItemLabel.TAG10
        lstPOInfo.TAG11 = objItemLabel.TAG11
        lstPOInfo.TAG12 = objItemLabel.TAG12
        lstPOInfo.TAG13 = objItemLabel.TAG13
        lstPOInfo.TAG14 = objItemLabel.TAG14
        lstPOInfo.TAG15 = objItemLabel.TAG15
        lstPOInfo.TAG16 = objItemLabel.TAG16
        lstPOInfo.TAG17 = objItemLabel.TAG17
        lstPOInfo.TAG18 = objItemLabel.TAG18
        lstPOInfo.TAG19 = objItemLabel.TAG19
        lstPOInfo.TAG20 = objItemLabel.TAG20
        lstPOInfo.TAG21 = objItemLabel.TAG21
        lstPOInfo.TAG22 = objItemLabel.TAG22
        lstPOInfo.TAG23 = objItemLabel.TAG23
        lstPOInfo.TAG24 = objItemLabel.TAG24
        lstPOInfo.TAG25 = objItemLabel.TAG25
        lstPOInfo.TAG26 = objItemLabel.TAG26
        lstPOInfo.TAG27 = objItemLabel.TAG27
        lstPOInfo.TAG28 = objItemLabel.TAG28
        lstPOInfo.TAG29 = objItemLabel.TAG29
        lstPOInfo.TAG30 = objItemLabel.TAG30
        lstPOInfo.TAG31 = objItemLabel.TAG31
        lstPOInfo.PRINTED = objItemLabel.PRINTED
        lstPOInfo.CREATE_USER = objItemLabel.CREATE_USER

        ItemLabelList.ItemLabelInfo.Add(lstPOInfo)
      Next
      msg_ItemLabelManagement.Body.Action = Action
      msg_ItemLabelManagement.Body.ItemLabelList = ItemLabelList

      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_MSG(Of MSG_T6F5U1_ItemLabelManagement)(strXML, msg_ItemLabelManagement, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(T2F3U1_SKUManagement)"
        End If
        Return False
      End If

      '写Command给WMS 超过四千字判定
      O_Send_MessageToWMS(strXML, msg_ItemLabelManagement.Header, dicHost_Command)
      'O_Send_MessageToWMS(strXML, msg_SKUManagement.Header, dicHost_Command)
      ''寫Command 
      'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")

      'Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Send_T6F5U2_ItemLabelPrint_to_WMS(ByRef Result_Message As String, ByVal dicItemLabel As Dictionary(Of String, clsItemLabel),
                           ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand), ByVal Action As String) As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      For Each objItemLabel In dicItemLabel.Values
        Dim UUID = objUUID.Get_NewUUID
        Dim EventID = "T6F5U2_ItemLabelPrint"
        '將單據宜並送給WMS 取得回復為OK後才將單據更新
        Dim msg_ItemLabelManagement As New MSG_T6F5U2_ItemLabelPrint
        msg_ItemLabelManagement.Header = New clsHeader
        msg_ItemLabelManagement.Header.UUID = UUID
        msg_ItemLabelManagement.Header.EventID = EventID
        msg_ItemLabelManagement.Header.Direction = "Primary"
        msg_ItemLabelManagement.Header.ClientInfo = New clsHeader.clsClientInfo
        msg_ItemLabelManagement.Header.ClientInfo.ClientID = "Hangler"
        msg_ItemLabelManagement.Header.ClientInfo.UserID = ""
        msg_ItemLabelManagement.Header.ClientInfo.IP = ""
        msg_ItemLabelManagement.Header.ClientInfo.MachineID = ""
        msg_ItemLabelManagement.Body = New MSG_T6F5U2_ItemLabelPrint.ItemLabelDataList

        Dim ItemLabelList As New MSG_T6F5U2_ItemLabelPrint.ItemLabelDataList.ItemLabelDataInfoList
        ItemLabelList.ItemLabelInfo = New List(Of MSG_T6F5U2_ItemLabelPrint.ItemLabelDataList.ItemLabelDataInfoList.ItemLabelDataInfo)
        Dim lstPOInfo As New MSG_T6F5U2_ItemLabelPrint.ItemLabelDataList.ItemLabelDataInfoList.ItemLabelDataInfo
        lstPOInfo.ITEM_LABEL_TYPE = objItemLabel.ITEM_LABEL_TYPE
        lstPOInfo.PO_ID = objItemLabel.PO_ID
        lstPOInfo.TAG1 = objItemLabel.TAG1
        lstPOInfo.TAG2 = objItemLabel.TAG2
        lstPOInfo.TAG3 = objItemLabel.TAG3
        lstPOInfo.TAG4 = objItemLabel.TAG4
        lstPOInfo.TAG5 = objItemLabel.TAG5
        lstPOInfo.TAG6 = objItemLabel.TAG6
        lstPOInfo.TAG7 = objItemLabel.TAG7
        lstPOInfo.TAG8 = objItemLabel.TAG8
        lstPOInfo.TAG9 = objItemLabel.TAG9
        lstPOInfo.TAG10 = objItemLabel.TAG10
        lstPOInfo.TAG11 = objItemLabel.TAG11
        lstPOInfo.TAG12 = objItemLabel.TAG12
        lstPOInfo.TAG13 = objItemLabel.TAG13
        lstPOInfo.TAG14 = objItemLabel.TAG14
        lstPOInfo.TAG15 = objItemLabel.TAG15
        lstPOInfo.TAG16 = objItemLabel.TAG16
        lstPOInfo.TAG17 = objItemLabel.TAG17
        lstPOInfo.TAG18 = objItemLabel.TAG18
        lstPOInfo.TAG19 = objItemLabel.TAG19
        lstPOInfo.TAG20 = objItemLabel.TAG20
        lstPOInfo.TAG21 = objItemLabel.TAG21
        lstPOInfo.TAG22 = objItemLabel.TAG22
        lstPOInfo.TAG23 = objItemLabel.TAG23
        lstPOInfo.TAG24 = objItemLabel.TAG24
        lstPOInfo.TAG25 = objItemLabel.TAG25
        lstPOInfo.TAG26 = objItemLabel.TAG26
        lstPOInfo.TAG27 = objItemLabel.TAG27
        lstPOInfo.TAG28 = objItemLabel.TAG28
        lstPOInfo.TAG29 = objItemLabel.TAG29
        lstPOInfo.TAG30 = objItemLabel.TAG30
        lstPOInfo.TAG31 = objItemLabel.TAG31

        ItemLabelList.ItemLabelInfo.Add(lstPOInfo)

        msg_ItemLabelManagement.Body.Action = Action
        msg_ItemLabelManagement.Body.ItemLabelList = ItemLabelList

        '將物件轉成xml
        Dim strXML = ""
        If PrepareMessage_MSG(Of MSG_T6F5U2_ItemLabelPrint)(strXML, msg_ItemLabelManagement, Result_Message) = False Then
          If Result_Message = "" Then
            Result_Message = "轉XML錯誤(T2F3U2_ItemLabelPrint)"
          End If
          Return False
        End If

        '写Command给WMS 超过四千字判定
        O_Send_MessageToWMS(strXML, msg_ItemLabelManagement.Header, dicHost_Command)
        'O_Send_MessageToWMS(strXML, msg_SKUManagement.Header, dicHost_Command)
        ''寫Command 
        'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")

      Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function Send_T5F1U1_POManagement_to_WMS(ByRef Result_Message As String,
                                                          ByVal dicPO As Dictionary(Of String, clsPO),
                                                          ByRef dicAdd_PO_Line As Dictionary(Of String, clsPO_LINE),
                                                          ByRef dicAddPO_DTL As Dictionary(Of String, clsPO_DTL),
                                                          ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand),
                                                          ByVal Action As String, Optional ByRef ret_Wait_UUID As String = "") As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      For Each objPO In dicPO.Values
        Dim UUID = objUUID.Get_NewUUID
        Dim EventID = "T5F1U1_POManagement"
        '將單據宜並送給WMS 取得回復為OK後才將單據更新
        Dim msg_POManagement As New MSG_T5F1U1_PO_Management
        msg_POManagement.Header = New clsHeader
        ret_Wait_UUID = UUID                                                          'Vito_20203
        msg_POManagement.Header.UUID = UUID
        msg_POManagement.Header.EventID = EventID
        msg_POManagement.Header.Direction = "Primary"
        msg_POManagement.Header.ClientInfo = New clsHeader.clsClientInfo
        msg_POManagement.Header.ClientInfo.ClientID = "Handler"
        msg_POManagement.Header.ClientInfo.UserID = objPO.User_ID
        msg_POManagement.Header.ClientInfo.IP = ""
        msg_POManagement.Header.ClientInfo.MachineID = ""
        msg_POManagement.Body = New MSG_T5F1U1_PO_Management.clsBody

        Dim PO_ID = ""
        Dim PO As New MSG_T5F1U1_PO_Management.clsBody.clsPOInfo
        Dim POList As New MSG_T5F1U1_PO_Management.clsBody.clsPOInfo.clsPODetailList
        POList.PODetailInfo = New List(Of MSG_T5F1U1_PO_Management.clsBody.clsPOInfo.clsPODetailList.clsPODetailInfo)
        ' Dim lstPO_DTLInfo As New MSG_T5F2U61_CreatePOInByReceipt.clsBody.clsPOInfo.clsPODetailList.clsPODetailInfo
        'Dim SKU_NO = ""
        ';Dim SKUList As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList
        'SKUList.SKUInfo = New List(Of MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo)
        'Dim lstPOInfo As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo
        PO.PO_ID = objPO.PO_ID
        PO.WRITE_OFF_NO = objPO.Write_Off_No
        PO.PO_TYPE1 = objPO.PO_Type1
        PO.PO_TYPE2 = objPO.PO_Type2
        PO.PO_TYPE3 = objPO.PO_Type3
        PO.WO_TYPE = objPO.WO_Type
        PO.PRIORITY = objPO.Priority
        PO.CUSTOMER_NO = objPO.Customer_No
        PO.SUPPLIER_NO = ""
        PO.CLASS_NO = objPO.Class_No
        PO.SHIPPING_NO = objPO.Shipping_No
        PO.H_PO_ORDER_TYPE = objPO.H_PO_ORDER_TYPE
        PO.H_PO1 = objPO.H_PO1
        PO.H_PO2 = objPO.H_PO2
        PO.H_PO3 = objPO.H_PO3
        PO.H_PO4 = objPO.H_PO4
        PO.H_PO5 = objPO.H_PO5
        PO.H_PO6 = objPO.H_PO6
        PO.H_PO7 = objPO.H_PO7
        PO.H_PO8 = objPO.H_PO8
        PO.H_PO9 = objPO.H_PO9
        PO.H_PO10 = objPO.H_PO10
        PO.H_PO11 = objPO.H_PO11
        PO.H_PO12 = objPO.H_PO12
        PO.H_PO13 = objPO.H_PO13
        PO.H_PO14 = objPO.H_PO14
        PO.H_PO15 = objPO.H_PO15
        PO.H_PO16 = objPO.H_PO16
        PO.H_PO17 = objPO.H_PO17
        PO.H_PO18 = objPO.H_PO18
        PO.H_PO19 = objPO.H_PO19
        PO.H_PO20 = objPO.H_PO20
        PO.PO_KEY1 = objPO.PO_KEY1
        PO.PO_KEY2 = objPO.PO_KEY2
        PO.PO_KEY3 = objPO.PO_KEY3
        PO.PO_KEY4 = objPO.PO_KEY4
        PO.PO_KEY5 = objPO.PO_KEY5
        For Each objPO_DTL In dicAddPO_DTL.Values

          Dim lstPO_DTLInfo As New MSG_T5F1U1_PO_Management.clsBody.clsPOInfo.clsPODetailList.clsPODetailInfo

          lstPO_DTLInfo.PO_LINE_NO = objPO_DTL.PO_LINE_NO
          lstPO_DTLInfo.PO_SERIAL_NO = objPO_DTL.PO_SERIAL_NO
          lstPO_DTLInfo.SKU_NO = objPO_DTL.SKU_NO
          lstPO_DTLInfo.LOT_NO = objPO_DTL.LOT_NO
          lstPO_DTLInfo.QTY = objPO_DTL.QTY
          lstPO_DTLInfo.COMMENTS = objPO_DTL.COMMENTS
          lstPO_DTLInfo.PACKAGE_ID = objPO_DTL.PACKAGE_ID
          lstPO_DTLInfo.ITEM_COMMON1 = objPO_DTL.ITEM_COMMON1
          lstPO_DTLInfo.ITEM_COMMON2 = objPO_DTL.ITEM_COMMON2
          lstPO_DTLInfo.ITEM_COMMON3 = objPO_DTL.ITEM_COMMON3
          lstPO_DTLInfo.ITEM_COMMON4 = objPO_DTL.ITEM_COMMON4
          lstPO_DTLInfo.ITEM_COMMON5 = objPO_DTL.ITEM_COMMON5
          lstPO_DTLInfo.ITEM_COMMON6 = objPO_DTL.ITEM_COMMON6
          lstPO_DTLInfo.ITEM_COMMON7 = objPO_DTL.ITEM_COMMON7
          lstPO_DTLInfo.ITEM_COMMON8 = objPO_DTL.ITEM_COMMON8
          lstPO_DTLInfo.ITEM_COMMON9 = objPO_DTL.ITEM_COMMON9
          lstPO_DTLInfo.ITEM_COMMON10 = objPO_DTL.ITEM_COMMON10
          lstPO_DTLInfo.SORT_ITEM_COMMON1 = objPO_DTL.SORT_ITEM_COMMON1
          lstPO_DTLInfo.SORT_ITEM_COMMON2 = objPO_DTL.SORT_ITEM_COMMON2
          lstPO_DTLInfo.SORT_ITEM_COMMON3 = objPO_DTL.SORT_ITEM_COMMON3
          lstPO_DTLInfo.SORT_ITEM_COMMON4 = objPO_DTL.SORT_ITEM_COMMON4
          lstPO_DTLInfo.SORT_ITEM_COMMON5 = objPO_DTL.SORT_ITEM_COMMON5
          lstPO_DTLInfo.STORAGE_TYPE = objPO_DTL.STORAGE_TYPE
          lstPO_DTLInfo.BND = objPO_DTL.BND
          lstPO_DTLInfo.QC_STATUS = objPO_DTL.QC_STATUS
          lstPO_DTLInfo.FROM_OWNER_ID = objPO_DTL.FROM_OWNER_ID
          lstPO_DTLInfo.FROM_SUB_OWNER_ID = objPO_DTL.FROM_SUB_OWNER_ID
          lstPO_DTLInfo.TO_OWNER_ID = objPO_DTL.TO_OWNER_ID
          lstPO_DTLInfo.TO_SUB_OWNER_ID = objPO_DTL.TO_SUB_OWNER_ID
          lstPO_DTLInfo.FACTORY_ID = objPO_DTL.FACTORY_ID
          lstPO_DTLInfo.DEST_AREA_ID = objPO_DTL.DEST_AREA_ID
          lstPO_DTLInfo.DEST_LOCATION_ID = objPO_DTL.DEST_LOCATION_ID
          lstPO_DTLInfo.H_POD1 = objPO_DTL.H_POD1
          lstPO_DTLInfo.H_POD2 = objPO_DTL.H_POD2
          lstPO_DTLInfo.H_POD3 = objPO_DTL.H_POD3
          lstPO_DTLInfo.H_POD4 = objPO_DTL.H_POD4
          lstPO_DTLInfo.H_POD5 = objPO_DTL.H_POD5
          lstPO_DTLInfo.H_POD6 = objPO_DTL.H_POD6
          lstPO_DTLInfo.H_POD7 = objPO_DTL.H_POD7
          lstPO_DTLInfo.H_POD8 = objPO_DTL.H_POD8
          lstPO_DTLInfo.H_POD9 = objPO_DTL.H_POD9
          lstPO_DTLInfo.H_POD10 = objPO_DTL.H_POD10
          lstPO_DTLInfo.H_POD11 = objPO_DTL.H_POD11
          lstPO_DTLInfo.H_POD12 = objPO_DTL.H_POD12
          lstPO_DTLInfo.H_POD13 = objPO_DTL.H_POD13
          lstPO_DTLInfo.H_POD14 = objPO_DTL.H_POD14
          lstPO_DTLInfo.H_POD15 = objPO_DTL.H_POD15
          lstPO_DTLInfo.H_POD16 = objPO_DTL.H_POD16
          lstPO_DTLInfo.H_POD17 = objPO_DTL.H_POD17
          lstPO_DTLInfo.H_POD18 = objPO_DTL.H_POD18
          lstPO_DTLInfo.H_POD19 = objPO_DTL.H_POD19
          lstPO_DTLInfo.H_POD20 = objPO_DTL.H_POD20
          lstPO_DTLInfo.H_POD21 = objPO_DTL.H_POD21
          lstPO_DTLInfo.H_POD22 = objPO_DTL.H_POD22
          lstPO_DTLInfo.H_POD23 = objPO_DTL.H_POD23
          lstPO_DTLInfo.H_POD24 = objPO_DTL.H_POD24
          lstPO_DTLInfo.H_POD25 = objPO_DTL.H_POD25
          lstPO_DTLInfo.EXPIRED_DATE = ""

          POList.PODetailInfo.Add(lstPO_DTLInfo)
        Next

        msg_POManagement.Body.Action = Action
        PO.PODetailList = POList
        msg_POManagement.Body.POInfo = PO
        'msg_POManagement.Body.POInfo.PODetailList = POList
        'msg_SKUManagement.Body.Action = Action
        'msg_SKUManagement.Body.SKUList = SKUList '資料填寫完成

        '將物件轉成xml
        Dim strXML = ""
        If PrepareMessage_MSG(Of MSG_T5F1U1_PO_Management)(strXML, msg_POManagement, Result_Message) = False Then
          If Result_Message = "" Then
            Result_Message = "轉XML錯誤(MSG_T5F1U1_PO_Management)"
          End If
        End If

        'If PrepareMessage_MSG(Of MSG_T2F3U1_SKUManagement)(strXML, msg_SKUManagement, Result_Message) = False Then
        '  If Result_Message = "" Then
        '    Result_Message = "轉XML錯誤(T2F3U1_SKUManagement)"
        '  End If
        '  Return False
        'End If


        '写Command给WMS 超过四千字判定
        O_Send_MessageToWMS(strXML, msg_POManagement.Header, dicHost_Command)
        'O_Send_ToWMSCommand_N(strXML, msg_POManagement.Header)
        'O_Send_MessageToWMS(strXML, msg_SKUManagement.Header, dicHost_Command)
        ''寫Command 
        'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")

      Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Send_T5F5U1_TransactionOederManagement_to_WMS(ByRef Result_Message As String,
                                                          ByVal dicPO As Dictionary(Of String, clsPO),
                                                          ByRef dicAdd_PO_Line As Dictionary(Of String, clsPO_LINE),
                                                          ByRef dicAddPO_DTL As Dictionary(Of String, clsPO_DTL),
                                                          ByRef dicAddPO_DTL_TRANSACION As Dictionary(Of String, clsWMS_T_PO_DTL_TRANSACTION),
                                                          ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand),
                                                          ByVal Action As String, Optional ByRef ret_Wait_UUID As String = "") As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      For Each objPO In dicPO.Values
        Dim UUID = objUUID.Get_NewUUID
        Dim EventID = "T5F5U1_TransactionOederManagement"
        '將單據宜並送給WMS 取得回復為OK後才將單據更新
        Dim msg_POManagement As New MSG_T5F5U1_TransactionOederManagement
        msg_POManagement.Header = New clsHeader
        ret_Wait_UUID = UUID                                                          'Vito_20203
        msg_POManagement.Header.UUID = UUID
        msg_POManagement.Header.EventID = EventID
        msg_POManagement.Header.Direction = "Primary"
        msg_POManagement.Header.ClientInfo = New clsHeader.clsClientInfo
        msg_POManagement.Header.ClientInfo.ClientID = "Handler"
        msg_POManagement.Header.ClientInfo.UserID = ""
        msg_POManagement.Header.ClientInfo.IP = ""
        msg_POManagement.Header.ClientInfo.MachineID = ""
        msg_POManagement.Body = New MSG_T5F5U1_TransactionOederManagement.clsBody

        Dim PO_ID = ""
        Dim PO As New MSG_T5F5U1_TransactionOederManagement.clsBody.clsPOInfo
        Dim POList As New MSG_T5F5U1_TransactionOederManagement.clsBody.clsPOInfo.clsPODetailList
        POList.PODetailInfo = New List(Of MSG_T5F5U1_TransactionOederManagement.clsBody.clsPOInfo.clsPODetailList.clsPODetailInfo)
        ' Dim lstPO_DTLInfo As New MSG_T5F2U61_CreatePOInByReceipt.clsBody.clsPOInfo.clsPODetailList.clsPODetailInfo
        'Dim SKU_NO = ""
        ';Dim SKUList As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList
        'SKUList.SKUInfo = New List(Of MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo)
        'Dim lstPOInfo As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo
        PO.PO_ID = objPO.PO_ID
        PO.WRITE_OFF_NO = objPO.Write_Off_No
        PO.PO_TYPE1 = objPO.PO_Type1
        PO.PO_TYPE2 = objPO.PO_Type2
        PO.PO_TYPE3 = objPO.PO_Type3
        PO.WO_TYPE = objPO.WO_Type
        PO.PRIORITY = objPO.Priority
        PO.CUSTOMER_NO = objPO.Customer_No
        PO.SUPPLIER_NO = ""
        PO.CLASS_NO = objPO.Class_No
        PO.SHIPPING_NO = objPO.Shipping_No
        PO.H_PO_ORDER_TYPE = objPO.H_PO_ORDER_TYPE
        PO.H_PO1 = objPO.H_PO1
        PO.H_PO2 = objPO.H_PO2
        PO.H_PO3 = objPO.H_PO3
        PO.H_PO4 = objPO.H_PO4
        PO.H_PO5 = objPO.H_PO5
        PO.H_PO6 = objPO.H_PO6
        PO.H_PO7 = objPO.H_PO7
        PO.H_PO8 = objPO.H_PO8
        PO.H_PO9 = objPO.H_PO9
        PO.H_PO10 = objPO.H_PO10
        PO.H_PO11 = objPO.H_PO11
        PO.H_PO12 = objPO.H_PO12
        PO.H_PO13 = objPO.H_PO13
        PO.H_PO14 = objPO.H_PO14
        PO.H_PO15 = objPO.H_PO15
        PO.H_PO16 = objPO.H_PO16
        PO.H_PO17 = objPO.H_PO17
        PO.H_PO18 = objPO.H_PO18
        PO.H_PO19 = objPO.H_PO19
        PO.H_PO20 = objPO.H_PO20
        For Each objPO_DTL In dicAddPO_DTL.Values

          Dim lstPO_DTLInfo As New MSG_T5F5U1_TransactionOederManagement.clsBody.clsPOInfo.clsPODetailList.clsPODetailInfo

          lstPO_DTLInfo.PO_LINE_NO = objPO_DTL.PO_LINE_NO
          lstPO_DTLInfo.PO_SERIAL_NO = objPO_DTL.PO_SERIAL_NO
          lstPO_DTLInfo.SKU_NO = objPO_DTL.SKU_NO
          lstPO_DTLInfo.LOT_NO = objPO_DTL.LOT_NO
          lstPO_DTLInfo.QTY = objPO_DTL.QTY
          lstPO_DTLInfo.COMMENTS = objPO_DTL.COMMENTS
          lstPO_DTLInfo.PACKAGE_ID = objPO_DTL.PACKAGE_ID
          lstPO_DTLInfo.ITEM_COMMON1 = objPO_DTL.ITEM_COMMON1
          lstPO_DTLInfo.ITEM_COMMON2 = objPO_DTL.ITEM_COMMON2
          lstPO_DTLInfo.ITEM_COMMON3 = objPO_DTL.ITEM_COMMON3
          lstPO_DTLInfo.ITEM_COMMON4 = objPO_DTL.ITEM_COMMON4
          lstPO_DTLInfo.ITEM_COMMON5 = objPO_DTL.ITEM_COMMON5
          lstPO_DTLInfo.ITEM_COMMON6 = objPO_DTL.ITEM_COMMON6
          lstPO_DTLInfo.ITEM_COMMON7 = objPO_DTL.ITEM_COMMON7
          lstPO_DTLInfo.ITEM_COMMON8 = objPO_DTL.ITEM_COMMON8
          lstPO_DTLInfo.ITEM_COMMON9 = objPO_DTL.ITEM_COMMON9
          lstPO_DTLInfo.ITEM_COMMON10 = objPO_DTL.ITEM_COMMON10
          lstPO_DTLInfo.SORT_ITEM_COMMON1 = objPO_DTL.SORT_ITEM_COMMON1
          lstPO_DTLInfo.SORT_ITEM_COMMON2 = objPO_DTL.SORT_ITEM_COMMON2
          lstPO_DTLInfo.SORT_ITEM_COMMON3 = objPO_DTL.SORT_ITEM_COMMON3
          lstPO_DTLInfo.SORT_ITEM_COMMON4 = objPO_DTL.SORT_ITEM_COMMON4
          lstPO_DTLInfo.SORT_ITEM_COMMON5 = objPO_DTL.SORT_ITEM_COMMON5
          lstPO_DTLInfo.STORAGE_TYPE = objPO_DTL.STORAGE_TYPE
          lstPO_DTLInfo.BND = objPO_DTL.BND
          lstPO_DTLInfo.QC_STATUS = objPO_DTL.QC_STATUS
          lstPO_DTLInfo.FROM_OWNER_ID = objPO_DTL.FROM_OWNER_ID
          lstPO_DTLInfo.FROM_SUB_OWNER_ID = objPO_DTL.FROM_SUB_OWNER_ID
          lstPO_DTLInfo.TO_OWNER_ID = objPO_DTL.TO_OWNER_ID

          lstPO_DTLInfo.TO_SUB_OWNER_ID = objPO_DTL.TO_SUB_OWNER_ID
          lstPO_DTLInfo.FACTORY_ID = objPO_DTL.FACTORY_ID
          lstPO_DTLInfo.DEST_AREA_ID = objPO_DTL.DEST_AREA_ID
          lstPO_DTLInfo.DEST_LOCATION_ID = objPO_DTL.DEST_LOCATION_ID
          lstPO_DTLInfo.H_POD1 = objPO_DTL.H_POD1
          lstPO_DTLInfo.H_POD2 = objPO_DTL.H_POD2
          lstPO_DTLInfo.H_POD3 = objPO_DTL.H_POD3
          lstPO_DTLInfo.H_POD4 = objPO_DTL.H_POD4
          lstPO_DTLInfo.H_POD5 = objPO_DTL.H_POD5
          lstPO_DTLInfo.H_POD6 = objPO_DTL.H_POD6
          lstPO_DTLInfo.H_POD7 = objPO_DTL.H_POD7
          lstPO_DTLInfo.H_POD8 = objPO_DTL.H_POD8
          lstPO_DTLInfo.H_POD9 = objPO_DTL.H_POD9
          lstPO_DTLInfo.H_POD10 = objPO_DTL.H_POD10
          lstPO_DTLInfo.H_POD11 = objPO_DTL.H_POD11
          lstPO_DTLInfo.H_POD12 = objPO_DTL.H_POD12
          lstPO_DTLInfo.H_POD13 = objPO_DTL.H_POD13
          lstPO_DTLInfo.H_POD14 = objPO_DTL.H_POD14
          lstPO_DTLInfo.H_POD15 = objPO_DTL.H_POD15
          lstPO_DTLInfo.H_POD16 = objPO_DTL.H_POD16
          lstPO_DTLInfo.H_POD17 = objPO_DTL.H_POD17
          lstPO_DTLInfo.H_POD18 = objPO_DTL.H_POD18
          lstPO_DTLInfo.H_POD19 = objPO_DTL.H_POD19
          lstPO_DTLInfo.H_POD20 = objPO_DTL.H_POD20
          lstPO_DTLInfo.H_POD21 = objPO_DTL.H_POD21
          lstPO_DTLInfo.H_POD22 = objPO_DTL.H_POD22
          lstPO_DTLInfo.H_POD23 = objPO_DTL.H_POD23
          lstPO_DTLInfo.H_POD24 = objPO_DTL.H_POD24
          lstPO_DTLInfo.H_POD25 = objPO_DTL.H_POD25

          Dim PO_DTL_TRANSACTION_Key = clsWMS_T_PO_DTL_TRANSACTION.Get_Combination_Key(PO.PO_ID, objPO_DTL.PO_SERIAL_NO)
          Dim objPO_DTL_TRANSACTION As clsWMS_T_PO_DTL_TRANSACTION = Nothing
          If dicAddPO_DTL_TRANSACION.TryGetValue(PO_DTL_TRANSACTION_Key, objPO_DTL_TRANSACTION) Then
            lstPO_DTLInfo.POTransactionInfo.TRANSACTION_TYPE = objPO_DTL_TRANSACTION.TRANSACTION_TYPE
            lstPO_DTLInfo.POTransactionInfo.SKU_NO = objPO_DTL_TRANSACTION.SKU_NO
            lstPO_DTLInfo.POTransactionInfo.LOT_NO = objPO_DTL_TRANSACTION.LOT_NO
            lstPO_DTLInfo.POTransactionInfo.QTY = objPO_DTL_TRANSACTION.QTY
            lstPO_DTLInfo.POTransactionInfo.PACKAGE_ID = objPO_DTL_TRANSACTION.PACKAGE_ID
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON1 = objPO_DTL_TRANSACTION.ITEM_COMMON1
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON2 = objPO_DTL_TRANSACTION.ITEM_COMMON2
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON3 = objPO_DTL_TRANSACTION.ITEM_COMMON3
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON4 = objPO_DTL_TRANSACTION.ITEM_COMMON4
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON5 = objPO_DTL_TRANSACTION.ITEM_COMMON5
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON6 = objPO_DTL_TRANSACTION.ITEM_COMMON6
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON7 = objPO_DTL_TRANSACTION.ITEM_COMMON7
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON8 = objPO_DTL_TRANSACTION.ITEM_COMMON8
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON9 = objPO_DTL_TRANSACTION.ITEM_COMMON9
            lstPO_DTLInfo.POTransactionInfo.ITEM_COMMON10 = objPO_DTL_TRANSACTION.ITEM_COMMON10
            lstPO_DTLInfo.POTransactionInfo.SORT_ITEM_COMMON1 = objPO_DTL_TRANSACTION.SORT_ITEM_COMMON1
            lstPO_DTLInfo.POTransactionInfo.SORT_ITEM_COMMON2 = objPO_DTL_TRANSACTION.SORT_ITEM_COMMON2
            lstPO_DTLInfo.POTransactionInfo.SORT_ITEM_COMMON3 = objPO_DTL_TRANSACTION.SORT_ITEM_COMMON3
            lstPO_DTLInfo.POTransactionInfo.SORT_ITEM_COMMON4 = objPO_DTL_TRANSACTION.SORT_ITEM_COMMON4
            lstPO_DTLInfo.POTransactionInfo.SORT_ITEM_COMMON5 = objPO_DTL_TRANSACTION.SORT_ITEM_COMMON5
            lstPO_DTLInfo.POTransactionInfo.STORAGE_TYPE = objPO_DTL_TRANSACTION.STORAGE_TYPE
            lstPO_DTLInfo.POTransactionInfo.BND = objPO_DTL_TRANSACTION.BND
            lstPO_DTLInfo.POTransactionInfo.QC_STATUS = objPO_DTL_TRANSACTION.QC_STATUS
            lstPO_DTLInfo.POTransactionInfo.FROM_OWNER_ID = objPO_DTL_TRANSACTION.FROM_OWNER_ID
            lstPO_DTLInfo.POTransactionInfo.FROM_SUB_OWNER_ID = objPO_DTL_TRANSACTION.FROM_SUB_OWNER_ID
            lstPO_DTLInfo.POTransactionInfo.TO_OWNER_ID = objPO_DTL_TRANSACTION.TO_OWNER_ID
            lstPO_DTLInfo.POTransactionInfo.TO_SUB_OWNER_ID = objPO_DTL_TRANSACTION.TO_SUB_OWNER_ID
            lstPO_DTLInfo.POTransactionInfo.FACTORY_ID = objPO_DTL_TRANSACTION.FACTORY_ID
            lstPO_DTLInfo.POTransactionInfo.DEST_AREA_ID = objPO_DTL_TRANSACTION.DEST_AREA_ID
            lstPO_DTLInfo.POTransactionInfo.DEST_LOCATION_ID = objPO_DTL_TRANSACTION.DEST_LOCATION_ID
            lstPO_DTLInfo.POTransactionInfo.H_POD1 = objPO_DTL_TRANSACTION.H_POD1
            lstPO_DTLInfo.POTransactionInfo.H_POD2 = objPO_DTL_TRANSACTION.H_POD2
            lstPO_DTLInfo.POTransactionInfo.H_POD3 = objPO_DTL_TRANSACTION.H_POD3
            lstPO_DTLInfo.POTransactionInfo.H_POD4 = objPO_DTL_TRANSACTION.H_POD4
            lstPO_DTLInfo.POTransactionInfo.H_POD5 = objPO_DTL_TRANSACTION.H_POD5
            lstPO_DTLInfo.POTransactionInfo.H_POD6 = objPO_DTL_TRANSACTION.H_POD6
            lstPO_DTLInfo.POTransactionInfo.H_POD7 = objPO_DTL_TRANSACTION.H_POD7
            lstPO_DTLInfo.POTransactionInfo.H_POD8 = objPO_DTL_TRANSACTION.H_POD8
            lstPO_DTLInfo.POTransactionInfo.H_POD9 = objPO_DTL_TRANSACTION.H_POD9
            lstPO_DTLInfo.POTransactionInfo.H_POD10 = objPO_DTL_TRANSACTION.H_POD10
            lstPO_DTLInfo.POTransactionInfo.H_POD11 = objPO_DTL_TRANSACTION.H_POD11
            lstPO_DTLInfo.POTransactionInfo.H_POD12 = objPO_DTL_TRANSACTION.H_POD12
            lstPO_DTLInfo.POTransactionInfo.H_POD13 = objPO_DTL_TRANSACTION.H_POD13
            lstPO_DTLInfo.POTransactionInfo.H_POD14 = objPO_DTL_TRANSACTION.H_POD14
            lstPO_DTLInfo.POTransactionInfo.H_POD15 = objPO_DTL_TRANSACTION.H_POD15
            lstPO_DTLInfo.POTransactionInfo.H_POD16 = objPO_DTL_TRANSACTION.H_POD16
            lstPO_DTLInfo.POTransactionInfo.H_POD17 = objPO_DTL_TRANSACTION.H_POD17
            lstPO_DTLInfo.POTransactionInfo.H_POD18 = objPO_DTL_TRANSACTION.H_POD18
            lstPO_DTLInfo.POTransactionInfo.H_POD19 = objPO_DTL_TRANSACTION.H_POD19
            lstPO_DTLInfo.POTransactionInfo.H_POD20 = objPO_DTL_TRANSACTION.H_POD20
            lstPO_DTLInfo.POTransactionInfo.H_POD21 = objPO_DTL_TRANSACTION.H_POD21
            lstPO_DTLInfo.POTransactionInfo.H_POD22 = objPO_DTL_TRANSACTION.H_POD22
            lstPO_DTLInfo.POTransactionInfo.H_POD23 = objPO_DTL_TRANSACTION.H_POD23
            lstPO_DTLInfo.POTransactionInfo.H_POD24 = objPO_DTL_TRANSACTION.H_POD24
            lstPO_DTLInfo.POTransactionInfo.H_POD25 = objPO_DTL_TRANSACTION.H_POD25
          End If

          POList.PODetailInfo.Add(lstPO_DTLInfo)
        Next

        msg_POManagement.Body.Action = Action
        PO.PODetailList = POList
        msg_POManagement.Body.POInfo = PO
        'msg_POManagement.Body.POInfo.PODetailList = POList
        'msg_SKUManagement.Body.Action = Action
        'msg_SKUManagement.Body.SKUList = SKUList '資料填寫完成

        '將物件轉成xml
        Dim strXML = ""
        If PrepareMessage_MSG(Of MSG_T5F5U1_TransactionOederManagement)(strXML, msg_POManagement, Result_Message) = False Then
          If Result_Message = "" Then
            Result_Message = "轉XML錯誤(MSG_T5F1U1_PO_Management)"
          End If
        End If

        'If PrepareMessage_MSG(Of MSG_T2F3U1_SKUManagement)(strXML, msg_SKUManagement, Result_Message) = False Then
        '  If Result_Message = "" Then
        '    Result_Message = "轉XML錯誤(T2F3U1_SKUManagement)"
        '  End If
        '  Return False
        'End If


        '写Command给WMS 超过四千字判定
        O_Send_MessageToWMS(strXML, msg_POManagement.Header, dicHost_Command)
        'O_Send_MessageToWMS(strXML, msg_SKUManagement.Header, dicHost_Command)
        ''寫Command 
        'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")

      Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Send_MessageToWMS(ByVal strMessage As String, HeaderInfo As clsHeader, ByRef dicWMSCommand As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try
      'Dim dicWMSCommand As New Dictionary(Of String, clsHOST_T_WMS_Command)
      Dim DBMaxLength As Long = 2500
      If HeaderInfo.ClientInfo IsNot Nothing Then
        '如果strMessage超過DBMaxLength個字就要分成多個Message(如果只有一個Message則SEQ填0)
        If strMessage.Length < DBMaxLength Then
          Dim objWMSCommand As clsFromHostCommand = Nothing
          If HeaderInfo.ClientInfo IsNot Nothing Then
            objWMSCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.WMS, HeaderInfo.EventID, 0, HeaderInfo.ClientInfo.UserID, "", "", GetNewTime_DBFormat, strMessage, "", "", "")
          Else
            objWMSCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.WMS, HeaderInfo.EventID, 0, "", "", "", GetNewTime_DBFormat, strMessage, "", "", "")
          End If
          If dicWMSCommand.ContainsKey(objWMSCommand.gid) = False Then
            dicWMSCommand.Add(objWMSCommand.gid, objWMSCommand)
          End If
        Else
          Dim Seq As Long = 1
          Do
            Dim NewMessage As String = ""
            If strMessage.Length > DBMaxLength Then
              NewMessage = strMessage.Substring(0, 2500)
              strMessage = strMessage.Substring(2500)
            Else
              NewMessage = strMessage.Substring(0, strMessage.Length)
              strMessage = ""
            End If
            Dim objWMSCommand As clsFromHostCommand = Nothing
            If HeaderInfo.ClientInfo IsNot Nothing Then
              objWMSCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.WMS, HeaderInfo.EventID, Seq, HeaderInfo.ClientInfo.UserID, "", "", GetNewTime_DBFormat, NewMessage, "", "", "")
            Else
              objWMSCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.WMS, HeaderInfo.EventID, Seq, "", "", "", GetNewTime_DBFormat, NewMessage, "", "", "")
            End If
            If dicWMSCommand.ContainsKey(objWMSCommand.gid) = False Then
              dicWMSCommand.Add(objWMSCommand.gid, objWMSCommand)
            End If
            Seq = Seq + 1
          Loop While (strMessage.Length > 0)
        End If
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function Send_T2F7U1_OwnerManagement_to_WMS(ByRef Result_Message As String, ByVal dicOwner As Dictionary(Of String, clsOwner),
                         ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand), ByVal Action As String) As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)
      Dim UUID = objUUID.Get_NewUUID

      Dim EventID = "T2F7U1_OwnerManagement"
      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim msg_OWNERManagement As New MSG_T2F7U1_OwnerManagement
      msg_OWNERManagement.Header = New clsHeader
      msg_OWNERManagement.Header.UUID = UUID
      msg_OWNERManagement.Header.EventID = EventID
      msg_OWNERManagement.Header.Direction = "Primary"
      msg_OWNERManagement.Header.ClientInfo = New clsHeader.clsClientInfo
      msg_OWNERManagement.Header.ClientInfo.ClientID = "Handler"
      msg_OWNERManagement.Header.ClientInfo.UserID = ""
      msg_OWNERManagement.Header.ClientInfo.IP = ""
      msg_OWNERManagement.Header.ClientInfo.MachineID = ""
      msg_OWNERManagement.Body = New MSG_T2F7U1_OwnerManagement.SKUDataList


      Dim OWNERList As New MSG_T2F7U1_OwnerManagement.SKUDataList.clsOwnerList
      OWNERList.OwnerInfo = New List(Of MSG_T2F7U1_OwnerManagement.SKUDataList.clsOwnerList.clsOwnerInfo)
      For Each objOwner In dicOwner.Values
        Dim lstOWNERInfo As New MSG_T2F7U1_OwnerManagement.SKUDataList.clsOwnerList.clsOwnerInfo
        lstOWNERInfo.OWNER_NO = objOwner.OWNER_NO
        lstOWNERInfo.OWNER_ID1 = objOwner.OWNER_ID1
        lstOWNERInfo.OWNER_ID2 = objOwner.OWNER_ID2
        lstOWNERInfo.OWNER_ID3 = objOwner.OWNER_ID3
        lstOWNERInfo.OWNER_ALIS1 = objOwner.OWNER_ALIS1
        lstOWNERInfo.OWNER_ALIS2 = objOwner.OWNER_ALIS2
        lstOWNERInfo.OWNER_DESC = objOwner.OWNER_DESC
        lstOWNERInfo.OWNER_TYPE = objOwner.OWNER_TYPE
        lstOWNERInfo.OWNER_COMMON1 = objOwner.OWNER_COMMON1
        lstOWNERInfo.OWNER_COMMON2 = objOwner.OWNER_COMMON2
        lstOWNERInfo.OWNER_COMMON3 = objOwner.OWNER_COMMON3
        lstOWNERInfo.OWNER_COMMON4 = objOwner.OWNER_COMMON4
        lstOWNERInfo.OWNER_COMMON5 = objOwner.OWNER_COMMON5
        lstOWNERInfo.OWNER_COMMON6 = objOwner.OWNER_COMMON6
        lstOWNERInfo.OWNER_COMMON7 = objOwner.OWNER_COMMON7
        lstOWNERInfo.OWNER_COMMON8 = objOwner.OWNER_COMMON8
        lstOWNERInfo.OWNER_COMMON9 = objOwner.OWNER_COMMON9
        lstOWNERInfo.OWNER_COMMON10 = objOwner.OWNER_COMMON10
        lstOWNERInfo.ENABLE = BooleanConvertToInteger(objOwner.ENABLE)
        lstOWNERInfo.COMMENTS = objOwner.COMMENTS
        OWNERList.OwnerInfo.Add(lstOWNERInfo)
      Next
      msg_OWNERManagement.Body.Action = Action
      msg_OWNERManagement.Body.OwnerList = OWNERList '資料填寫完成

      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_MSG(Of MSG_T2F7U1_OwnerManagement)(strXML, msg_OWNERManagement, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(T2F7U1_OwnerManagement)"
        End If
        Return False
      End If


      '写Command给WMS 超过四千字判定
      O_Send_MessageToWMS(strXML, msg_OWNERManagement.Header, dicHost_Command)
      ''寫Command 
      'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")



      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function Send_T10F2U1_StocktakingManagement_to_WMS(ByRef Result_Message As String, ByVal dicStockTaking As Dictionary(Of String, clsTSTOCKTAKING), ByVal dicStockTaking_dtl As Dictionary(Of String, clsTSTOCKTAKINGDTL),
                          ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand), ByVal Action As String) As Boolean
    Try

      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)



      Dim EventID = "T10F2U1_StocktakingManagement"
      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim msg_StockTakingManagement As New MSG_T10F2U1_StocktakingManagement
      For Each objStockTaking In dicStockTaking.Values
        Dim UUID = objUUID.Get_NewUUID
        msg_StockTakingManagement.Header = New clsHeader
        msg_StockTakingManagement.Header.UUID = UUID
        msg_StockTakingManagement.Header.EventID = EventID
        msg_StockTakingManagement.Header.Direction = "Primary"
        msg_StockTakingManagement.Header.ClientInfo = New clsHeader.clsClientInfo
        msg_StockTakingManagement.Header.ClientInfo.ClientID = "Handler"
        msg_StockTakingManagement.Header.ClientInfo.UserID = ""
        msg_StockTakingManagement.Header.ClientInfo.IP = ""
        msg_StockTakingManagement.Header.ClientInfo.MachineID = ""
        msg_StockTakingManagement.Body = New MSG_T10F2U1_StocktakingManagement.clsBody
        msg_StockTakingManagement.Body.Action = Action
        msg_StockTakingManagement.Body.StocktakingInfo.STOCKTAKING_ID = objStockTaking.STOCKTAKING_ID
        msg_StockTakingManagement.Body.StocktakingInfo.LOCATION_GROUP_NO = objStockTaking.LOCATION_GROUP_NO
        msg_StockTakingManagement.Body.StocktakingInfo.PRIORITY = objStockTaking.PRIORITY
        msg_StockTakingManagement.Body.StocktakingInfo.STOCKTAKING_TYPE1 = objStockTaking.STOCKTAKING_TYPE1
        msg_StockTakingManagement.Body.StocktakingInfo.STOCKTAKING_TYPE2 = objStockTaking.STOCKTAKING_TYPE2
        msg_StockTakingManagement.Body.StocktakingInfo.STOCKTAKING_TYPE3 = objStockTaking.STOCKTAKING_TYPE3
        msg_StockTakingManagement.Body.StocktakingInfo.SEND_TO_HOST = objStockTaking.SEND_TO_HOST
        msg_StockTakingManagement.Body.StocktakingInfo.CHANGE_INVENTORY = objStockTaking.CHANGE_INVENTORY
        Dim StockTakingDTLList As New MSG_T10F2U1_StocktakingManagement.clsBody.clsStocktakingInfo.clsStocktakingDTLList
        StockTakingDTLList.StocktakingDTLInfo = New List(Of MSG_T10F2U1_StocktakingManagement.clsBody.clsStocktakingInfo.clsStocktakingDTLList.clsStocktakingDTLInfo)
        For Each objStockTaking_dtl In dicStockTaking_dtl.Values
          Dim StockTakingDTLInfo As New MSG_T10F2U1_StocktakingManagement.clsBody.clsStocktakingInfo.clsStocktakingDTLList.clsStocktakingDTLInfo
          If objStockTaking_dtl.STOCKTAKING_ID = objStockTaking.STOCKTAKING_ID Then
            StockTakingDTLInfo.STOCKTAKING_SERIAL_NO = objStockTaking_dtl.STOCKTAKING_SERIAL_NO
            StockTakingDTLInfo.AREA_NO = objStockTaking_dtl.AREA_NO
            StockTakingDTLInfo.BLOCK_NO = objStockTaking_dtl.BLOCK_NO
            StockTakingDTLInfo.SKU_NO = objStockTaking_dtl.SKU_NO
            StockTakingDTLInfo.BND = objStockTaking_dtl.BND
            StockTakingDTLInfo.SL_NO = objStockTaking_dtl.SL_NO
            StockTakingDTLInfo.CARRIER_ID = objStockTaking_dtl.CARRIER_ID
            StockTakingDTLInfo.PERCENTAGE = objStockTaking_dtl.PERCENTAGE
            StockTakingDTLInfo.LOT_NO = objStockTaking_dtl.LOT_NO
            StockTakingDTLInfo.ITEM_COMMON1 = objStockTaking_dtl.ITEM_COMMON1
            StockTakingDTLInfo.ITEM_COMMON2 = objStockTaking_dtl.ITEM_COMMON2
            StockTakingDTLInfo.ITEM_COMMON3 = objStockTaking_dtl.ITEM_COMMON3
            StockTakingDTLInfo.ITEM_COMMON4 = objStockTaking_dtl.ITEM_COMMON4
            StockTakingDTLInfo.ITEM_COMMON5 = objStockTaking_dtl.ITEM_COMMON5
            StockTakingDTLInfo.ITEM_COMMON6 = objStockTaking_dtl.ITEM_COMMON6
            StockTakingDTLInfo.ITEM_COMMON7 = objStockTaking_dtl.ITEM_COMMON7
            StockTakingDTLInfo.ITEM_COMMON8 = objStockTaking_dtl.ITEM_COMMON8
            StockTakingDTLInfo.ITEM_COMMON9 = objStockTaking_dtl.ITEM_COMMON9
            StockTakingDTLInfo.ITEM_COMMON10 = objStockTaking_dtl.ITEM_COMMON10
            StockTakingDTLInfo.SORT_ITEM_COMMON1 = objStockTaking_dtl.SORT_ITEM_COMMON1
            StockTakingDTLInfo.SORT_ITEM_COMMON2 = objStockTaking_dtl.SORT_ITEM_COMMON2
            StockTakingDTLInfo.SORT_ITEM_COMMON3 = objStockTaking_dtl.SORT_ITEM_COMMON3
            StockTakingDTLInfo.SORT_ITEM_COMMON4 = objStockTaking_dtl.SORT_ITEM_COMMON4
            StockTakingDTLInfo.SORT_ITEM_COMMON5 = objStockTaking_dtl.SORT_ITEM_COMMON5
            StockTakingDTLInfo.OWNER_NO = objStockTaking_dtl.OWNER_NO
            StockTakingDTLInfo.SUB_OWNER_NO = objStockTaking_dtl.SUB_OWNER_NO
            StockTakingDTLInfo.SUPPLIER_NO = objStockTaking_dtl.SUPPLIER_NO
            StockTakingDTLInfo.CUSTOMER_NO = objStockTaking_dtl.CUSTOMER_NO
            StockTakingDTLInfo.RECEIPT_DATE = objStockTaking_dtl.RECEIPT_DATE
            StockTakingDTLInfo.MANUFACETURE_DATE = objStockTaking_dtl.MANUFACETURE_DATE
            StockTakingDTLInfo.EXPIRED_DATE = objStockTaking_dtl.EXPIRED_DATE
            StockTakingDTLInfo.ERP_QTY = objStockTaking_dtl.ERP_QTY
            StockTakingDTLList.StocktakingDTLInfo.Add(StockTakingDTLInfo)
          End If
        Next
        msg_StockTakingManagement.Body.StocktakingInfo.StocktakingDTLList = StockTakingDTLList
        ''將物件轉成xml
        Dim strXML = ""
        If PrepareMessage_MSG(Of MSG_T10F2U1_StocktakingManagement)(strXML, msg_StockTakingManagement, Result_Message) = False Then
          If Result_Message = "" Then
            Result_Message = "轉XML錯誤(T2F3U3_SKULanguageManagement)"
          End If
          Return False
        End If
        strXML = strXML.Replace("'", "''")

        '写Command给WMS 超过四千字判定
        O_Send_MessageToWMS(strXML, msg_StockTakingManagement.Header, dicHost_Command)







      Next
      'Dim SKU_NO = ""
      'Dim SKUList As New MSG_T2F3U3_SKULanguageManagement.SKUDataList.SKUDataInfoList
      'SKUList.SKUInfo = New List(Of MSG_T2F3U3_SKULanguageManagement.SKUDataList.SKUDataInfoList.SKUDataInfo)


      'For Each objSKU_LANGUAGE_INFO In dicSKU_LANGUAGE_INFO.Values

      '  Dim lstPOInfo As New MSG_T2F3U3_SKULanguageManagement.SKUDataList.SKUDataInfoList.SKUDataInfo
      '  lstPOInfo.SKU_NO = objSKU_LANGUAGE_INFO.SKU_NO
      '  lstPOInfo.LANGUAGE_TYPE = objSKU_LANGUAGE_INFO.LANGUAGE_TYPE
      '  lstPOInfo.COMMON1 = objSKU_LANGUAGE_INFO.COMMON1
      '  lstPOInfo.COMMON2 = objSKU_LANGUAGE_INFO.COMMON2
      '  lstPOInfo.COMMON3 = objSKU_LANGUAGE_INFO.COMMON3
      '  lstPOInfo.COMMON4 = objSKU_LANGUAGE_INFO.COMMON4
      '  lstPOInfo.COMMON5 = objSKU_LANGUAGE_INFO.COMMON5
      '  lstPOInfo.COMMON6 = objSKU_LANGUAGE_INFO.COMMON6
      '  lstPOInfo.COMMON7 = objSKU_LANGUAGE_INFO.COMMON7
      '  lstPOInfo.COMMON8 = objSKU_LANGUAGE_INFO.COMMON8
      '  lstPOInfo.COMMON9 = objSKU_LANGUAGE_INFO.COMMON9
      '  lstPOInfo.COMMON10 = objSKU_LANGUAGE_INFO.COMMON10

      '  SKUList.SKUInfo.Add(lstPOInfo)

      '  msg_StockTakingManagement.Body.Action = Action
      '  msg_StockTakingManagement.Body.SKUList = SKUList '資料填寫完成


      'Next


      ''寫Command 
      'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")


      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Send_T5F2U62_AutoInbound_to_WMS(ByRef Result_Message As String,
                                                          ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand),
                                                          ByRef dicGluePO_DTL As Dictionary(Of String, clsPO_DTL)) As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      '先取出項次是膠塊的PO_DTL




      Dim UUID = objUUID.Get_NewUUID
      Dim EventID = "T5F2U62_AutoInbound"
      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim msg_POManagement As New MSG_T5F2U62_AutoInbound
      msg_POManagement.Header = New clsHeader
      'ret_Wait_UUID = UUID                                                          'Vito_20203
      msg_POManagement.Header.UUID = UUID
      msg_POManagement.Header.EventID = EventID
      msg_POManagement.Header.Direction = "Primary"
      msg_POManagement.Header.ClientInfo = New clsHeader.clsClientInfo
      msg_POManagement.Header.ClientInfo.ClientID = "Handler"
      msg_POManagement.Header.ClientInfo.UserID = ""
      msg_POManagement.Header.ClientInfo.IP = ""
      msg_POManagement.Header.ClientInfo.MachineID = ""
      msg_POManagement.Body = New MSG_T5F2U62_AutoInbound.clsBody



      ' Dim lstPO_DTLInfo As New MSG_T5F2U61_CreatePOInByReceipt.clsBody.clsPOInfo.clsPODetailList.clsPODetailInfo
      'Dim SKU_NO = ""
      ';Dim SKUList As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList
      'SKUList.SKUInfo = New List(Of MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo)
      'Dim lstPOInfo As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo
      Dim PODatailList As New MSG_T5F2U62_AutoInbound.clsBody.clsPODetailList
      PODatailList.PODetailInfo = New List(Of MSG_T5F2U62_AutoInbound.clsBody.clsPODetailList.clsPODetailInfo)
      For Each objPO_DTL In dicGluePO_DTL.Values

        Dim lstPO_DTLInfo As New MSG_T5F2U62_AutoInbound.clsBody.clsPODetailList.clsPODetailInfo
        lstPO_DTLInfo.PO_ID = objPO_DTL.PO_ID
        lstPO_DTLInfo.PO_LINE_NO = objPO_DTL.PO_LINE_NO
        lstPO_DTLInfo.PO_SERIAL_NO = objPO_DTL.PO_SERIAL_NO
        lstPO_DTLInfo.SKU_NO = objPO_DTL.SKU_NO
        lstPO_DTLInfo.LOT_NO = objPO_DTL.LOT_NO
        lstPO_DTLInfo.QTY = objPO_DTL.QTY
        lstPO_DTLInfo.COMMENTS = objPO_DTL.COMMENTS
        lstPO_DTLInfo.PACKAGE_ID = objPO_DTL.PACKAGE_ID
        lstPO_DTLInfo.ITEM_COMMON1 = objPO_DTL.ITEM_COMMON1
        lstPO_DTLInfo.ITEM_COMMON2 = objPO_DTL.ITEM_COMMON2
        lstPO_DTLInfo.ITEM_COMMON3 = objPO_DTL.ITEM_COMMON3
        lstPO_DTLInfo.ITEM_COMMON4 = objPO_DTL.ITEM_COMMON4
        lstPO_DTLInfo.ITEM_COMMON5 = objPO_DTL.ITEM_COMMON5
        lstPO_DTLInfo.ITEM_COMMON6 = objPO_DTL.ITEM_COMMON6
        lstPO_DTLInfo.ITEM_COMMON7 = objPO_DTL.ITEM_COMMON7
        lstPO_DTLInfo.ITEM_COMMON8 = objPO_DTL.ITEM_COMMON8
        lstPO_DTLInfo.ITEM_COMMON9 = objPO_DTL.ITEM_COMMON9
        lstPO_DTLInfo.ITEM_COMMON10 = objPO_DTL.ITEM_COMMON10
        lstPO_DTLInfo.SORT_ITEM_COMMON1 = objPO_DTL.SORT_ITEM_COMMON1
        lstPO_DTLInfo.SORT_ITEM_COMMON2 = objPO_DTL.SORT_ITEM_COMMON2
        lstPO_DTLInfo.SORT_ITEM_COMMON3 = objPO_DTL.SORT_ITEM_COMMON3
        lstPO_DTLInfo.SORT_ITEM_COMMON4 = objPO_DTL.SORT_ITEM_COMMON4
        lstPO_DTLInfo.SORT_ITEM_COMMON5 = objPO_DTL.SORT_ITEM_COMMON5
        lstPO_DTLInfo.STORAGE_TYPE = objPO_DTL.STORAGE_TYPE
        lstPO_DTLInfo.BND = objPO_DTL.BND
        lstPO_DTLInfo.QC_STATUS = objPO_DTL.QC_STATUS
        lstPO_DTLInfo.FROM_OWNER_ID = objPO_DTL.FROM_OWNER_ID
        lstPO_DTLInfo.FROM_SUB_OWNER_ID = objPO_DTL.FROM_SUB_OWNER_ID
        lstPO_DTLInfo.TO_OWNER_ID = objPO_DTL.TO_OWNER_ID
        lstPO_DTLInfo.TO_SUB_OWNER_ID = objPO_DTL.TO_SUB_OWNER_ID
        lstPO_DTLInfo.FACTORY_ID = objPO_DTL.FACTORY_ID
        lstPO_DTLInfo.DEST_AREA_ID = objPO_DTL.DEST_AREA_ID
        lstPO_DTLInfo.DEST_LOCATION_ID = objPO_DTL.DEST_LOCATION_ID
        lstPO_DTLInfo.H_POD1 = objPO_DTL.H_POD1
        lstPO_DTLInfo.H_POD2 = objPO_DTL.H_POD2
        lstPO_DTLInfo.H_POD3 = objPO_DTL.H_POD3
        lstPO_DTLInfo.H_POD4 = objPO_DTL.H_POD4
        lstPO_DTLInfo.H_POD5 = objPO_DTL.H_POD5
        lstPO_DTLInfo.H_POD6 = objPO_DTL.H_POD6
        lstPO_DTLInfo.H_POD7 = objPO_DTL.H_POD7
        lstPO_DTLInfo.H_POD8 = objPO_DTL.H_POD8
        lstPO_DTLInfo.H_POD9 = objPO_DTL.H_POD9
        lstPO_DTLInfo.H_POD10 = objPO_DTL.H_POD10
        lstPO_DTLInfo.H_POD11 = objPO_DTL.H_POD11
        lstPO_DTLInfo.H_POD12 = objPO_DTL.H_POD12
        lstPO_DTLInfo.H_POD13 = objPO_DTL.H_POD13
        lstPO_DTLInfo.H_POD14 = objPO_DTL.H_POD14
        lstPO_DTLInfo.H_POD15 = objPO_DTL.H_POD15
        lstPO_DTLInfo.H_POD16 = objPO_DTL.H_POD16
        lstPO_DTLInfo.H_POD17 = objPO_DTL.H_POD17
        lstPO_DTLInfo.H_POD18 = objPO_DTL.H_POD18
        lstPO_DTLInfo.H_POD19 = objPO_DTL.H_POD19
        lstPO_DTLInfo.H_POD20 = objPO_DTL.H_POD20
        lstPO_DTLInfo.H_POD21 = objPO_DTL.H_POD21
        lstPO_DTLInfo.H_POD22 = objPO_DTL.H_POD22
        lstPO_DTLInfo.H_POD23 = objPO_DTL.H_POD23
        lstPO_DTLInfo.H_POD24 = objPO_DTL.H_POD24
        lstPO_DTLInfo.H_POD25 = objPO_DTL.H_POD25
        lstPO_DTLInfo.EXPIRED_DATE = ""

        PODatailList.PODetailInfo.Add(lstPO_DTLInfo)
      Next


      msg_POManagement.Body.PODetailList = PODatailList
      'msg_POManagement.Body.POInfo.PODetailList = POList
      'msg_SKUManagement.Body.Action = Action
      'msg_SKUManagement.Body.SKUList = SKUList '資料填寫完成

      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_MSG(Of MSG_T5F2U62_AutoInbound)(strXML, msg_POManagement, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(MSG_T5F2U62_AutoInbound)"
        End If
      End If

      'If PrepareMessage_MSG(Of MSG_T2F3U1_SKUManagement)(strXML, msg_SKUManagement, Result_Message) = False Then
      '  If Result_Message = "" Then
      '    Result_Message = "轉XML錯誤(T2F3U1_SKUManagement)"
      '  End If
      '  Return False
      'End If


      '写Command给WMS 超过四千字判定
      O_Send_MessageToWMS(strXML, msg_POManagement.Header, dicHost_Command)
      'O_Send_MessageToWMS(strXML, msg_SKUManagement.Header, dicHost_Command)
      ''寫Command 
      'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")



      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function Send_T11F1U14_SwitchOnLocationLight_to_WMS(ByRef Result_Message As String,
                                                          ByRef dicHost_Command As Dictionary(Of String, clsFromHostCommand),
                                                          ByRef dicLocation_No As Dictionary(Of String, String)) As Boolean
    Try
      Dim dicUUID As New Dictionary(Of String, clsUUID)
      If gMain.objHandling.O_Get_UUID(enuUUID_No.HostHandler_Command.ToString, dicUUID) = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If dicUUID.Any = False Then
        Result_Message = "Get UUID False"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)

      '先取出項次是膠塊的PO_DTL




      Dim UUID = objUUID.Get_NewUUID
      Dim EventID = "T11F1U14_SwitchOnLocationLight"
      '將單據宜並送給WMS 取得回復為OK後才將單據更新
      Dim msg_SwitchOnLocationLight As New MSG_T11F1U14_SwitchOnLocationLight
      msg_SwitchOnLocationLight.Header = New clsHeader
      'ret_Wait_UUID = UUID                                                          'Vito_20203
      msg_SwitchOnLocationLight.Header.UUID = UUID
      msg_SwitchOnLocationLight.Header.EventID = EventID
      msg_SwitchOnLocationLight.Header.Direction = "Primary"
      msg_SwitchOnLocationLight.Header.ClientInfo = New clsHeader.clsClientInfo
      msg_SwitchOnLocationLight.Header.ClientInfo.ClientID = "Handler"
      msg_SwitchOnLocationLight.Header.ClientInfo.UserID = ""
      msg_SwitchOnLocationLight.Header.ClientInfo.IP = ""
      msg_SwitchOnLocationLight.Header.ClientInfo.MachineID = ""
      msg_SwitchOnLocationLight.Body = New MSG_T11F1U14_SwitchOnLocationLight.clsBody



      ' Dim lstPO_DTLInfo As New MSG_T5F2U61_CreatePOInByReceipt.clsBody.clsPOInfo.clsPODetailList.clsPODetailInfo
      'Dim SKU_NO = ""
      ';Dim SKUList As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList
      'SKUList.SKUInfo = New List(Of MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo)
      'Dim lstPOInfo As New MSG_T2F3U1_SKUManagement.SKUDataList.SKUDataInfoList.SKUDataInfo
      Dim lstLocation As New MSG_T11F1U14_SwitchOnLocationLight.clsBody.clsLocationList
      lstLocation.LocationInfo = New List(Of MSG_T11F1U14_SwitchOnLocationLight.clsBody.clsLocationList.clsLocationInfo)
      For Each LOCATION_NO In dicLocation_No.Values
        Dim LocationInfo As New MSG_T11F1U14_SwitchOnLocationLight.clsBody.clsLocationList.clsLocationInfo
        LocationInfo.LOCATION_NO = LOCATION_NO
        LocationInfo.SKU_NO = ""
        lstLocation.LocationInfo.Add(LocationInfo)
      Next
      msg_SwitchOnLocationLight.Body.LocationList = lstLocation


      '將物件轉成xml
      Dim strXML = ""
      If PrepareMessage_MSG(Of MSG_T11F1U14_SwitchOnLocationLight)(strXML, msg_SwitchOnLocationLight, Result_Message) = False Then
        If Result_Message = "" Then
          Result_Message = "轉XML錯誤(MSG_T11F1U14_SwitchOnLocationLight)"
        End If
      End If



      '写Command给WMS 超过四千字判定
      O_Send_MessageToWMS(strXML, msg_SwitchOnLocationLight.Header, dicHost_Command)
      'O_Send_MessageToWMS(strXML, msg_SKUManagement.Header, dicHost_Command)
      ''寫Command 
      'Host_Command = New clsHOST_T_WMS_Command(UUID, enuSystemType.WMS, EventID, 1, "", ModuleHelpFunc.GetNewTime_DBFormat, strXML, "", "")



      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
