'20180629
'V1.0.0
'Jerry

'结单

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T5F1S1_WOClose
  Public Function O_T5F1S1_WOClose(ByVal Receive_Msg As MSG_T5F1S1_WOClose,
                                   ByRef ret_strResultMsg As String) As Boolean
    Try
      '要更新的Table
      Dim dicAddPURXC As New Dictionary(Of String, clsPURXC) '採購單
      Dim dicUpdatePURXC As New Dictionary(Of String, clsPURXC) '採購單
      Dim dicAddMOCXB As New Dictionary(Of String, clsMOCXB) '製作單(成品入庫)
      Dim dicUpdateMOCXB As New Dictionary(Of String, clsMOCXB) '製作單(成品入庫)
      Dim dicUpdateEPSXB As New Dictionary(Of String, clsEPSXB) '出通單
      Dim dicUpdateMOCXD As New Dictionary(Of String, clsMOCXD) '領料單
      Dim dicUpdateINVXF As New Dictionary(Of String, clsINVXF) '轉調單
      '儲存要更新的SQL， 進行一次性更新
      Dim lstSql As New List(Of String)
      Dim lstSql_ERP As New List(Of String)
      '儲存要更新的SQL， 進行一次性更新
      Dim lstQueueSql As New List(Of String)
      '回報的xml
      Dim StrXML As String = ""
      '紀錄回報ERP回傳
      Dim Result As String = ""

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料處理
      If Process_Data(Receive_Msg, dicAddPURXC, dicUpdatePURXC, dicAddMOCXB, dicUpdateMOCXB, dicUpdateEPSXB, dicUpdateMOCXD, dicUpdateINVXF, ret_strResultMsg, StrXML) = False Then
        Return False
      End If

      If Get_SQL(ret_strResultMsg, dicAddPURXC, dicUpdatePURXC, dicAddMOCXB, dicUpdateMOCXB, dicUpdateEPSXB, dicUpdateMOCXD, dicUpdateINVXF, lstSql, lstSql_ERP, lstQueueSql) = False Then
        Return False
      End If
      If Execute_DataUpdate(ret_strResultMsg, lstSql, lstSql_ERP, lstQueueSql) = False Then
        Return False
      End If

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T5F1S1_WOClose,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        Dim PO_ID As String = objPOInfo.PO_ID
        '檢查PO_ID是否為空
        If PO_ID = "" Then
          ret_strResultMsg = "PO_ID is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '資料處理
  Private Function Process_Data(ByVal Receive_Msg As MSG_T5F1S1_WOClose,
                                ByRef ret_dicAddPURXC As Dictionary(Of String, clsPURXC),
                                ByRef ret_dicUpdatePURXC As Dictionary(Of String, clsPURXC),
                                ByRef ret_dicAddMOCXB As Dictionary(Of String, clsMOCXB),
                                ByRef ret_dicUpdateMOCXB As Dictionary(Of String, clsMOCXB),
                                ByRef ret_dicUpdateEPSXB As Dictionary(Of String, clsEPSXB),
                                ByRef ret_dicUpdateMOCXD As Dictionary(Of String, clsMOCXD),
                                ByRef ret_dicUpdateINVXF As Dictionary(Of String, clsINVXF),
                                ByRef ret_strResultMsg As String,
                                ByRef StrXML As String) As Boolean
    Try
      '資料處理
      If Receive_Msg.Body.POList Is Nothing Or Receive_Msg.Body.POList.POInfo.Count = 0 Then
        SendMessageToLog("WMS 给的结单资讯有缺(POList", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return True
      End If
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat_yyyymmdd()
      Dim dicEPSXB As New Dictionary(Of String, clsEPSXB)
      Dim dicMOCXD As New Dictionary(Of String, clsMOCXD)
      Dim dicINVXF As New Dictionary(Of String, clsINVXF)

      Dim UserID = Receive_Msg.Header.ClientInfo.UserID

      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        Dim dicPO As New Dictionary(Of String, clsPO)
        gMain.objHandling.O_GetDB_dicPOByPOID(objPOInfo.PO_ID, dicPO)
        If dicPO.Any = False Then
          SendMessageToLog($"PO_ID:{objPOInfo.PO_ID}有誤，無法取得PO資訊)", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        Dim objPO = dicPO.First.Value
        Dim PO_TYPE2 = objPO.PO_Type2
        Dim PO_ID As String = objPOInfo.PO_ID
        'Dim PO_TYPE1 = objPO.PO_Type1
        'Dim PO_TYPE2 = objPO.PO_Type2
        'Dim H_PO10 = objPO.H_PO10
        ''If CLng(PO_TYPE1) = enuPOType_1.Combination_in Then
        'Select Case PO_TYPE2
        '  Case enuPOType_2.Inbound_Data

        '  Case enuPOType_2.Product_In

        ''End Select
        ''ElseIf CLng(PO_TYPE1) = enuPOType_1.Picking_out Then
        Select Case PO_TYPE2
          Case enuPOType_2.Material_Out 'MOCXD
            Dim XD001 = objPO.PO_KEY1
            Dim XD002 = objPO.PO_KEY2
            gMain.objHandling.O_GetDB_dicMOCXDByXD001_XD002(XD001, XD002, dicMOCXD)
          Case enuPOType_2.Sell_Out 'EPSXB
            Dim XB001 = objPO.PO_KEY1
            Dim XB002 = objPO.PO_KEY2
            gMain.objHandling.O_GetDB_dicEPSXBByXB001_XB002(XB001, XB002, dicEPSXB)
          Case enuPOType_2.transfer_in, enuPOType_2.normal_in, enuPOType_2.transfer_out, enuPOType_2.normal_out
            Dim XF001 = objPO.PO_KEY1
            Dim XF002 = objPO.PO_KEY2
            gMain.objHandling.O_GetDB_dicINVXFByXF001_XF002(XF001, XF002, dicINVXF)
        End Select
        ''End If

        For Each objPO_DTL In objPOInfo.PO_DTLList.PO_DTLInfo
          Select Case PO_TYPE2
            Case enuPOType_2.Inbound_Data
              Dim dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
              gMain.objHandling.O_GetDB_dicPODTLByPOID_POSerialNo(PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_DTL)
              Dim objTmpPO_DTL = dicPO_DTL.First.Value

              Dim strSplit() = objTmpPO_DTL.PO_ID.Split("-")
              If strSplit.Length < 2 Then
                SendMessageToLog($"PO_ID:{objTmpPO_DTL.PO_ID}有誤，無法填出PURXC的XC001(單別)和XC002(單號)", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If

              'Dim day = Now_Time.Substring(6)
              'Dim WO_LAST = objPO_DTL.WO_ID.Substring(objPO_DTL.WO_ID.Length - 2) '取最後兩碼

              'Dim count = CInt(day & WO_LAST)

              Dim XC001 = ""
              Dim XC002 = ""
              Dim XC003 = ""
              Dim XC004 = objPO_DTL.SKU_NO
              Dim XC005 = objPO_DTL.QTY
              Dim XC006 = objTmpPO_DTL.H_POD1
              Dim XC007 = "C01"
              Dim XC008 = objPO.PO_KEY1
              Dim XC009 = objPO.PO_KEY2
              Dim XC010 = objTmpPO_DTL.PO_SERIAL_NO
              Dim XC011 = objPO_DTL.QTY
              Dim XC012 = 0
              Dim XC013 = 0
              Dim XC014 = "0"
              Dim XC015 = objPO_DTL.WO_ID
              Dim XC016 = objPO_DTL.WO_ID & LinkKey & objPO_DTL.PO_SERIAL_NO

              If XC016.Length > 20 Then
                XC016 = XC016.Substring(0, 20)
              End If
              Dim XC017 = UserID

              Dim dicInbound_DTL As New Dictionary(Of String, clsINBOUND_DTL)
              gMain.objHandling.O_GetDB_dicInboundDTLByPOID(PO_ID, dicInbound_DTL)
              Dim objInbound_DTL As clsINBOUND_DTL = Nothing
              If dicInbound_DTL.Count > 0 Then
                objInbound_DTL = dicInbound_DTL.First.Value
                XC017 = objInbound_DTL.USER_ID
              End If



              Dim tmp_dicPURXC As New Dictionary(Of String, clsPURXC)
              gMain.objHandling.O_GetDB_dicPURXCByKey(XC008, XC009, XC010, XC016, tmp_dicPURXC)

              'If tmp_dicPURXC IsNot Nothing AndAlso tmp_dicPURXC.Count > 0 Then
              '  Dim tmp_objPURXC = tmp_dicPURXC.Values.First.Clone

              '  Dim sum As Integer = CInt(tmp_objPURXC.XC005) + CInt(XC005)

              '  tmp_objPURXC.XC005 = sum.ToString

              '  If ret_dicUpdatePURXC.ContainsKey(tmp_objPURXC.gid) = False Then
              '    ret_dicUpdatePURXC.Add(tmp_objPURXC.gid, tmp_objPURXC)
              '  End If
              'Else
              If tmp_dicPURXC.Any Then
                Dim objUpdatePURXC = tmp_dicPURXC.First.Value.Clone
                objUpdatePURXC.XC001 = XC001
                objUpdatePURXC.XC001 = XC001
                objUpdatePURXC.XC002 = XC002
                objUpdatePURXC.XC003 = XC003
                objUpdatePURXC.XC004 = XC004
                objUpdatePURXC.XC005 = XC005
                objUpdatePURXC.XC006 = XC006
                objUpdatePURXC.XC007 = XC007
                'objUpdatePURXC.XC008 = XC008
                'objUpdatePURXC.XC009 = XC009
                'objUpdatePURXC.XC010 = XC010
                objUpdatePURXC.XC011 = objUpdatePURXC.XC011 + XC011
                objUpdatePURXC.XC012 = XC012
                objUpdatePURXC.XC013 = XC013
                objUpdatePURXC.XC014 = XC014
                objUpdatePURXC.XC015 = XC015
                'objUpdatePURXC.XC016 = XC016
                objUpdatePURXC.XC017 = XC017

                If ret_dicUpdatePURXC.ContainsKey(objUpdatePURXC.gid) = False Then
                  ret_dicUpdatePURXC.Add(objUpdatePURXC.gid, objUpdatePURXC)
                End If
              Else
                Dim objNewPURXC = New clsPURXC(XC001, XC002, XC003, XC004, XC005, XC006, XC007, XC008, XC009, XC010, XC011, XC012, XC013, XC014, XC015, XC016, XC017)
                If ret_dicAddPURXC.ContainsKey(objNewPURXC.gid) = False Then
                  ret_dicAddPURXC.Add(objNewPURXC.gid, objNewPURXC)
                End If
              End If

              'Next
              'End If
            Case enuPOType_2.Product_In
              Dim dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
              gMain.objHandling.O_GetDB_dicPODTLByPOID_POSerialNo(PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_DTL)
              Dim objTmpPO_DTL = dicPO_DTL.First.Value

              Dim strSplit() = objTmpPO_DTL.PO_ID.Split("-")
              If strSplit.Length < 2 Then
                SendMessageToLog($"PO_ID:{objTmpPO_DTL.PO_ID}有誤，無法填出MOCXB的XB001(單別)和XB002(單號)", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
              End If

              'Dim day = Now_Time.Substring(6)
              'Dim WO_LAST = objPO_DTL.WO_ID.Substring(objPO_DTL.WO_ID.Length - 2) '取最後兩碼

              'Dim count = CInt(day & WO_LAST)

              'For Each objWO_DTL In lst_WO_DTL
              Dim XB001 = ""
              Dim XB002 = ""
              Dim XB003 = Now_Time
              Dim XB004 = "1"
              Dim XB005 = objPO_DTL.SKU_NO
              Dim XB006 = ""
              Dim XB007 = "C01"
              Dim XB008 = objPO_DTL.QTY
              Dim XB009 = objTmpPO_DTL.H_POD1
              Dim XB010 = objPO.PO_KEY1
              Dim XB011 = objPO.PO_KEY2
              Dim XB012 = "0"
              Dim XB013 = objPO_DTL.WO_ID & LinkKey & objPO_DTL.PO_SERIAL_NO

              If XB013.Length > 20 Then
                XB013 = XB013.Substring(0, 20)
              End If

              Dim XB014 = objPO_DTL.WO_ID

              Dim XB015 = UserID

              Dim dicInbound_DTL As New Dictionary(Of String, clsINBOUND_DTL)
              gMain.objHandling.O_GetDB_dicInboundDTLByPOID(PO_ID, dicInbound_DTL)
              Dim objInbound_DTL As clsINBOUND_DTL = Nothing
              If dicInbound_DTL.Count > 0 Then
                objInbound_DTL = dicInbound_DTL.First.Value
                XB015 = objInbound_DTL.USER_ID
              End If
              '假如一張訂單組多次，會導致工單INSERT寫入結果失敗，因此要先查看看原本的單據是否已有紀錄，若有，要用修改的，QTY用累加的
              Dim tmp_dicMOCXB As New Dictionary(Of String, clsMOCXB)
              gMain.objHandling.O_GetDB_dicMOCXBByKey(XB010, XB011, XB013, tmp_dicMOCXB)

              'If tmp_dicMOCXB IsNot Nothing AndAlso tmp_dicMOCXB.Count > 0 Then
              '  Dim tmp_objMOCXB = tmp_dicMOCXB.First.Value.Clone

              '  Dim sum As Integer = CInt(tmp_objMOCXB.XB008) + CInt(XB008)

              '  tmp_objMOCXB.XB008 = sum.ToString

              '  If ret_dicUpdateMOCXB.ContainsKey(tmp_objMOCXB.gid) = False Then
              '    ret_dicUpdateMOCXB.Add(tmp_objMOCXB.gid, tmp_objMOCXB)
              '  End If
              'Else
              If tmp_dicMOCXB.Any Then
                Dim objUpdateMOCXB = tmp_dicMOCXB.First.Value.Clone

                objUpdateMOCXB.XB001 = XB001
                objUpdateMOCXB.XB002 = XB002
                objUpdateMOCXB.XB003 = XB003
                objUpdateMOCXB.XB004 = XB004
                objUpdateMOCXB.XB005 = XB005
                objUpdateMOCXB.XB006 = XB006
                objUpdateMOCXB.XB007 = XB007
                objUpdateMOCXB.XB008 = objUpdateMOCXB.XB008 + XB008
                objUpdateMOCXB.XB009 = XB009
                'objUpdateMOCXB.XB010 = XB010
                'objUpdateMOCXB.XB011 = XB011
                objUpdateMOCXB.XB012 = XB012
                'objUpdateMOCXB.XB013 = XB013
                objUpdateMOCXB.XB014 = XB014
                objUpdateMOCXB.XB015 = XB015

                If ret_dicUpdateMOCXB.ContainsKey(objUpdateMOCXB.gid) = False Then
                  ret_dicUpdateMOCXB.Add(objUpdateMOCXB.gid, objUpdateMOCXB)
                End If
              Else
                Dim objNewMOCXB = New clsMOCXB(XB001, XB002, XB003, XB004, XB005, XB006, XB007, XB008, XB009, XB010, XB011, XB012, XB013, XB014, XB015)
                If ret_dicAddMOCXB.ContainsKey(objNewMOCXB.gid) = False Then
                  ret_dicAddMOCXB.Add(objNewMOCXB.gid, objNewMOCXB)
                End If
              End If
              'Next
              'End If
            Case enuPOType_2.Material_Out
              Dim objUpdateMOCXD As clsMOCXD = Nothing
              For Each objMOCXD In dicMOCXD.Values
                If objMOCXD.XD001 = objPO.PO_KEY1 AndAlso objMOCXD.XD002 = objPO.PO_KEY2 AndAlso objMOCXD.XD013 = objPO_DTL.PO_SERIAL_NO Then
                  objUpdateMOCXD = objMOCXD.Clone

                  objUpdateMOCXD.XD014 = objUpdateMOCXD.XD014 + objPO_DTL.QTY
                  If ret_dicUpdateMOCXD.ContainsKey(objUpdateMOCXD.gid) = False Then
                    ret_dicUpdateMOCXD.Add(objUpdateMOCXD.gid, objUpdateMOCXD)
                  End If
                End If
              Next
            Case enuPOType_2.Sell_Out
              Dim objUpdateEPSXB As clsEPSXB = Nothing
              For Each objEPSXB In dicEPSXB.Values
                If objEPSXB.XB001 = objPO.PO_KEY1 AndAlso objEPSXB.XB002 = objPO.PO_KEY2 AndAlso objEPSXB.XB003 = objPO_DTL.PO_SERIAL_NO Then
                  objUpdateEPSXB = objEPSXB.Clone

                  objUpdateEPSXB.XB006 = objUpdateEPSXB.XB006 + objPO_DTL.QTY
                  If ret_dicUpdateEPSXB.ContainsKey(objUpdateEPSXB.gid) = False Then
                    ret_dicUpdateEPSXB.Add(objUpdateEPSXB.gid, objUpdateEPSXB)
                  End If
                End If
              Next
            Case enuPOType_2.transfer_in, enuPOType_2.normal_in, enuPOType_2.transfer_out, enuPOType_2.normal_out
              Dim objUpdateINVXF As clsINVXF = Nothing
              For Each objINVXF In dicINVXF.Values
                If objINVXF.XF001 = objPO.PO_KEY1 AndAlso objINVXF.XF002 = objPO.PO_KEY2 AndAlso objINVXF.XF004 = objPO_DTL.PO_SERIAL_NO Then
                  objUpdateINVXF = objINVXF.Clone

                  objUpdateINVXF.XF013 = objUpdateINVXF.XF013 + objPO_DTL.QTY
                  If ret_dicUpdateINVXF.ContainsKey(objUpdateINVXF.gid) = False Then
                    ret_dicUpdateINVXF.Add(objUpdateINVXF.gid, objUpdateINVXF)
                  End If
                End If
              Next
          End Select
        Next
      Next

#Region "配合頂新ERP，用CONFIG的FLAG控制是否運行此處"
      If f_CommonERPSwitch = True Then
        For Each objPOInfo In Receive_Msg.Body.POList.POInfo
          Dim dicPO As New Dictionary(Of String, clsPO)
          Dim dicPO_DTL As New Dictionary(Of String, clsPO_DTL)
          Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)

          For Each objPO In dicPO.Values
            Dim PO_ID As String = objPO.PO_ID
            Dim PO_TYPE1 = objPO.PO_Type1
            Dim PO_TYPE2 = objPO.PO_Type2
            Dim H_PO10 = objPO.H_PO10
            Dim dicItem_Label As New Dictionary(Of String, clsItemLabel)

            '回報單據

            If CLng(PO_TYPE1) = enuPOType_1.Combination_in Then
              Select Case PO_TYPE2
#Region "採購單"
                Case "101"
                  If SendBuyData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "採購入庫單"
                Case "102", "181" '採購入庫單
                  If PO_TYPE2 = "181" And H_PO10 <> "0" Then
                    Exit Select
                  End If
                  If InboundData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "生產入庫單"
                Case "103" '生產入庫單
                  If ProduceInData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "雜收單"
                Case "104" '雜收單
                  If OtherInData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "調撥入庫單"
                Case "144" '調撥入庫單
                  Dim dicInbound_DTL As New Dictionary(Of String, clsINBOUND_DTL)
                  Dim dicOutbound_DTL As New Dictionary(Of String, clsOUTBOUND_DTL)
                  If gMain.objHandling.O_GetDB_dicInboundDTLByPOID(PO_ID, dicInbound_DTL) = False Then
                    ret_strResultMsg = String.Format("Get WMS_T_INBOUND_DTL Failed PO_ID={0}", PO_ID)
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If

                  If TransferData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg, dicInbound_DTL, dicOutbound_DTL) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "銷退單"
                Case enuPOType_2.Sell_Back
                  If SellReturnData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "手工入庫單"

#End Region
              End Select
            ElseIf CLng(PO_TYPE1) = enuPOType_1.Picking_out Then
              Select Case PO_TYPE2
#Region "雜發單"
                Case "303" '雜發單
                  If OtherOutData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "銷貨單"
                Case "304", "381" '銷貨單
                  If PO_TYPE2 = "381" And H_PO10 <> "0" Then
                    Exit Select
                  End If
                  If SellData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "調撥出庫"
                Case "344" '調撥出庫
                  Dim dicOutbound_DTL As New Dictionary(Of String, clsOUTBOUND_DTL)
                  Dim dicInbound_DTL As New Dictionary(Of String, clsINBOUND_DTL)
                  If gMain.objHandling.O_GetDB_dicOutboundDTLByPOID(objPO.PO_ID, dicOutbound_DTL) = False Then
                    ret_strResultMsg = String.Format("Get WMS_T_OUTBOUND_DTL Failed PO_ID={0}", objPO.PO_ID)
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If

                  If TransferData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg, dicInbound_DTL, dicOutbound_DTL) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
#Region "退供應商"
                  'Case enuPOType_2.InboundReturn_Data
                  '  If InboundReturnData_Report(objPO, dicPO_DTL, StrXML, ret_strResultMsg) = False Then
                  '    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  '    Return False
                  '  End If
#End Region
#Region "手工出庫單"

#End Region
              End Select

            ElseIf CLng(PO_TYPE1) = enuPOType_1.Transaction Then
              Select Case PO_TYPE2
#Region "貨主調撥"
                Case "631" '貨主調撥
                  If TransferOwnerData_Report(objPO, dicPO_DTL, dicPO_Merge, StrXML, ret_strResultMsg) = False Then
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
#End Region
              End Select
            End If 'If CLng(PO_TYPE1) = enuPOType_1.Combination_in Then


          Next  'For Each objPO In dicPO.Values

        Next  'For Each objPOInfo In Receive_Msg.Body.POList.POInfo

      End If 'If gMain.f_Common_ERP_Switch = True Then
#End Region
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef ret_dicAddPURXC As Dictionary(Of String, clsPURXC),
                           ByRef ret_dicUpdatePURXC As Dictionary(Of String, clsPURXC),
                           ByRef ret_dicAddMOCXB As Dictionary(Of String, clsMOCXB),
                           ByRef ret_dicUpdateMOCXB As Dictionary(Of String, clsMOCXB),
                           ByRef ret_dicUpdateEPSXB As Dictionary(Of String, clsEPSXB),
                           ByRef ret_dicUpdateMOCXD As Dictionary(Of String, clsMOCXD),
                           ByRef ret_dicUpdateINVXF As Dictionary(Of String, clsINVXF),
                           ByRef lstSql As List(Of String), ByRef lstSql_ERP As List(Of String), ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      For Each obj In ret_dicAddPURXC.Values
        If obj.O_Add_Insert_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Insert PURXC SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      For Each obj In ret_dicUpdatePURXC.Values
        If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Update PURXC SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      For Each obj In ret_dicAddMOCXB.Values
        If obj.O_Add_Insert_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Insert MOCXB SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      For Each obj In ret_dicUpdateMOCXB.Values
        If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Update MOCXB SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      For Each obj In ret_dicUpdateEPSXB.Values
        If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Update EPSXB SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      For Each obj In ret_dicUpdateMOCXD.Values
        If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Update MOCXD SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      For Each obj In ret_dicUpdateINVXF.Values
        If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Update INVXF SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行新增的Carrier和Carrier_Status的SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef Result_Message As String,
                                         ByRef lstSql As List(Of String), ByRef lstSql_ERP As List(Of String), ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSql.Any = True Then
        If Common_DBManagement.BatchUpdate(lstSql) = False Then
          '更新DB失敗則回傳False
          Result_Message = "eHOST 更新资料库失败"
          Return False
        End If
      End If
      If lstSql_ERP.Any = True Then
        If ERP_DBManagement.BatchUpdate(lstSql_ERP) = False Then
          '更新DB失敗則回傳False
          Result_Message = "ERP 更新资料库失败"
          Return False
        End If
      End If
      Common_DBManagement.AddQueued(lstQueueSql)

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 採購單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function SendBuyData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      Dim Now_time = GetNewTime_DBFormat()
      '表頭
      Dim InDataReport As New MSG_SendInData
      Dim InDataList As New MSG_SendInData.clsInDataList
      InDataList.InDataInfo = New List(Of MSG_SendInData.clsInDataList.clsInDataInfo)
      Dim InDataInfo As New MSG_SendInData.clsInDataList.clsInDataInfo
      InDataInfo.POType = ObjPO.PO_KEY1
      InDataInfo.POID = ObjPO.PO_KEY2
      InDataInfo.InDateTime = Now_time
      InDataInfo.FactoryID = ObjPO.H_PO3

      '加入表身
      Dim InDetailDataList As New MSG_SendInData.clsInDataList.clsInDataInfo.clsInDetailDataList
      InDetailDataList.InDetailDataInfo = New List(Of MSG_SendInData.clsInDataList.clsInDataInfo.clsInDetailDataList.clsInDetailDataInfo)
      For Each objPO_DTL In dicPO_DTL.Values
        Dim InDetailDataInfo As New MSG_SendInData.clsInDataList.clsInDataInfo.clsInDetailDataList.clsInDetailDataInfo
        InDetailDataInfo.SerialID = objPO_DTL.PO_SERIAL_NO
        InDetailDataInfo.SKU = objPO_DTL.SKU_NO
        InDetailDataInfo.LotId = objPO_DTL.LOT_NO
        InDetailDataInfo.CheckQty = objPO_DTL.QTY_FINISH

        'BuyDetailDataInfo.Item_Common1 = objPO_DTL.ITEM_COMMON1
        'BuyDetailDataInfo.Item_Common2 = objPO_DTL.ITEM_COMMON2
        InDetailDataList.InDetailDataInfo.Add(InDetailDataInfo)
      Next
      InDataInfo.InDetailDataList = InDetailDataList
      InDataList.InDataInfo.Add(InDataInfo)
      InDataReport.InDataList = InDataList

      '將物件轉成xml     
      If PrepareMessage_MSG(Of MSG_SendInData)(StrXML, InDataReport, result_Message) = False Then
        If result_Message = "" Then
          result_Message = "轉XML錯誤(SendBuyData_採購單回報)"
        End If
        Return False
      End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 採購入庫單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function InboundData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      Dim Now_time = GetNewTime_DBFormat()
      Dim FactoryId = ""
      '表頭
      Dim InboundDataReport As New MSG_SendInboundData
      InboundDataReport.WebService_ID = "WMS"
      InboundDataReport.EventID = "InboundData"
      Dim InboundDataList As New MSG_SendInboundData.clsInboundDataList
      InboundDataList.InboundDataInfo = New List(Of MSG_SendInboundData.clsInboundDataList.clsInboundDataInfo)
      Dim InboundDataInfo As New MSG_SendInboundData.clsInboundDataList.clsInboundDataInfo
      InboundDataInfo.POType = ObjPO.PO_KEY1
      InboundDataInfo.POId = ObjPO.PO_KEY2
      InboundDataInfo.InboundDateTime = Now_time
      InboundDataInfo.FactoryId = ObjPO.H_PO3

      Dim warehouse = ""
      '加入表身
      Dim InboundDetailDataList As New MSG_SendInboundData.clsInboundDataList.clsInboundDataInfo.clsInboundDetailDataList
      InboundDetailDataList.InboundDetailDataInfo = New List(Of MSG_SendInboundData.clsInboundDataList.clsInboundDataInfo.clsInboundDetailDataList.clsInboundDetailDataInfo)
      For Each objPO_DTL In dicPO_DTL.Values
        Dim InboundDetailDataInfo As New MSG_SendInboundData.clsInboundDataList.clsInboundDataInfo.clsInboundDetailDataList.clsInboundDetailDataInfo
        InboundDetailDataInfo.SerialId = objPO_DTL.PO_SERIAL_NO
        Dim SKU_NO As String() = objPO_DTL.SKU_NO.Split("_")
        InboundDetailDataInfo.SKU = SKU_NO(0)
        InboundDetailDataInfo.LotId = objPO_DTL.LOT_NO
        Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)
        gMain.objHandling.O_GetDB_dicPO_MergeByPO_ID_PO_Serial_No(objPO_DTL.PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_Merge)
        If dicPO_Merge.Any Then
          InboundDetailDataInfo.CheckQty = dicPO_Merge.First.Value.QTY_FINISH
        Else
          Return False
        End If
        InboundDetailDataInfo.Item_Common1 = objPO_DTL.ITEM_COMMON1
        InboundDetailDataInfo.Item_Common2 = objPO_DTL.ITEM_COMMON2
        InboundDetailDataList.InboundDetailDataInfo.Add(InboundDetailDataInfo)
        FactoryId = objPO_DTL.TO_OWNER_ID
        warehouse = objPO_DTL.ITEM_COMMON10
      Next
      InboundDataInfo.FactoryId = FactoryId
      'InboundDataInfo.Warehouse = warehouse
      InboundDataInfo.InboundDetailDataList = InboundDetailDataList
      InboundDataList.InboundDataInfo.Add(InboundDataInfo)
      InboundDataReport.InboundDataList = InboundDataList

      '將物件轉成xml     
      If PrepareMessage_MSG(Of MSG_SendInboundData)(StrXML, InboundDataReport, result_Message) = False Then
        If result_Message = "" Then
          result_Message = "轉XML錯誤(SendInboundData_採購入庫單回報)"
        End If
        Return False
      End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 調撥入庫單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function TransferData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String, ByVal dicInbound_DTL As Dictionary(Of String, clsINBOUND_DTL), ByRef dicOutobund_DTL As Dictionary(Of String, clsOUTBOUND_DTL)) As Boolean
    Try
      Dim Now_time = GetNewTime_DBFormat()
      '表頭
      Dim TransferDataReport As New MSG_SendTransferData
      TransferDataReport.WebService_ID = "WMS"
      TransferDataReport.EventID = "TransferData"
      Dim TransferDataList As New MSG_SendTransferData.clsTransferDataList
      TransferDataList.TransferDataInfo = New List(Of MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo)
      Dim TransferDataInfo As New MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo
      TransferDataInfo.POType = ObjPO.PO_KEY1
      TransferDataInfo.POId = ObjPO.PO_KEY2
      TransferDataInfo.TransferDateTime = Now_time
      TransferDataInfo.FactoryId = ObjPO.H_PO3
      TransferDataInfo.TransferOutWarehouse = ObjPO.H_PO7
      TransferDataInfo.TransferInWarehouse = ObjPO.H_PO8

      Dim FactoryId = ""
      Dim warehouse = ""
      '加入表身
      Dim TransferDetailDataList As New MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo.clsTransferDetailDataList
      TransferDetailDataList.TransferDetailDataInfo = New List(Of MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo.clsTransferDetailDataList.clsTransferDetailDataInfo)
      If dicOutobund_DTL.Any Then
        Dim lstPackage_ID As New List(Of String)
        For Each objOutbound_DTL In dicOutobund_DTL.Values
          If lstPackage_ID.Contains(objOutbound_DTL.PACKAGE_ID) = True Then Continue For
          lstPackage_ID.Add(objOutbound_DTL.PACKAGE_ID)
          Dim TransferDetailDataInfo As New MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo.clsTransferDetailDataList.clsTransferDetailDataInfo
          TransferDetailDataInfo.SerialId = objOutbound_DTL.PO_SERIAL_NO
          Dim SKU_NO As String() = objOutbound_DTL.SKU_NO.Split("_")
          TransferDetailDataInfo.SKU = SKU_NO(0)
          TransferDetailDataInfo.LotId = objOutbound_DTL.LOT_NO
          Dim CheckQty As Decimal = objOutbound_DTL.QTY_OUTBOUND
          For Each objTmpOutbound_DTL In dicOutobund_DTL.Values
            If objTmpOutbound_DTL.KEY_NO <> objOutbound_DTL.KEY_NO AndAlso objTmpOutbound_DTL.PACKAGE_ID = objOutbound_DTL.PACKAGE_ID Then
              CheckQty += objTmpOutbound_DTL.QTY_OUTBOUND
            End If
          Next
          TransferDetailDataInfo.CheckQty = CheckQty
          'TransferDetailDataInfo.SN = objOutbound_DTL.PACKAGE_ID
          TransferDetailDataInfo.Item_Common3 = ""
          For Each objPO_DTL In dicPO_DTL.Values
            If objPO_DTL.PO_SERIAL_NO = objOutbound_DTL.PO_SERIAL_NO Then
              TransferDetailDataInfo.Item_Common3 = objPO_DTL.SORT_ITEM_COMMON5
              warehouse = objPO_DTL.ITEM_COMMON10
              FactoryId = objPO_DTL.FROM_OWNER_ID
              Exit For
            End If
          Next
          'TransferDetailDataInfo.TransferOutWarehouse = objPO_DTL.H_POD8
          'TransferDetailDataInfo.TransferInWarehouse = objPO_DTL.H_POD9
          TransferDetailDataList.TransferDetailDataInfo.Add(TransferDetailDataInfo)
        Next
      Else
        For Each objInbound_DTL In dicInbound_DTL.Values
          Dim TransferDetailDataInfo As New MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo.clsTransferDetailDataList.clsTransferDetailDataInfo
          TransferDetailDataInfo.SerialId = objInbound_DTL.PO_SERIAL_NO
          Dim SKU_NO As String() = objInbound_DTL.SKU_NO.Split("_")
          TransferDetailDataInfo.SKU = SKU_NO(0)
          TransferDetailDataInfo.LotId = objInbound_DTL.LOT_NO
          TransferDetailDataInfo.CheckQty = objInbound_DTL.QTY_INBOUND
          'TransferDetailDataInfo.SN = objInbound_DTL.PACKAGE_ID
          TransferDetailDataInfo.Item_Common3 = ""
          For Each objPO_DTL In dicPO_DTL.Values
            If objPO_DTL.PO_SERIAL_NO = objInbound_DTL.PO_SERIAL_NO Then
              TransferDetailDataInfo.Item_Common3 = objPO_DTL.SORT_ITEM_COMMON5
              warehouse = objPO_DTL.ITEM_COMMON10
              FactoryId = objPO_DTL.TO_OWNER_ID
              Exit For
            End If
          Next
          'TransferDetailDataInfo.TransferOutWarehouse = objPO_DTL.H_POD8
          'TransferDetailDataInfo.TransferInWarehouse = objPO_DTL.H_POD9
          TransferDetailDataList.TransferDetailDataInfo.Add(TransferDetailDataInfo)
        Next

      End If

      TransferDataInfo.FactoryId = FactoryId
      TransferDataInfo.TransferDetailDataList = TransferDetailDataList
      TransferDataList.TransferDataInfo.Add(TransferDataInfo)
      TransferDataReport.TransferDataList = TransferDataList

      '將物件轉成xml     
      If PrepareMessage_MSG(Of MSG_SendTransferData)(StrXML, TransferDataReport, result_Message) = False Then
        If result_Message = "" Then
          result_Message = "轉XML錯誤(SendTransferData_雕撥入庫單回報)"
        End If
        Return False
      End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 生產入庫單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function ProduceInData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      Dim Now_time = GetNewTime_DBFormat()
      '表頭
      Dim ProduceInDataReport As New MSG_SendProduceInData
      ProduceInDataReport.WebService_ID = "WMS"
      ProduceInDataReport.EventID = "ProduceInData"
      Dim ProduceInDataList As New MSG_SendProduceInData.clsProduceInDataList
      ProduceInDataList.ProduceInDataInfo = New List(Of MSG_SendProduceInData.clsProduceInDataList.clsProduceInDataInfo)
      Dim ProduceInDataInfo As New MSG_SendProduceInData.clsProduceInDataList.clsProduceInDataInfo
      ProduceInDataInfo.POType = ObjPO.PO_KEY1
      ProduceInDataInfo.POId = ObjPO.PO_KEY2
      ProduceInDataInfo.ProduceInDateTime = Now_time
      ProduceInDataInfo.FactoryId = ObjPO.H_PO3
      Dim FactoryId = ""
      Dim warehouse = ""
      '加入表身
      Dim ProduceInDetailDataList As New MSG_SendProduceInData.clsProduceInDataList.clsProduceInDataInfo.clsProduceInDetailDataList
      ProduceInDetailDataList.ProduceInDetailDataInfo = New List(Of MSG_SendProduceInData.clsProduceInDataList.clsProduceInDataInfo.clsProduceInDetailDataList.clsProduceInDetailDataInfo)
      For Each objPO_DTL In dicPO_DTL.Values
        Dim ProduceInDetailDataInfo As New MSG_SendProduceInData.clsProduceInDataList.clsProduceInDataInfo.clsProduceInDetailDataList.clsProduceInDetailDataInfo
        ProduceInDetailDataInfo.SerialId = objPO_DTL.PO_SERIAL_NO
        Dim SKU_NO As String() = objPO_DTL.SKU_NO.Split("_")
        ProduceInDetailDataInfo.SKU = SKU_NO(0)
        ProduceInDetailDataInfo.LotId = objPO_DTL.LOT_NO
        Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)
        gMain.objHandling.O_GetDB_dicPO_MergeByPO_ID_PO_Serial_No(objPO_DTL.PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_Merge)
        'ProduceInDetailDataInfo.CheckQty = objPO_DTL.QTY_PROCESS
        If dicPO_Merge.Any Then
          ProduceInDetailDataInfo.CheckQty = dicPO_Merge.First.Value.QTY_FINISH
        Else
          Return False
        End If
        'ProduceInDetailDataInfo.SN = objPO_DTL.ITEM_COMMON5
        ProduceInDetailDataInfo.Item_Common3 = objPO_DTL.ITEM_COMMON3
        ProduceInDetailDataList.ProduceInDetailDataInfo.Add(ProduceInDetailDataInfo)
        FactoryId = objPO_DTL.TO_OWNER_ID
        warehouse = objPO_DTL.ITEM_COMMON10
      Next
      ProduceInDataInfo.FactoryId = FactoryId
      'ProduceInDataInfo.Warehouse = warehouse
      ProduceInDataInfo.ProduceInDetailDataList = ProduceInDetailDataList
      ProduceInDataList.ProduceInDataInfo.Add(ProduceInDataInfo)
      ProduceInDataReport.ProduceInDataList = ProduceInDataList

      '將物件轉成xml     
      If PrepareMessage_MSG(Of MSG_SendProduceInData)(StrXML, ProduceInDataReport, result_Message) = False Then
        If result_Message = "" Then
          result_Message = "轉XML錯誤(MSG_SendProduceInData_生產入庫單回報)"
        End If
        Return False
      End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 雜收單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function OtherInData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      Dim Now_time = GetNewTime_DBFormat()
      '表頭
      Dim OtherInDataReport As New MSG_SendOtherInData
      OtherInDataReport.WebService_ID = "WMS"
      OtherInDataReport.EventID = "OtherInData"
      Dim OtherInDataList As New MSG_SendOtherInData.clsOtherInDataList
      OtherInDataList.OtherInDataInfo = New List(Of MSG_SendOtherInData.clsOtherInDataList.clsOtherInDataInfo)
      Dim OtherInDataInfo As New MSG_SendOtherInData.clsOtherInDataList.clsOtherInDataInfo
      OtherInDataInfo.POType = ObjPO.PO_KEY1
      OtherInDataInfo.POId = ObjPO.PO_KEY2
      OtherInDataInfo.OtherInDateTime = Now_time
      OtherInDataInfo.FactoryId = ObjPO.H_PO3
      Dim FactoryId = ""
      Dim warehouse = ""
      '加入表身
      Dim OtherInDetailDataList As New MSG_SendOtherInData.clsOtherInDataList.clsOtherInDataInfo.clsOtherInDetailDataList
      OtherInDetailDataList.OtherInDetailDataInfo = New List(Of MSG_SendOtherInData.clsOtherInDataList.clsOtherInDataInfo.clsOtherInDetailDataList.clsOtherInDetailDataInfo)
      For Each objPO_DTL In dicPO_DTL.Values
        Dim OtherInDetailDataInfo As New MSG_SendOtherInData.clsOtherInDataList.clsOtherInDataInfo.clsOtherInDetailDataList.clsOtherInDetailDataInfo
        OtherInDetailDataInfo.SerialId = objPO_DTL.PO_SERIAL_NO
        Dim SKU_NO As String() = objPO_DTL.SKU_NO.Split("_")
        OtherInDetailDataInfo.SKU = SKU_NO(0)
        OtherInDetailDataInfo.LotId = objPO_DTL.LOT_NO
        Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)
        gMain.objHandling.O_GetDB_dicPO_MergeByPO_ID_PO_Serial_No(objPO_DTL.PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_Merge)
        If dicPO_Merge.Any Then
          OtherInDetailDataInfo.CheckQty = dicPO_Merge.First.Value.QTY_FINISH
        Else
          Return False
        End If
        OtherInDetailDataInfo.SN = objPO_DTL.ITEM_COMMON5
        OtherInDetailDataList.OtherInDetailDataInfo.Add(OtherInDetailDataInfo)
        warehouse = objPO_DTL.ITEM_COMMON10
        FactoryId = objPO_DTL.TO_OWNER_ID
      Next
      OtherInDataInfo.FactoryId = FactoryId
      OtherInDataInfo.Warehouse = warehouse
      OtherInDataInfo.OtherInDetailDataList = OtherInDetailDataList
      OtherInDataList.OtherInDataInfo.Add(OtherInDataInfo)
      OtherInDataReport.OtherInDataList = OtherInDataList

      '將物件轉成xml     
      If PrepareMessage_MSG(Of MSG_SendOtherInData)(StrXML, OtherInDataReport, result_Message) = False Then
        If result_Message = "" Then
          result_Message = "轉XML錯誤(MSG_SendOtherInData_雜收單回報)"
        End If
        Return False
      End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 調撥入庫單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function SellReturnData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      Dim Now_time = GetNewTime_DBFormat()
      '表頭
      Dim SellReturnDataReport As New MSG_SendSellReturnData
      SellReturnDataReport.WebService_ID = "WMS"
      SellReturnDataReport.EventID = "SellReturnData"
      Dim SellReturnDataList As New MSG_SendSellReturnData.clsSellReturnDataList
      SellReturnDataList.SellReturnDataInfo = New List(Of MSG_SendSellReturnData.clsSellReturnDataList.clsSellReturnDataInfo)
      Dim SellReturnDataInfo As New MSG_SendSellReturnData.clsSellReturnDataList.clsSellReturnDataInfo
      SellReturnDataInfo.POType = ObjPO.PO_KEY1
      SellReturnDataInfo.POId = ObjPO.PO_KEY2
      SellReturnDataInfo.SellReturnDateTime = Now_time
      SellReturnDataInfo.FactoryId = ObjPO.H_PO3
      Dim FactoryId = ""
      Dim warehouse = ""
      '加入表身
      'Dim TransferDetailDataList As New MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo.clsTransferDetailDataList
      Dim SellReturnDetailDataList As New MSG_SendSellReturnData.clsSellReturnDataList.clsSellReturnDataInfo.clsSellReturnDetailDataList
      SellReturnDetailDataList.SellReturnDetailDataInfo = New List(Of MSG_SendSellReturnData.clsSellReturnDataList.clsSellReturnDataInfo.clsSellReturnDetailDataList.clsSellReturnDetailDataInfo)
      'TransferDetailDataList.TransferDetailDataInfo = New List(Of MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo.clsTransferDetailDataList.clsTransferDetailDataInfo)
      For Each objPO_DTL In dicPO_DTL.Values
        'Dim TransferDetailDataInfo As New MSG_SendTransferData.clsTransferDataList.clsTransferDataInfo.clsTransferDetailDataList.clsTransferDetailDataInfo
        Dim SellReturnDetailDataInfo As New MSG_SendSellReturnData.clsSellReturnDataList.clsSellReturnDataInfo.clsSellReturnDetailDataList.clsSellReturnDetailDataInfo
        SellReturnDetailDataInfo.SerialId = objPO_DTL.PO_SERIAL_NO
        Dim SKU_NO As String() = objPO_DTL.SKU_NO.Split("_")

        SellReturnDetailDataInfo.SKU = SKU_NO(0)
        SellReturnDetailDataInfo.LotId = objPO_DTL.LOT_NO
        Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)
        gMain.objHandling.O_GetDB_dicPO_MergeByPO_ID_PO_Serial_No(objPO_DTL.PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_Merge)
        If dicPO_Merge.Any Then
          SellReturnDetailDataInfo.CheckQty = dicPO_Merge.First.Value.QTY_FINISH
        Else
          Return False
        End If
        SellReturnDetailDataList.SellReturnDetailDataInfo.Add(SellReturnDetailDataInfo)
        warehouse = objPO_DTL.ITEM_COMMON10
        FactoryId = objPO_DTL.TO_OWNER_ID
      Next
      SellReturnDataInfo.FactoryId = FactoryId
      SellReturnDataInfo.Warehouse = warehouse
      SellReturnDataInfo.SellReturnDetailDataList = SellReturnDetailDataList
      SellReturnDataList.SellReturnDataInfo.Add(SellReturnDataInfo)
      SellReturnDataReport.SellReturnDataList = SellReturnDataList
      '將物件轉成xml     
      If PrepareMessage_MSG(Of MSG_SendSellReturnData)(StrXML, SellReturnDataReport, result_Message) = False Then
        If result_Message = "" Then
          result_Message = "轉XML錯誤(SendSellReturnData_雕撥入庫單回報)"
        End If
        Return False
      End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 雜發單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function OtherOutData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      Dim Now_time = GetNewTime_DBFormat()
      '表頭
      Dim OtherOutDataReport As New MSG_SendOtherOutData
      OtherOutDataReport.WebService_ID = "WMS"
      OtherOutDataReport.EventID = "OtherOutData"
      Dim OtherOutDataList As New MSG_SendOtherOutData.clsOtherOutDataList
      OtherOutDataList.OtherOutDataInfo = New List(Of MSG_SendOtherOutData.clsOtherOutDataList.clsOtherOutDataInfo)
      Dim OtherOutDataInfo As New MSG_SendOtherOutData.clsOtherOutDataList.clsOtherOutDataInfo
      OtherOutDataInfo.POType = ObjPO.PO_KEY1
      OtherOutDataInfo.POId = ObjPO.PO_KEY2
      OtherOutDataInfo.OtherOutDateTime = Now_time
      OtherOutDataInfo.FactoryId = ObjPO.H_PO3
      Dim FactoryId = ""
      Dim warehouse = ""
      '加入表身
      Dim OtherOutDetailDataList As New MSG_SendOtherOutData.clsOtherOutDataList.clsOtherOutDataInfo.clsOtherOutDetailDataList
      OtherOutDetailDataList.OtherOutDetailDataInfo = New List(Of MSG_SendOtherOutData.clsOtherOutDataList.clsOtherOutDataInfo.clsOtherOutDetailDataList.clsOtherOutDetailDataInfo)
      For Each objPO_DTL In dicPO_DTL.Values
        Dim OtherOutDetailDataInfo As New MSG_SendOtherOutData.clsOtherOutDataList.clsOtherOutDataInfo.clsOtherOutDetailDataList.clsOtherOutDetailDataInfo
        OtherOutDetailDataInfo.SerialId = objPO_DTL.PO_SERIAL_NO
        Dim SKU_NO As String() = objPO_DTL.SKU_NO.Split("_")
        OtherOutDetailDataInfo.SKU = SKU_NO(0)
        OtherOutDetailDataInfo.LotId = objPO_DTL.LOT_NO
        Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)
        gMain.objHandling.O_GetDB_dicPO_MergeByPO_ID_PO_Serial_No(objPO_DTL.PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_Merge)
        If dicPO_Merge.Any Then
          OtherOutDetailDataInfo.CheckQty = dicPO_Merge.First.Value.QTY_FINISH
        Else
          Return False
        End If
        OtherOutDetailDataInfo.SN = objPO_DTL.ITEM_COMMON5
        OtherOutDetailDataList.OtherOutDetailDataInfo.Add(OtherOutDetailDataInfo)
        warehouse = objPO_DTL.ITEM_COMMON10
        FactoryId = objPO_DTL.FROM_OWNER_ID
      Next
      OtherOutDataInfo.FactoryId = FactoryId
      OtherOutDataInfo.Warehouse = warehouse
      OtherOutDataInfo.OtherOutDetailDataList = OtherOutDetailDataList
      OtherOutDataList.OtherOutDataInfo.Add(OtherOutDataInfo)
      OtherOutDataReport.OtherOutDataList = OtherOutDataList

      '將物件轉成xml     
      If PrepareMessage_MSG(Of MSG_SendOtherOutData)(StrXML, OtherOutDataReport, result_Message) = False Then
        If result_Message = "" Then
          result_Message = "轉XML錯誤(MSG_SendOtherOutData_雜發單回報)"
        End If
        Return False
      End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 銷貨單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function SellData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      'Dim Now_time = GetNewTime_DBFormat()
      ''表頭
      'Dim SellDataReport As New MSG_SendSellData
      'SellDataReport.WebService_ID = "ERP"
      'SellDataReport.EventID = "SellData"
      'Dim SellDataList As New MSG_SendSellData.clsSellDataList
      'SellDataList.SellDataInfo = New List(Of MSG_SendSellData.clsSellDataList.clsSellDataInfo)
      'Dim SellDataInfo As New MSG_SendSellData.clsSellDataList.clsSellDataInfo
      'SellDataInfo.POType = ObjPO.PO_KEY1
      'SellDataInfo.POId = ObjPO.PO_KEY2
      'SellDataInfo.SellDateTime = Now_time
      'SellDataInfo.FactoryId = ObjPO.H_PO3
      'Dim FactoryId = ""
      'Dim warehouse = ""
      ''加入表身
      'Dim SellDetailDataList As New MSG_SendSellData.clsSellDataList.clsSellDataInfo.clsSellDetailDataList
      'SellDetailDataList.SellDetailDataInfo = New List(Of MSG_SendSellData.clsSellDataList.clsSellDataInfo.clsSellDetailDataList.clsSellDetailDataInfo)
      'For Each objPO_DTL In dicPO_DTL.Values
      '  Dim SellDetailDataInfo As New MSG_SendSellData.clsSellDataList.clsSellDataInfo.clsSellDetailDataList.clsSellDetailDataInfo
      '  SellDetailDataInfo.SerialId = objPO_DTL.PO_SERIAL_NO
      '  Dim SKU_NO As String() = objPO_DTL.SKU_NO.Split("_")
      '  SellDetailDataInfo.SKU = SKU_NO(0)
      '  SellDetailDataInfo.LotId = objPO_DTL.LOT_NO
      '  Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)
      '  gMain.objHandling.O_GetDB_dicPO_MergeByPO_ID_PO_Serial_No(objPO_DTL.PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_Merge)
      '  If dicPO_Merge.Any Then
      '    SellDetailDataInfo.CheckQty = dicPO_Merge.First.Value.QTY_FINISH
      '  Else
      '    Return False
      '  End If
      '  SellDetailDataInfo.SN = objPO_DTL.ITEM_COMMON5
      '  SellDetailDataList.SellDetailDataInfo.Add(SellDetailDataInfo)
      '  warehouse = objPO_DTL.ITEM_COMMON10
      '  FactoryId = objPO_DTL.FROM_OWNER_ID
      'Next
      'SellDataInfo.FactoryId = FactoryId
      'SellDataInfo.Warehouse = warehouse
      'SellDataInfo.SellDetailDataList = SellDetailDataList
      'SellDataList.SellDataInfo.Add(SellDataInfo)
      'SellDataReport.SellDataList = SellDataList

      ''將物件轉成xml     
      'If PrepareMessage_MSG(Of MSG_SendSellData)(StrXML, SellDataReport, result_Message) = False Then
      '  If result_Message = "" Then
      '    result_Message = "轉XML錯誤(MSG_SendSellData_銷貨單回報)"
      '  End If
      '  Return False
      'End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 銷貨單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function InboundReturnData_Report(ByRef ObjPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      'Dim Now_time = GetNewTime_DBFormat()
      ''表頭
      'Dim InboundReturnDataReport As New MSG_SendInboundReturnData
      'InboundReturnDataReport.WebService_ID = "WMS"
      'InboundReturnDataReport.EventID = "InboundReturnData"
      'Dim InboundReturnDataList As New MSG_SendInboundReturnData.clsInboundReturnDataList
      'InboundReturnDataList.InboundReturnDataInfo = New List(Of MSG_SendInboundReturnData.clsInboundReturnDataList.clsInboundReturnDataInfo)
      'Dim InboundReturnDataInfo As New MSG_SendInboundReturnData.clsInboundReturnDataList.clsInboundReturnDataInfo
      'InboundReturnDataInfo.POType = ObjPO.PO_KEY1
      'InboundReturnDataInfo.POId = ObjPO.PO_KEY2
      'InboundReturnDataInfo.InboundReturnDateTime = Now_time
      'InboundReturnDataInfo.FactoryId = ObjPO.H_PO3
      'Dim FactoryId = ""
      'Dim warehouse = ""
      ''加入表身
      'Dim SellDetailDataList As New MSG_SendSellData.clsSellDataList.clsSellDataInfo.clsSellDetailDataList
      'SellDetailDataList.SellDetailDataInfo = New List(Of MSG_SendSellData.clsSellDataList.clsSellDataInfo.clsSellDetailDataList.clsSellDetailDataInfo)
      'Dim InboundReturnDetailDataList As New MSG_SendInboundReturnData.clsInboundReturnDataList.clsInboundReturnDataInfo.clsInboundReturnDetailDataList
      'InboundReturnDetailDataList.InboundReturnDetailDataInfo = New List(Of MSG_SendInboundReturnData.clsInboundReturnDataList.clsInboundReturnDataInfo.clsInboundReturnDetailDataList.clsInboundReturnDetailDataInfo)
      'For Each objPO_DTL In dicPO_DTL.Values

      '  Dim InboundReturnDetailDataInfo As New MSG_SendInboundReturnData.clsInboundReturnDataList.clsInboundReturnDataInfo.clsInboundReturnDetailDataList.clsInboundReturnDetailDataInfo
      '  InboundReturnDetailDataInfo.SerialId = objPO_DTL.PO_SERIAL_NO
      '  Dim SKU_NO As String() = objPO_DTL.SKU_NO.Split("_")
      '  InboundReturnDetailDataInfo.SKU = SKU_NO(0)
      '  InboundReturnDetailDataInfo.SN = objPO_DTL.PACKAGE_ID
      '  Dim dicPO_Merge As New Dictionary(Of String, clsPO_MERGE)
      '  gMain.objHandling.O_GetDB_dicPO_MergeByPO_ID_PO_Serial_No(objPO_DTL.PO_ID, objPO_DTL.PO_SERIAL_NO, dicPO_Merge)
      '  If dicPO_Merge.Any Then
      '    InboundReturnDetailDataInfo.CheckQty = dicPO_Merge.First.Value.QTY_FINISH
      '  Else
      '    Return False
      '  End If
      '  InboundReturnDetailDataList.InboundReturnDetailDataInfo.Add(InboundReturnDetailDataInfo)
      '  warehouse = objPO_DTL.ITEM_COMMON10
      '  FactoryId = objPO_DTL.FROM_OWNER_ID
      'Next
      'InboundReturnDataInfo.FactoryId = FactoryId
      'InboundReturnDataInfo.Warehouse = warehouse
      'InboundReturnDataInfo.InboundReturnDetailDataList = InboundReturnDetailDataList
      'InboundReturnDataList.InboundReturnDataInfo.Add(InboundReturnDataInfo)
      'InboundReturnDataReport.InboundReturnDataList = InboundReturnDataList
      ''將物件轉成xml     
      'If PrepareMessage_MSG(Of MSG_SendInboundReturnData)(StrXML, InboundReturnDataReport, result_Message) = False Then
      '  If result_Message = "" Then
      '    result_Message = "轉XML錯誤(MSG_SendSellData_銷貨單回報)"
      '  End If
      '  Return False
      'End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 貨主調撥單回報
  ''' </summary>
  ''' <param name="ObjPO"></param>
  ''' <param name="dicPO_DTL"></param>
  ''' <param name="StrXML"></param>
  ''' <param name="result_Message"></param>
  ''' <returns></returns>
  Public Function TransferOwnerData_Report(ByRef objPO As clsPO, ByRef dicPO_DTL As Dictionary(Of String, clsPO_DTL), ByRef dicPO_Merge As Dictionary(Of String, clsPO_MERGE), ByRef StrXML As String, ByRef result_Message As String) As Boolean
    Try
      Dim Now_time = GetNewTime_DBFormat()
      '表頭
      Dim TransferOwnerDataReport As New MSG_SendTransferOwnerData
      TransferOwnerDataReport.WebService_ID = "WMS"
      TransferOwnerDataReport.EventID = "TransferOwnerData"
      Dim TransferOwnerDataList As New MSG_SendTransferOwnerData.clsTransferOwnerDataList
      TransferOwnerDataList.TransferOwnerDataInfo = New List(Of MSG_SendTransferOwnerData.clsTransferOwnerDataList.clsTransferOwnerDataInfo)
      Dim TransferOwnerDataInfo As New MSG_SendTransferOwnerData.clsTransferOwnerDataList.clsTransferOwnerDataInfo
      TransferOwnerDataInfo.POType = objPO.PO_KEY1
      TransferOwnerDataInfo.POId = objPO.PO_KEY2
      TransferOwnerDataInfo.TransferOwnerDateTime = Now_time
      TransferOwnerDataInfo.FactoryId = objPO.H_PO3
      TransferOwnerDataInfo.TransferOutOwner = objPO.H_PO7
      TransferOwnerDataInfo.TransferInOwner = objPO.H_PO8
      Dim FactoryId = ""
      '加入表身
      Dim TransferOwnerDetailDataList As New MSG_SendTransferOwnerData.clsTransferOwnerDataList.clsTransferOwnerDataInfo.clsTransferOwnerDetailDataList
      TransferOwnerDetailDataList.TransferOwnerDetailDataInfo = New List(Of MSG_SendTransferOwnerData.clsTransferOwnerDataList.clsTransferOwnerDataInfo.clsTransferOwnerDetailDataList.clsTransferOwnerDetailDataInfo)
      For Each objPO_DTL In dicPO_DTL.Values
        Dim TransferOwnerDetailDataInfo As New MSG_SendTransferOwnerData.clsTransferOwnerDataList.clsTransferOwnerDataInfo.clsTransferOwnerDetailDataList.clsTransferOwnerDetailDataInfo
        TransferOwnerDetailDataInfo.SerialId = objPO_DTL.PO_SERIAL_NO
        Dim SKU_NO As String() = objPO_DTL.SKU_NO.Split("_")
        TransferOwnerDetailDataInfo.SKU = SKU_NO(0)
        TransferOwnerDetailDataInfo.LotId = objPO_DTL.LOT_NO
        Dim Report_QTY = 0
        For Each obj In dicPO_Merge.Values
          If obj.PO_ID = objPO_DTL.PO_ID And obj.PO_SERIAL_NO = objPO_DTL.PO_SERIAL_NO Then
            Report_QTY = obj.QTY_FINISH
          End If
        Next
        TransferOwnerDetailDataInfo.CheckQty = Report_QTY 'objPO_DTL.QTY_FINISH
        TransferOwnerDetailDataInfo.SN = objPO_DTL.PACKAGE_ID

        'TransferOwnerDetailDataInfo.Item_Common3 = objPO_DTL.ITEM_COMMON3
        'TransferDetailDataInfo.TransferOutWarehouse = objPO_DTL.H_POD8
        'TransferDetailDataInfo.TransferInWarehouse = objPO_DTL.H_POD9
        If objPO.PO_Type1 = enuPOType_1.Combination_in Then
          FactoryId = objPO_DTL.TO_OWNER_ID
        Else
          FactoryId = objPO_DTL.FROM_OWNER_ID
        End If

        TransferOwnerDetailDataList.TransferOwnerDetailDataInfo.Add(TransferOwnerDetailDataInfo)
      Next
      TransferOwnerDataInfo.FactoryId = FactoryId
      TransferOwnerDataInfo.TransferOwnerDetailDataList = TransferOwnerDetailDataList
      TransferOwnerDataList.TransferOwnerDataInfo.Add(TransferOwnerDataInfo)
      TransferOwnerDataReport.TransferOwnerDataList = TransferOwnerDataList

      '將物件轉成xml     
      If PrepareMessage_MSG(Of MSG_SendTransferOwnerData)(StrXML, TransferOwnerDataReport, result_Message) = False Then
        If result_Message = "" Then
          result_Message = "轉XML錯誤(SendTransferOwnerData_貨主雕撥回報)"
        End If
        Return False
      End If
      Return True
    Catch ex As Exception
      result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module
