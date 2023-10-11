'20180629
'V1.0.0
'Jerry

'提單(強制或不強制)

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T11F1U1_PODownload
  Public Function O_T11F1U1_PODownload(ByVal Receive_Msg As MSG_T11F1U1_PODownload,
                                       ByRef ret_strResultMsg As String,
                                       ByRef ret_Wait_UUID As String) As Boolean
    Try
      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料處理
      If Process_Data(Receive_Msg, ret_strResultMsg, ret_Wait_UUID) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T11F1U1_PODownload,
                              ByRef ret_strResultMsg As String) As Boolean

    Try
      '先進行資料邏輯檢查
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        Dim FORCED_UPDATE As String = objPOInfo.FORCED_UPDATE
        '檢查PO_ID是否為空
        If PO_ID = "" Then
          ret_strResultMsg = "PO_ID is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '資料處理
  Private Function Process_Data(ByVal Receive_Msg As MSG_T11F1U1_PODownload,
                                ByRef ret_strResultMsg As String, ByRef ret_Wait_UUID As String) As Boolean
    Try
      Dim USER_ID As String = Receive_Msg.Header.ClientInfo.UserID
      '先進行資料邏輯檢查
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        Dim PO_ID As String = objPOInfo.PO_ID
        If PO_ID.Contains("-") = False Then
          ret_strResultMsg = "單號格式錯誤"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        Dim PO_ID_Str As String() = PO_ID.Split("-")
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        Dim FORCED_UPDATE As String = objPOInfo.FORCED_UPDATE
        Dim FORCED_FLAG As Boolean = False
        Dim dicPO_ID As New Dictionary(Of String, String)
        Dim tmp_PO_ID = PO_ID
        Dim ERP_ORDER_TYPE = objPOInfo.ERP_ORDER_TYPE

        Dim tmp_dicPOID As New Dictionary(Of String, String)
        Dim tmp_dicPO As New Dictionary(Of String, clsPO)
        Dim tmp_dicPO_Line As New Dictionary(Of String, clsPO_LINE)
        Dim tmp_dicPO_DTL As New Dictionary(Of String, clsPO_DTL)

        If tmp_dicPOID.ContainsKey(PO_ID) = False Then
          tmp_dicPOID.Add(PO_ID, PO_ID)
        End If
        '使用dicPO取得資料庫裡的PO資料
        If gMain.objHandling.O_GetDB_dicHPOBydicPO_ID(tmp_dicPOID, tmp_dicPO) = False Then
          ret_strResultMsg = "WMS get PO data From DB Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If tmp_dicPO.Any Then
          ret_strResultMsg = $"單據:{PO_ID_Str(1)}，已執行過，勿重覆提單。"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '使用dicPO取得資料庫裡的PO資料
        If gMain.objHandling.O_GetDB_dicPOBydicPO_ID(tmp_dicPOID, tmp_dicPO) = False Then
          ret_strResultMsg = "WMS get PO data From DB Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '使用dicPO取得資料庫裡的PO_Line資料
        If gMain.objHandling.O_GetDB_dicPOLineBydicPO_ID(tmp_dicPOID, tmp_dicPO_Line) = False Then
          ret_strResultMsg = "WMS get PO_Line data From DB Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        '使用dicPO取得資料庫裡的PO_DTL資料
        If gMain.objHandling.O_GetDB_dicPODTLBydicPO_ID(tmp_dicPOID, tmp_dicPO_DTL) = False Then
          ret_strResultMsg = "WMS get PO_DTL data From DB Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If

        '20221012 修改採購單和製令單的取法
        Dim dicPURTC As New Dictionary(Of String, clsPURTC)
        Dim dicPURTD As New Dictionary(Of String, clsPURTD)
        Dim dicPURTE As New Dictionary(Of String, clsPURTE)
        Dim dicPURTF As New Dictionary(Of String, clsPURTF)

        Dim dicMOCTA As New Dictionary(Of String, clsMOCTA)
        Dim dicMOCTB As New Dictionary(Of String, clsMOCTB)
        Dim dicMOCTO As New Dictionary(Of String, clsMOCTO)
        Dim dicMOCTP As New Dictionary(Of String, clsMOCTP)

        '20221012 因為客戶無法用單號確認單別，因此改用客戶輸入字串用"-"切割後的中間段=單號
        '用單號去同時查採購和製令，客戶確定單號不會重複，因此只會有一邊查到表頭資料
        '接著取出表頭完整資訊再用單號 + 單別 去查表身(若是採購，客戶切割後第三段文字是序號，也要當成條件)
        '程式製作過程中，已將採購和製令的PO_ID改成客戶上傳字串

        Dim split_PO_TYPE = ""
        Dim split_PO_ID = ""
        Dim split_Serial = ""

        split_PO_TYPE = PO_ID_Str(0).Trim
        split_PO_ID = PO_ID_Str(1).Trim
        If (PO_ID_Str.Length = 3) Then
          split_Serial = PO_ID_Str(2).Trim
        End If

        If tmp_dicPO.Any = False Then
          '無單據，讀取新增用的 採購和製令單 TABLE
          gMain.objHandling.O_GetDB_dicPURTCByPOID(split_PO_TYPE, split_PO_ID, dicPURTC)

          If dicPURTC.Count = 0 Then '沒有採購才讀取製令
            gMain.objHandling.O_GetDB_dicMOCTAByPOID(split_PO_TYPE, split_PO_ID, dicMOCTA)
          End If
        Else
          '有單據，讀取修改用的 TABLE
          gMain.objHandling.O_GetDB_dicPURTEByPOID(split_PO_TYPE, split_PO_ID, dicPURTE)

          If dicPURTE.Count = 0 Then '沒有採購才讀取製令
            gMain.objHandling.O_GetDB_dicMOCTOByPOID(split_PO_TYPE, split_PO_ID, dicMOCTO)
          End If
        End If


        '若採購單和製令單的表頭都查不到東西，代表單號有問題
        If dicPURTC.Count = 0 AndAlso dicPURTE.Count = 0 AndAlso dicMOCTA.Count = 0 AndAlso dicMOCTO.Count = 0 Then
          If tmp_dicPO.Any = False Then
            ret_strResultMsg = $"新增用表頭檔查無單據，單據號:{PO_ID}"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          Else
            ret_strResultMsg = $"修改用表頭檔查無單據，單據號:{PO_ID}"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If

        Else
          If dicPURTC.Any Then '採購單，新增
            If PO_ID_Str.Length <> 3 Then
              SendMessageToLog($"PO_ID:{PO_ID}，用 '-' 切割後矩陣長度不為3，無法取得採購單序號", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If

            gMain.objHandling.O_GetDB_dicPURTDByPOID(split_PO_TYPE, split_PO_ID, split_Serial, dicPURTD)

            If Module_POManagement_HTG_Buy.O_POManagement_HTG_Buy(tmp_dicPOID, tmp_dicPO, tmp_dicPO_Line, tmp_dicPO_DTL, dicPURTC, dicPURTD, dicPURTE, dicPURTF, USER_ID, ret_strResultMsg, ret_Wait_UUID) = False Then
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If


          ElseIf dicPURTE.Any Then '採購單，修改
            If PO_ID_Str.Length <> 3 Then
              SendMessageToLog($"PO_ID:{PO_ID}，用 '-' 切割後矩陣長度不為3，無法取得採購單序號", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If

            gMain.objHandling.O_GetDB_dicPURTFByPOID(split_PO_TYPE, split_PO_ID, split_Serial, dicPURTF)

            '檢查單據狀態是否可以修改。已處理數量大於修改後的項目，則不能修改
            '有錯誤
            For Each objPO_DTL In tmp_dicPO_DTL.Values
              For Each objPURTF In dicPURTF.Values
                If objPO_DTL.PO_SERIAL_NO = objPURTF.TF004 Then
                  If objPO_DTL.QTY_PROCESS <> 0 And objPO_DTL.QTY_PROCESS > objPURTF.TF009 Then
                    ret_strResultMsg = "項次：" & objPO_DTL.PO_SERIAL_NO & ",已處理數量:" & objPO_DTL.QTY_PROCESS & ",最後單據數量：" & objPURTF.TF009
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
                  Exit For
                End If
              Next
            Next
            '無錯誤，改單據
            If Module_POManagement_HTG_Buy.O_POManagement_HTG_Buy(tmp_dicPOID, tmp_dicPO, tmp_dicPO_Line, tmp_dicPO_DTL, dicPURTC, dicPURTD, dicPURTE, dicPURTF, USER_ID, ret_strResultMsg, ret_Wait_UUID) = False Then
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If


          ElseIf dicMOCTA.Any Then '製令單，新增
            gMain.objHandling.O_GetDB_dicMOCTBByPOID(split_PO_TYPE, split_PO_ID, dicMOCTB)

            If Module_POManagement_HTG_Produce_In.O_POManagement_HTG_Produce_In(tmp_dicPOID, tmp_dicPO, tmp_dicPO_Line, tmp_dicPO_DTL, dicMOCTA, dicMOCTB, dicMOCTO, dicMOCTP, USER_ID, ret_strResultMsg, ret_Wait_UUID) = False Then
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If


          ElseIf dicMOCTO.Any Then '製令單，修改
            gMain.objHandling.O_GetDB_dicMOCTPByPOID(split_PO_TYPE, split_PO_ID, dicMOCTP)

            '檢查單據狀態是否可以修改。已處理數量大於修改後的項目，則不能修改
            '有錯誤
            For Each objPO_DTL In tmp_dicPO_DTL.Values
              For Each objMOCTP In dicMOCTP.Values
                If objPO_DTL.SKU_NO = objMOCTP.TP004 Then
                  If objPO_DTL.QTY_PROCESS <> 0 And objPO_DTL.QTY_PROCESS > objMOCTP.TP005 Then
                    ret_strResultMsg = "項次：" & objPO_DTL.PO_SERIAL_NO & ",已處理數量:" & objPO_DTL.QTY_PROCESS & ",最後單據數量：" & objMOCTP.TP005
                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                    Return False
                  End If
                  Exit For
                End If
              Next
            Next
            '無錯誤，改單據
            If Module_POManagement_HTG_Produce_In.O_POManagement_HTG_Produce_In(tmp_dicPOID, tmp_dicPO, tmp_dicPO_Line, tmp_dicPO_DTL, dicMOCTA, dicMOCTB, dicMOCTO, dicMOCTP, USER_ID, ret_strResultMsg, ret_Wait_UUID) = False Then
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If
          End If
        End If 'If dicPURTC.Count = 0 AndAlso dicPURTE.Count = 0 AndAlso dicMOCTA.Count = 0 AndAlso dicMOCTO.Count = 0 Then

        '        If H_PO_ORDER_TYPE = enuOrderType.Inbound_Data Then
        '          If PO_ID_Str.Length <> 3 Then
        '            SendMessageToLog($"PO_ID:{PO_ID}，用 '_' 切割後矩陣長度不為3", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '            Return False
        '          End If

        '#Region "採購單"
        '          '無單據則進行提單
        '          Dim dicPURTC As New Dictionary(Of String, clsPURTC)
        '          Dim dicPURTD As New Dictionary(Of String, clsPURTD)
        '          Dim dicPURTE As New Dictionary(Of String, clsPURTE)
        '          Dim dicPURTF As New Dictionary(Of String, clsPURTF)

        '          If tmp_dicPO.Any = False Then
        '            '無單據，新增PO
        '            gMain.objHandling.O_GetDB_dicPURTCByPOID(PO_ID_Str(1), dicPURTC)
        '            gMain.objHandling.O_GetDB_dicPURTDByPOID(PO_ID_Str(1), PO_ID_Str(2), dicPURTD)
        '            If Module_POManagement_HTG_Buy.O_POManagement_HTG_Buy(tmp_dicPOID, tmp_dicPO, tmp_dicPO_Line, tmp_dicPO_DTL, dicPURTC, dicPURTD, dicPURTE, dicPURTF, USER_ID, ret_strResultMsg) = False Then
        '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '              Return False
        '            End If
        '          Else
        '            '有單據，判斷項次數量是否有錯誤(已處理數量大最新單據數量)

        '            gMain.objHandling.O_GetDB_dicPURTEByPOID(PO_ID_Str(1), dicPURTE)
        '            gMain.objHandling.O_GetDB_dicPURTFByPOID(PO_ID_Str(1), PO_ID_Str(2), dicPURTF)

        '            '有錯誤
        '            For Each objPO_DTL In tmp_dicPO_DTL.Values
        '              For Each objPURTF In dicPURTF.Values
        '                If objPO_DTL.PO_SERIAL_NO = objPURTF.TF004 Then
        '                  If objPO_DTL.QTY_PROCESS <> 0 And objPO_DTL.QTY_PROCESS > objPURTF.TF009 Then
        '                    ret_strResultMsg = "項次：" & objPO_DTL.PO_SERIAL_NO & ",已處理數量:" & objPO_DTL.QTY_PROCESS & ",最後單據數量：" & objPURTF.TF009
        '                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '                    Return False
        '                  End If
        '                  Exit For
        '                End If
        '              Next

        '            Next
        '            '無錯誤，改單據
        '            If Module_POManagement_HTG_Buy.O_POManagement_HTG_Buy(tmp_dicPOID, tmp_dicPO, tmp_dicPO_Line, tmp_dicPO_DTL, dicPURTC, dicPURTD, dicPURTE, dicPURTF, USER_ID, ret_strResultMsg) = False Then
        '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '              Return False
        '            End If
        '          End If
        '#End Region

        '        ElseIf H_PO_ORDER_TYPE = enuOrderType.Product_In Then
        '#Region "製令單(成品入庫單)"
        '          Dim dicMOCTA As New Dictionary(Of String, clsMOCTA)
        '          Dim dicMOCTB As New Dictionary(Of String, clsMOCTB)
        '          Dim dicMOCTP As New Dictionary(Of String, clsMOCTP)
        '          Dim dicMOCTO As New Dictionary(Of String, clsMOCTO)
        '          '無單據則進行提單
        '          If tmp_dicPO.Any = False Then

        '            gMain.objHandling.O_GetDB_dicMOCTAByPOID(PO_ID_Str(1), dicMOCTA)
        '            gMain.objHandling.O_GetDB_dicMOCTBByPOID(PO_ID_Str(1), dicMOCTB)
        '            If Module_POManagement_HTG_Produce_In.O_POManagement_HTG_Produce_In(tmp_dicPOID, tmp_dicPO, tmp_dicPO_Line, tmp_dicPO_DTL, dicMOCTA, dicMOCTB, dicMOCTO, dicMOCTP, USER_ID, ret_strResultMsg) = False Then
        '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '              Return False
        '            End If
        '          Else
        '            '有單據，判斷項次數量是否有錯誤(已處理數量大最新單據數量)
        '            gMain.objHandling.O_GetDB_dicMOCTOByPOID(PO_ID_Str(1), dicMOCTO)
        '            gMain.objHandling.O_GetDB_dicMOCTPByPOID(PO_ID_Str(1), dicMOCTP)


        '            '有錯誤
        '            For Each objPO_DTL In tmp_dicPO_DTL.Values
        '              For Each objMOCTP In dicMOCTP.Values
        '                If objPO_DTL.SKU_NO = objMOCTP.TP004 Then
        '                  If objPO_DTL.QTY_PROCESS <> 0 And objPO_DTL.QTY_PROCESS > objMOCTP.TP005 Then
        '                    ret_strResultMsg = "項次：" & objPO_DTL.PO_SERIAL_NO & ",已處理數量:" & objPO_DTL.QTY_PROCESS & ",最後單據數量：" & objMOCTP.TP005
        '                    SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '                    Return False
        '                  End If
        '                  Exit For
        '                End If
        '              Next

        '            Next
        '            '無錯誤，改單據
        '            If Module_POManagement_HTG_Produce_In.O_POManagement_HTG_Produce_In(tmp_dicPOID, tmp_dicPO, tmp_dicPO_Line, tmp_dicPO_DTL, dicMOCTA, dicMOCTB, dicMOCTO, dicMOCTP, USER_ID, ret_strResultMsg) = False Then
        '              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '              Return False
        '            End If
        '          End If
        '#End Region
        '        End If






        If ERP_ORDER_TYPE = "VC" Then

        Else
          SendMessageToLog("Start Process Data T11F1U1_PODownload", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)     'Vito_12b30
          SyncLock gMain.objHandling.objCT_PO_DTLLock
            SendMessageToLog("T11F1U1_PODownload objCT_PO_DTLLock Locked", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)  'Vito_12b30
            Try
              '刪除對應的單據資訊
              Dim dicDeleteCT_PO_DTL = gMain.objHandling.gdicCT_PO_DTL.Where(Function(q)
                                                                               If q.Value.PO_ID = PO_ID Then Return True
                                                                               Return False
                                                                             End Function).ToDictionary(Function(q) q.Key, Function(q) q.Value)
              Dim lstSQL As New List(Of String)
              For Each obj In dicDeleteCT_PO_DTL.Values
                obj.O_Add_Delete_SQLString(lstSQL)
              Next
              If Common_DBManagement.BatchUpdate(lstSQL) = True Then
                For Each obj In dicDeleteCT_PO_DTL.Values
                  obj.Remove_Relationship()
                Next
              End If
            Catch ex As Exception
              SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            End Try
          End SyncLock
          SendMessageToLog("End Process Data T11F1U1_PODownload", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)     'Vito_12b30
          Dim Str_USER = ""
          'If Mod_WCFHost.ASRS_singleMatDoc(PO_ID, ret_strResultMsg, Str_USER) = True Then
          '  ret_strResultMsg = Str_USER
          '  Return True
          'End If
        End If
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function



End Module
