'20200106
'V1.0.0
'Vito
'Vito_20106
'WMS向HostHandler進行收料資訊的回報

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T10F2S1_StocktakingReport
  Public Function O_T10F2S1_StocktakingReport(ByVal Receive_Msg As MSG_T10F2S1_StocktakingReport,
                                          ByRef ret_strResultMsg As String) As Boolean
    Try

      Dim lstSql As New List(Of String)
      Dim InventoryData As New MSG_InventoryData
      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      Dim Stocktaking_ID As String = ""
      '進行資料處理
      If Process_Data(Receive_Msg, ret_strResultMsg, InventoryData, Stocktaking_ID) = False Then
        Return False
      End If

      Dim strXML As String = ""
      If PrepareMessage_InventoryToERP(strXML, InventoryData, ret_strResultMsg) = False Then
        Return False '將obj轉成xml
      End If
      If STD_IN(strXML, ret_strResultMsg) = False Then
        SendMessageToLog("盤點單回報失敗，盤點單號:" & Stocktaking_ID, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        Return False
      End If
      'Dim Json_Str = Obj_To_Json(GainHowStocktakingReport)
      'If HttpHost_Test(Json_Str, "T10F2S1_StocktakingReport", ret_strResultMsg) = False Then
      '  ret_strResultMsg = "上報ERP錯誤"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      'If Get_SQL(ret_strResultMsg, lstSql) = False Then
      '  Return False
      'End If
      'If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
      '  Return False
      'End If

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '檢查相關資料是否正確
  Private Function Check_Data(ByVal Receive_Msg As MSG_T10F2S1_StocktakingReport,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      For Each objStocktaking In Receive_Msg.Body.StocktakingList.StocktakingInfo
        Dim STOCKTAKING_ID = objStocktaking.STOCKTAKING_ID
        'Dim STOCKTAKING_SERIAL_NO = objStocktaking.STOCKTAKING_SERIAL_NO
        If STOCKTAKING_ID = "" Then
          ret_strResultMsg = "STOCKTAKING_ID is empty"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        'If STOCKTAKING_SERIAL_NO = "" Then
        '  ret_strResultMsg = "STOCKTAKING_SERIAL_NO is empty"
        '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        '  Return False
        'End If
      Next


      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '資料處理
  Private Function Process_Data(ByVal Receive_Msg As MSG_T10F2S1_StocktakingReport,
                                ByRef ret_strResultMsg As String, ByRef StocktakingReport As MSG_InventoryData, Optional ByRef Stocktaking_ID As String = "") As Boolean
    Try
      '先進行資料邏輯檢查
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim USER_ID = Receive_Msg.Header.ClientInfo.UserID
      Dim UUID = Receive_Msg.Header.UUID

      '表頭
      StocktakingReport.WebService_ID = "WMS"
      StocktakingReport.EventID = "InventoryData"
      Dim InventoryDataList As New MSG_InventoryData.clsInventoryDataList
      InventoryDataList.InventoryDataInfo = New List(Of MSG_InventoryData.clsInventoryDataList.clsInventoryDataInfo)
      Dim InventoryInfo As New MSG_InventoryData.clsInventoryDataList.clsInventoryDataInfo
      '表身
      Dim InventoryDetailList As New MSG_InventoryData.clsInventoryDataList.clsInventoryDataInfo.clsInventoryDetailDataList
      InventoryDetailList.InventoryDetailDataInfo = New List(Of MSG_InventoryData.clsInventoryDataList.clsInventoryDataInfo.clsInventoryDetailDataList.clsInventoryDetailDataInfo)

      For Each objStocktakingInfo In Receive_Msg.Body.StocktakingList.StocktakingInfo

        Stocktaking_ID = objStocktakingInfo.STOCKTAKING_ID
        Dim dicStocktaking As New Dictionary(Of String, clsHSTOCKTAKING)
        gMain.objHandling.O_GetDB_dicHStocktakingByStocktakingID(Stocktaking_ID, dicStocktaking)
        Dim objStocktaking As clsHSTOCKTAKING = Nothing
        If dicStocktaking.Any = False Then
          SendMessageToLog("查無相關單據，STOCKTAKING_ID：" & Stocktaking_ID, eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        Else
          objStocktaking = dicStocktaking.First.Value
        End If
        InventoryInfo.Inventory_Id = objStocktakingInfo.STOCKTAKING_ID
        '用盤點單號取的STOCKTAKING_DTL
        Dim dicSTOCKTAKING_DTL As New Dictionary(Of String, clsHSTOCKTAKINGDTL)
        gMain.objHandling.O_GetDB_dicHStocktaking_DTLByStocktakingID(objStocktakingInfo.STOCKTAKING_ID, dicSTOCKTAKING_DTL)
        If dicSTOCKTAKING_DTL.Any = False Then
          SendMessageToLog("查無相關單據，STOCKTAKING_ID：" & Stocktaking_ID, eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        End If
        If objStocktaking.STATUS = 4 Then
          For Each objSTOCKTAKING_DTL In dicSTOCKTAKING_DTL.Values
            Dim InventoryDetailInfo As New MSG_InventoryData.clsInventoryDataList.clsInventoryDataInfo.clsInventoryDetailDataList.clsInventoryDetailDataInfo
            Dim QTY As Decimal = 0

            InventoryDetailInfo.SerialId = objSTOCKTAKING_DTL.STOCKTAKING_SERIAL_NO
            Dim lstSKU = objSTOCKTAKING_DTL.SKU_NO.Split("_")
            InventoryDetailInfo.SKU = lstSKU(0)
            InventoryDetailInfo.LotId = objSTOCKTAKING_DTL.LOT_NO
            InventoryDetailInfo.InventoryQty = QTY
            'InventoryDetailInfo.LotId = objStocktakingInfo.LOT_NO
            'InventoryDetailInfo.CarrierID = objStocktakingInfo.CARRIER_ID
            'InventoryDetailInfo.InventoryQty = objStocktakingInfo.REPORT_QTY
            InventoryDetailList.InventoryDetailDataInfo.Add(InventoryDetailInfo)
          Next
        Else
          Dim dicStocktaking_Carrier As New Dictionary(Of String, clsHSTOCKTAKINGCARRIER)
          gMain.objHandling.O_GetDB_dicStocktaking_CarrierByStocktakingID(Stocktaking_ID, dicStocktaking_Carrier)
          If dicStocktaking_Carrier.Any = False Then
            SendMessageToLog("查無相關盤點棧板,STOCKTAKING_ID:" & Stocktaking_ID, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          For Each objSTOCKTAKING_DTL In dicSTOCKTAKING_DTL.Values
            Dim InventoryDetailInfo As New MSG_InventoryData.clsInventoryDataList.clsInventoryDataInfo.clsInventoryDetailDataList.clsInventoryDetailDataInfo
            Dim QTY As Decimal = 0
            For Each objStocktakingCarrier In dicStocktaking_Carrier.Values
              If objStocktakingCarrier.STOCKTAKING_SERIAL_NO = objSTOCKTAKING_DTL.STOCKTAKING_SERIAL_NO AndAlso objStocktakingCarrier.SKU_NO = objSTOCKTAKING_DTL.SKU_NO Then
                If objStocktakingCarrier.STOCKTAKING_STATUS = 0 Then
                  QTY += objStocktakingCarrier.QTY
                Else
                  QTY += objStocktakingCarrier.REPORT_QTY
                End If
              End If
            Next
            InventoryDetailInfo.SerialId = objSTOCKTAKING_DTL.STOCKTAKING_SERIAL_NO
            Dim lstSKU = objSTOCKTAKING_DTL.SKU_NO.Split("_")
            InventoryDetailInfo.SKU = lstSKU(0)
            InventoryDetailInfo.LotId = objSTOCKTAKING_DTL.LOT_NO
            InventoryDetailInfo.InventoryQty = QTY
            'InventoryDetailInfo.LotId = objStocktakingInfo.LOT_NO
            'InventoryDetailInfo.CarrierID = objStocktakingInfo.CARRIER_ID
            'InventoryDetailInfo.InventoryQty = objStocktakingInfo.REPORT_QTY
            InventoryDetailList.InventoryDetailDataInfo.Add(InventoryDetailInfo)
          Next
        End If


      Next
      Stocktaking_ID = InventoryInfo.Inventory_ID
      InventoryInfo.InventoryDetailDataList = InventoryDetailList
      InventoryDataList.InventoryDataInfo.Add(InventoryInfo)
      StocktakingReport.InventoryDataList = InventoryDataList


      Return True

    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      'For Each objtrans_control In dicAddTrans_Control.Values
      '  If objtrans_control.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "get insert trans_control sql failed"
      '    Return False
      '  End If
      'Next
      'For Each objRcvm In dicAddW2E_Rcvm.Values
      '  If objRcvm.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_Rcvm SQL Failed"
      '    Return False
      '  End If
      'Next
      'For Each objRcvd In dicAddW2E_Rcvd.Values
      '  If objRcvd.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_Rcvd SQL Failed"
      '    Return False
      '  End If
      'Next

      'For Each objW2E_RetnResult_Head In ret_dicAddW2E_RetnResult_Head.Values
      '  If objW2E_RetnResult_Head.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_RetnResult_Head SQL Failed"
      '    Return False
      '  End If
      'Next
      'For Each objW2E_RetnResult_Item In ret_dicAddW2E_RetnResult_item.Values
      '  If objW2E_RetnResult_Item.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_RetnResult_Item SQL Failed"
      '    Return False
      '  End If
      'Next

      'For Each obj In ret_dicAddW2E_Damage_Head.Values
      '  If obj.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_Damage_Head SQL Failed"
      '    Return False
      '  End If
      'Next
      'For Each obj In ret_dicAddW2E_Damage_Item.Values
      '  If obj.O_Add_Insert_SQLString(lstSql) = False Then
      '    Result_Message = "Get Insert W2E_Damage_Item SQL Failed"
      '    Return False
      '  End If
      'Next

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行新增的Carrier和Carrier_Status的SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef Result_Message As String,
                                         ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL   更新杏一的資料庫
      'If MedFirst_DBManagement.BatchUpdate(lstSql) = False Then
      '  '更新DB失敗則回傳False
      '  Result_Message = "eHOST 更新资料库失败"
      '  Return False
      'End If

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Getcls_browByItem_Common3(ByVal ITEM_COMMON3 As String) As String

    Select Case ITEM_COMMON3
      Case "21"
        Return "1"  '遺失
      Case "22"
        Return "2"  '破損
      Case "23"
        Return "3"  '浸水
    End Select
    Return "0"  '
  End Function
  Private Function CombinationBatchID(ByVal nowtime As String, ByVal BatchNum As String)
    Try
      Dim ret As String = ""
      ret = "R" + nowtime + BatchNum
      Return ret
    Catch ex As Exception
      Return ""
    End Try
  End Function
End Module



