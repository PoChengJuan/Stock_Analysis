'20190423
'V1.0.0
'Jerry

'盘点单提单

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T11F3U1_StocktakingDownload
  Public Function O_T11F3U1_StocktakingDownload(ByVal Receive_Msg As MSG_T11F3U1_StocktakingDownload,
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_T11F3U1_StocktakingDownload,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      Dim FromDate = Receive_Msg.Body.DataInfo.FromDate
      Dim ToDate = Receive_Msg.Body.DataInfo.ToDate
      '檢查 FromDate 是否為空
      If FromDate = "" Then
        ret_strResultMsg = "FromDate is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If ToDate = "" Then
        ret_strResultMsg = "ToDate is empty"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '資料處理
  Private Function Process_Data(ByVal Receive_Msg As MSG_T11F3U1_StocktakingDownload,
                              ByRef ret_strResultMsg As String, ByRef ret_Wait_UUID As String) As Boolean
    Try
      '先進行資料邏輯檢查
      Dim FromDate = Receive_Msg.Body.DataInfo.FromDate
      Dim ToDate = Receive_Msg.Body.DataInfo.ToDate

      'If Mod_WCFHost.ASRS_stockCheck(FromDate, ToDate, ret_strResultMsg) = False Then
      '  Return False
      'End If


      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  'Private Function Update_PO_LINE_POL1(ByVal PO_ID As String) As Boolean
  '    Try
  '        Dim lstSQL As New List(Of String)
  '        Dim dicPO_LINE As New Dictionary(Of String, clsPO_LINE)
  '        If gMain.objHandling.O_Get_dicPOLineByPOID(PO_ID, dicPO_LINE) = False Then
  '            Return False
  '        End If
  '        For Each objPO_LINE In dicPO_LINE.Values
  '            objPO_LINE.Update_Data_H_POL1()
  '            If objPO_LINE.O_Add_Update_SQLString(lstSQL) = False Then
  '                SendMessageToLog("GET Update PO_LINE POH1 SQL Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '                Return False
  '            End If
  '        Next
  '        If Common_DBManagement.BatchUpdate_DynamicConnection(lstSQL) = False Then
  '            SendMessageToLog("Update PO_LINE POH1 to DB Failed", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '            Return False
  '        End If



  '        Return True
  '    Catch ex As Exception
  '        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '        Return False
  '    End Try
  'End Function


End Module
