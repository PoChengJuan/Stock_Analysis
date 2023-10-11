
Imports eCA_TransactionMessage

Module Module_ProcessWMSMessage
  '處理所有傳入的GUICommand
  Public Function O_ProcessWMSCommand(ByVal strFunction_ID As String,
                                      ByVal strXmlMessage As String,
                                      ByRef ret_strResultMsg As String,
                                      ByRef Wait_UUID As String) As Boolean
    Try
      Dim blnProcessResult As Boolean = False
      SendMessageToLog("Process WMS Command Start... Function_ID = " & strFunction_ID & " XML = " & strXmlMessage, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      SendMessageToLog("Message XML: " & strXmlMessage, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      Select Case strFunction_ID
        Case enuWMSMessageFunctionID.T5F1S11_POClose.ToString
          If O_Process_T11F1S1_POClose(strXmlMessage, ret_strResultMsg) = True Then
            blnProcessResult = True
          End If
        Case enuWMSMessageFunctionID.T10F2S1_StocktakingReport.ToString
          If O_Process_T10F2S1_StocktakingReport(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case enuWMSMessageFunctionID.T5F1S1_WOClose.ToString
          If O_Process_T5F1S1_WOClose(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case enuWMSMessageFunctionID.T5F1S13_TransationPOExecution.ToString
          If O_Process_T5F1S13_TransationPOExecution(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case enuWMSMessageFunctionID.T5F1U90_WOExcuting.ToString
          If O_ProcessResult_T5F1U90_WOExcuting(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case enuWMSMessageFunctionID.T5F1S31_CarrierProduceReport.ToString
          'WMS那邊棧板出庫後會自動揀貨，會發這個訊息過來要上報，但輝庭不用上報揀貨結果，直接給他TRUE就好，避免動到WMS的CODE
          blnProcessResult = True
        Case Else
          blnProcessResult = False
          ret_strResultMsg = "Not Defines Function_ID, Function_ID=" & strFunction_ID
      End Select
      SendMessageToLog("Process WMS Command Finish... Function_ID = " & strFunction_ID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Return blnProcessResult
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''' <summary>
  ''' POClose
  ''' </summary>
  ''' <param name="strXmlMessage"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <returns></returns>
  Public Function O_Process_T10F2S1_StocktakingReport(ByVal strXmlMessage As String,
                                                     ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T10F2S1_StocktakingReport = Nothing
      If ParseXmlString.ParseMessage_T10F2S1_StocktakingReport(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T10F2S1_StocktakingReport.O_T10F2S1_StocktakingReport(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T10F2S1_StocktakingReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T10F2S1_StocktakingReport Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T11F1S1_POClose(ByVal strXmlMessage As String,
                                                       ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T11F1S1_POClose = Nothing
      If ParseXmlString.ParseMessage_T11F1S1_POClose(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T11F1S1_POClose.O_T11F1S1_POClose(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T11F1S1_POClose Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T11F1S1_POClose Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T5F1S1_WOClose(ByVal strXmlMessage As String,
                                                       ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T5F1S1_WOClose = Nothing
      If ParseXmlString.ParseMessage_T5F1S1_WOClose(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T5F1S1_WOClose.O_T5F1S1_WOClose(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T5F1S1_WOClose Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T5F1S1_WOClose Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function O_Process_T5F1S13_TransationPOExecution(ByVal strXmlMessage As String,
                                                       ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim obj As MSG_T5F1S13_TransationPOExecution = Nothing
      If ParseXmlString.ParseMessage_T5F1S13_TransationPOExecution(strXmlMessage, obj, ret_strResultMsg) = True Then
        '執行相對應的事件
        If Module_T5F1S13_TransationPOExecution.O_T5F1S13_TransationPOExecution(obj, ret_strResultMsg) = True Then
          Return True
        Else
          gMain.SendMessageToLog("O_Process_T5F1S13_TransationPOExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Else
        gMain.SendMessageToLog("ParseMessage_T5F1S13_TransationPOExecution Failed, ResultMsg=" & ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      gMain.SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module