'20200106
'V1.0.0
'Vito
'Vito_20106
'WMS向HostHandler進行收料資訊的回報

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T5F1S4_PODeliveryReport
  Public Function O_T5F1S4_PODeliveryReport(ByVal Receive_Msg As MSG_T5F1S4_PODeliveryReport,
                                          ByRef ret_strResultMsg As String) As Boolean
    Try

      Dim lstSql As New List(Of String)
      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If

      '進行資料處理
      If Process_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If

      If Get_SQL(ret_strResultMsg, lstSql) = False Then
        Return False
      End If

      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_T5F1S4_PODeliveryReport,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '資料處理
  Private Function Process_Data(ByVal Receive_Msg As MSG_T5F1S4_PODeliveryReport,
                                ByRef ret_strResultMsg As String) As Boolean
    Try
      '資料處理

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
      '取得要新增的SQL語句

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

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module



