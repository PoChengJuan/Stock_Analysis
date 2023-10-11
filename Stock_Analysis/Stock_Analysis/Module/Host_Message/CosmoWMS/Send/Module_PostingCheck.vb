'20180718
'V1.0.0
'Jerry
'Cosmo_WMS 回覆的過帳結果


Module Module_PostingCheck

  '過帳結果的回覆
  Public Function O_PostingCheck(ByRef Receive_Msg As List(Of eCA_TransactionMessage.MSG_PostingCheck), ByRef Result_Message As String) As Boolean
    Try


      ''要新增的Carrier資料
      'Dim dicAdd_Carrier As New Dictionary(Of String, eCA_WMSObject.clsCarrier)

      ''儲存要更新的SQL，進行一次性更新
      'Dim lstSql As New List(Of String)

      ''先進行資料邏輯檢查
      'If Check_CarrierInstallData(Receive_Msg, Result_Message) = False Then
      '  Return False
      'End If
      ''取得要新增的Carrier和Carrier_Status的資料
      'If Get_CarrierInstallData(Receive_Msg, Result_Message, dicAdd_Carrier) = False Then
      '  Return False
      'End If
      ''取得要更新到DB的SQL(新增Carrier、Carrier_Status的Insert SQL)
      'If Get_Insert_Carrier_SQL(Result_Message, dicAdd_Carrier, lstSql) = False Then
      '  Return False
      'End If
      ''執行資料更新
      'If Execute_PO_DataUpdate(Result_Message, dicAdd_Carrier, lstSql) = False Then
      '  Return False
      'End If
      'Return True


      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module
