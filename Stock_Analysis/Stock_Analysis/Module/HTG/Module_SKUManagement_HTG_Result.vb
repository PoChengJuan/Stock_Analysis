'20230112
'V1.0.0
'Bom
'處理PO單據執行成功後，更新ERP中介檔狀態

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_SKUManagement_HTG_Result
  Public Function O_CheckMessageResult(ByVal obj As MSG_T2F3U1_SKUManagement,
                                       ByRef ret_strResultMsg As String) As Boolean

    Try
      '要變更的資料
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)
      Dim dicUpdateINVXB As New Dictionary(Of String, clsINVXB)
      Return True
      '檢查資料
      If Check_Data(obj, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(obj, dicUpdateINVXB, ret_strResultMsg) = False Then
        'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, dicUpdateINVXB, lstSql) = False Then
        'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        Return False
      End If
      '執行SQL與更新物件
      If Execute_DataUpdate(ret_strResultMsg, lstSql) = False Then
        'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        Return False
      End If

      'SendPurchaserData(enuRtnCode.Sucess, PO_TYPE, PO_ID, ret_strResultMsg)
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Check_Data(ByRef obj As MSG_T2F3U1_SKUManagement,
                              ByRef ret_strResultMsg As String) As Boolean
    Try


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Get_UpdateData(ByRef obj As MSG_T2F3U1_SKUManagement,
                                  ByRef ret_dicUpdateINVXB As Dictionary(Of String, clsINVXB),
                                  ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()

      For Each SKU_Info In obj.Body.SKUList.SKUInfo
        Dim SKU_NO = SKU_Info.SKU_NO
        Dim tmp_dicINVXB As New Dictionary(Of String, clsINVXB)
        If gMain.objHandling.O_GetDB_dicINVXBByKEY(SKU_NO, tmp_dicINVXB) = False Then
          ret_strResultMsg = $"Get INVXB ByKEY FAIL XB001:{SKU_NO}"
          Return False
        End If

        Dim objINVXB = tmp_dicINVXB.First.Value

        If ret_dicUpdateINVXB.ContainsKey(objINVXB.XB001) = False Then
          ret_dicUpdateINVXB.Add(objINVXB.gid, objINVXB.Clone)
        End If
      Next

      For Each objINVXB In ret_dicUpdateINVXB.Values
        objINVXB.XB008 = "1"
        objINVXB.XB002 = objINVXB.XB002.Replace("'", "''")
      Next

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef ret_dicUpdateINVXB As Dictionary(Of String, clsINVXB),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得Host_Command的SQL
      For Each obj In ret_dicUpdateINVXB.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get UPDATE INVXB SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                      ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If ERP_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        ret_strResultMsg = "WMS Update ERP DB Failed"
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
