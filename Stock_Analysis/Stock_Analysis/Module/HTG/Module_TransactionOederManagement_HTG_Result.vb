'20230112
'V1.0.0
'Bom
'處理PO單據執行成功後，更新ERP中介檔狀態

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_TransactionOederManagement_HTG_Result
  Public Function O_CheckMessageResult(ByVal obj As MSG_T5F5U1_TransactionOederManagement,
                                       ByRef ret_strResultMsg As String) As Boolean

    Try
      '要變更的資料
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)
      Dim dicUpdateINVXF As New Dictionary(Of String, clsINVXF)

      '檢查資料
      If Check_Data(obj, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(obj, dicUpdateINVXF, ret_strResultMsg) = False Then
        'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, dicUpdateINVXF, lstSql) = False Then
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

  Private Function Check_Data(ByRef obj As MSG_T5F5U1_TransactionOederManagement,
                              ByRef ret_strResultMsg As String) As Boolean
    Try


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Get_UpdateData(ByRef obj As MSG_T5F5U1_TransactionOederManagement,
                                  ByRef ret_dicUpdateINVXF As Dictionary(Of String, clsINVXF),
                                  ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()

      Dim PO_TYPE2 = obj.Body.POInfo.PO_TYPE2
      Dim str_PO_ID As String() = obj.Body.POInfo.PO_ID.Split("_")
      Dim PO_KEY1 = str_PO_ID(0)
      Dim PO_KEY2 = str_PO_ID(1)

      Dim ACTION = obj.Body.Action
      Dim STATUS As String = ""

      Select Case ACTION
        Case enuAction.Create.ToString
          STATUS = "1"
        Case enuAction.Modify.ToString
          STATUS = "8"
        Case enuAction.Delete.ToString
          STATUS = "A"
      End Select

      For Each PO_DTL In obj.Body.POInfo.PODetailList.PODetailInfo
        Dim PO_KEY3 = PO_DTL.PO_SERIAL_NO
        Dim tmp_dicINVXF As New Dictionary(Of String, clsINVXF)

        If gMain.objHandling.O_GetDB_dicINVXFByKEY(PO_KEY1, PO_KEY2, PO_KEY3, tmp_dicINVXF) = False Then
          ret_strResultMsg = $"Get INVXF ByKEY FAIL KEY1:{PO_KEY1},KEY1:{PO_KEY2}, KEY1:{PO_KEY3}"
          Return False
        End If

        For Each objUpdateINVXF In tmp_dicINVXF.Values
          objUpdateINVXF.XF009 = STATUS

          If ret_dicUpdateINVXF.ContainsKey(objUpdateINVXF.gid) = False Then
            ret_dicUpdateINVXF.Add(objUpdateINVXF.gid, objUpdateINVXF)
          End If
        Next
      Next

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef ret_dicUpdateINVXF As Dictionary(Of String, clsINVXF),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得Host_Command的SQL
      For Each obj In ret_dicUpdateINVXF.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get UPDATE INVXF SQL Failed"
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
