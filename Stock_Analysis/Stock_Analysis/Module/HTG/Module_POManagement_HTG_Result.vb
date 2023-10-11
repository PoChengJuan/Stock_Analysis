'20230112
'V1.0.0
'Bom
'處理PO單據執行成功後，更新ERP中介檔狀態

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_POManagement_HTG_Result
  Public Function O_CheckMessageResult(ByVal obj As MSG_T5F1U1_PO_Management,
                                       ByVal blnResult As Boolean,
                                       ByRef ret_strResultMsg As String) As Boolean

    Try
      '要變更的資料
      '儲存要更新的SQL，進行一次性更新
      Dim lstSql As New List(Of String)
      Dim dicUpdateEPSXB As New Dictionary(Of String, clsEPSXB)
      Dim dicUpdateMOCXD As New Dictionary(Of String, clsMOCXD)
      Dim dicUpdateINVXF As New Dictionary(Of String, clsINVXF)
      '檢查資料
      If Check_Data(obj, ret_strResultMsg) = False Then
        Return False
      End If
      '進行資料調整
      If Get_UpdateData(obj, blnResult, dicUpdateEPSXB, dicUpdateMOCXD, dicUpdateINVXF, ret_strResultMsg) = False Then
        'SendPurchaserData(enuRtnCode.Fail, PO_TYPE, PO_ID, ret_strResultMsg)
        Return False
      End If
      '取得SQL
      If Get_SQL(ret_strResultMsg, dicUpdateEPSXB, dicUpdateMOCXD, dicUpdateINVXF, lstSql) = False Then
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

  Private Function Check_Data(ByRef obj As MSG_T5F1U1_PO_Management,
                              ByRef ret_strResultMsg As String) As Boolean
    Try


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Get_UpdateData(ByRef obj As MSG_T5F1U1_PO_Management,
                                  ByVal blnResult As Boolean,
                                  ByRef ret_dicUpdateEPSXB As Dictionary(Of String, clsEPSXB),
                                  ByRef ret_dicUpdateMOCXD As Dictionary(Of String, clsMOCXD),
                                  ByRef ret_dicUpdateINVXF As Dictionary(Of String, clsINVXF),
                                  ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()

      Dim PO_TYPE2 = obj.Body.POInfo.PO_TYPE2
      Dim PO_KEY1 = obj.Body.POInfo.PO_KEY1
      Dim PO_KEY2 = obj.Body.POInfo.PO_KEY2

      Dim ACTION = obj.Body.Action
      Dim STATUS As String = ""

      If blnResult = False Then
#Region "失敗"
        For Each PO_DTL In obj.Body.POInfo.PODetailList.PODetailInfo
          Dim PO_KEY3 = PO_DTL.PO_SERIAL_NO

          Select Case PO_TYPE2
            Case enuPOType_2.Material_Out
              Dim tmp_dicMOCXD As New Dictionary(Of String, clsMOCXD)
              If gMain.objHandling.O_GetDB_dicMOCXDByKEY(PO_KEY1, PO_KEY2, PO_KEY3, tmp_dicMOCXD) = False Then
                ret_strResultMsg = $"Get MOCXB ByKEY FAIL KEY1:{PO_KEY1},KEY1:{PO_KEY2}, KEY1:{PO_KEY3}"
                Return False
              End If

              For Each objUpdateMOCXD In tmp_dicMOCXD.Values
                objUpdateMOCXD.XD011 = "3"

                If ret_dicUpdateMOCXD.ContainsKey(objUpdateMOCXD.gid) = False Then
                  ret_dicUpdateMOCXD.Add(objUpdateMOCXD.gid, objUpdateMOCXD.Clone)
                End If
              Next
            Case enuPOType_2.Sell_Out
              Dim tmp_dicEPSCB As New Dictionary(Of String, clsEPSXB)
              If gMain.objHandling.O_GetDB_dicEPSXBByKEY(PO_KEY1, PO_KEY2, PO_KEY3, tmp_dicEPSCB) = False Then
                ret_strResultMsg = $"Get EPSXB ByKEY FAIL KEY1:{PO_KEY1},KEY1:{PO_KEY2}, KEY1:{PO_KEY3}"
                Return False
              End If

              For Each objUpdateEPSXB In tmp_dicEPSCB.Values
                objUpdateEPSXB.XB010 = "3"
                objUpdateEPSXB.XB015 = Now_Time

                If ret_dicUpdateEPSXB.ContainsKey(objUpdateEPSXB.gid) = False Then
                  ret_dicUpdateEPSXB.Add(objUpdateEPSXB.gid, objUpdateEPSXB.Clone)
                End If
              Next
            Case enuPOType_2.transfer_in, enuPOType_2.normal_in, enuPOType_2.transfer_out, enuPOType_2.normal_out
              Dim tmp_dicINVXF As New Dictionary(Of String, clsINVXF)
              If gMain.objHandling.O_GetDB_dicINVXFByKEY(PO_KEY1, PO_KEY2, PO_KEY3, tmp_dicINVXF) = False Then
                ret_strResultMsg = $"Get INVXF ByKEY FAIL KEY1:{PO_KEY1},KEY1:{PO_KEY2}, KEY1:{PO_KEY3}"
                Return False
              End If

              For Each objUpdateINVXF In tmp_dicINVXF.Values
                objUpdateINVXF.XF009 = "3"

                If ret_dicUpdateINVXF.ContainsKey(objUpdateINVXF.gid) = False Then
                  ret_dicUpdateINVXF.Add(objUpdateINVXF.gid, objUpdateINVXF)
                End If
              Next
            Case Else

          End Select
        Next
#End Region
        Return True
      End If



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

        Select Case PO_TYPE2
          Case enuPOType_2.Material_Out
            Dim tmp_dicMOCXD As New Dictionary(Of String, clsMOCXD)
            If gMain.objHandling.O_GetDB_dicMOCXDByKEY(PO_KEY1, PO_KEY2, PO_KEY3, tmp_dicMOCXD) = False Then
              ret_strResultMsg = $"Get MOCXB ByKEY FAIL KEY1:{PO_KEY1},KEY1:{PO_KEY2}, KEY1:{PO_KEY3}"
              Return False
            End If

            For Each objUpdateMOCXD In tmp_dicMOCXD.Values
              objUpdateMOCXD.XD011 = STATUS

              If ret_dicUpdateMOCXD.ContainsKey(objUpdateMOCXD.gid) = False Then
                ret_dicUpdateMOCXD.Add(objUpdateMOCXD.gid, objUpdateMOCXD.Clone)
              End If
            Next
          Case enuPOType_2.Sell_Out
            Dim tmp_dicEPSCB As New Dictionary(Of String, clsEPSXB)
            If gMain.objHandling.O_GetDB_dicEPSXBByKEY(PO_KEY1, PO_KEY2, PO_KEY3, tmp_dicEPSCB) = False Then
              ret_strResultMsg = $"Get EPSXB ByKEY FAIL KEY1:{PO_KEY1},KEY1:{PO_KEY2}, KEY1:{PO_KEY3}"
              Return False
            End If

            For Each objUpdateEPSXB In tmp_dicEPSCB.Values
              objUpdateEPSXB.XB010 = STATUS
              objUpdateEPSXB.XB015 = Now_Time

              If ret_dicUpdateEPSXB.ContainsKey(objUpdateEPSXB.gid) = False Then
                ret_dicUpdateEPSXB.Add(objUpdateEPSXB.gid, objUpdateEPSXB.Clone)
              End If
            Next
          Case enuPOType_2.transfer_in, enuPOType_2.normal_in, enuPOType_2.transfer_out, enuPOType_2.normal_out
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
          Case Else

        End Select
      Next

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef ret_dicUpdateEPSXB As Dictionary(Of String, clsEPSXB),
                           ByRef ret_dicUpdateMOCXD As Dictionary(Of String, clsMOCXD),
                           ByRef ret_dicUpdateINVXF As Dictionary(Of String, clsINVXF),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得Host_Command的SQL
      For Each obj In ret_dicUpdateEPSXB.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get UPDATE EPSXB SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj In ret_dicUpdateMOCXD.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get UPDATE MOCXD SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
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
