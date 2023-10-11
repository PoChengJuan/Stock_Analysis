Imports eCA_HostObject
Imports eCA_TransactionMessage

Public Module Module_T5F1U90_WOExcuting
  Public Function O_T5F1U90_WOExcuting(ByVal Receive_Msg As MSG_T5F1S90_WOExcuting,
                                       ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim lstSql As New List(Of String)
      Dim lstSql_ERP As New List(Of String)
      Dim dicUpdate_EPSXB As New Dictionary(Of String, clsEPSXB)
      Dim dicUpdate_MOCXD As New Dictionary(Of String, clsMOCXD)
      Dim dicUpdate_INVXF As New Dictionary(Of String, clsINVXF)

      '先進行資料邏輯檢查
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If

      '進行資料處理
      If Process_Data(Receive_Msg, ret_strResultMsg, dicUpdate_EPSXB, dicUpdate_MOCXD, dicUpdate_INVXF) = False Then
        Return False
      End If

      If Get_SQL(ret_strResultMsg, lstSql, lstSql_ERP, dicUpdate_EPSXB, dicUpdate_MOCXD, dicUpdate_INVXF) = False Then
        Return False
      End If

      If Execute_DataUpdate(ret_strResultMsg, lstSql, lstSql_ERP) = False Then
        Return False
      End If

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function Check_Data(ByVal Receive_Msg As MSG_T5F1S90_WOExcuting,
                             ByRef ret_strResultMsg As String) As Boolean
    Try
      '先進行資料邏輯檢查
      For Each objPO In Receive_Msg.Body.POList.POInfo
        If objPO.PO_ID = "" Then
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

  Public Function Process_Data(ByVal Receive_Msg As MSG_T5F1S90_WOExcuting,
                               ByRef ret_strResultMsg As String,
                               ByRef dicUpdate_EPSXB As Dictionary(Of String, clsEPSXB),
                               ByRef dicUpdate_MOCXD As Dictionary(Of String, clsMOCXD),
                               ByRef dicUpdate_INVXF As Dictionary(Of String, clsINVXF)) As Boolean
    Try
      '準備資料
      Dim Now_Time As String = ModuleHelpFunc.GetNewTime_DBFormat()

      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        Dim dicPO As New Dictionary(Of String, clsPO)

        If gMain.objHandling.O_Get_dicPOByPO_ID(objPOInfo.PO_ID, dicPO) = False Then
          ret_strResultMsg = $"PO_ID:{objPOInfo.PO_ID}，取得PO失敗"

          Return False
        End If

        Dim dicPOStr As New Dictionary(Of String, String)
        dicPOStr.Add(objPOInfo.PO_ID, objPOInfo.PO_ID)
        Dim dicPODTL As New Dictionary(Of String, clsPO_DTL)
        If gMain.objHandling.O_Get_dicPODTLBydicPO_ID(dicPOStr, dicPODTL) = False Then
          ret_strResultMsg = $"PO_ID:{objPOInfo.PO_ID}，取得PO_DTL失敗"

          Return False
        End If

        Dim objPO = dicPO.First.Value

        Dim PO_ID As String = objPO.PO_ID
        Dim PO_TYPE2 As Integer = CInt(objPO.PO_Type2)
        Dim ERP_PO_TYPE As String = String.Empty
        Dim ERP_PO_ID As String = String.Empty

        Dim strSplit() = PO_ID.Split("_")
        If strSplit.Length = 2 Then
          ERP_PO_TYPE = strSplit(0)
          ERP_PO_ID = strSplit(1)
        Else
          ERP_PO_ID = PO_ID
        End If

        Select Case PO_TYPE2
          Case enuPOType_2.Material_Out '領料出庫(領退單)
            Dim dicMOCXD As New Dictionary(Of String, clsMOCXD)

            For Each objDTL In dicPODTL.Values
              If gMain.objHandling.O_GetDB_dicMOCXDByXD001_XD002_XD013(objPO.PO_KEY1, objPO.PO_KEY2, objDTL.PO_SERIAL_NO, dicMOCXD) = False Then
                ret_strResultMsg = $"ERP單號 :{ERP_PO_ID}，無法取得 MOCXD 資料"

                Return False
              Else
                For Each objMOCXD In dicMOCXD.Values
                  Dim obj = objMOCXD.Clone

                  'XD011 是更新碼
                  obj.XD011 = "4"

                  If dicUpdate_MOCXD.ContainsKey(obj.gid) = False Then
                    dicUpdate_MOCXD.Add(obj.gid, obj)
                  End If
                Next
              End If
            Next
          Case enuPOType_2.Sell_Out '成品出庫(出通單)
            Dim dicEPSXB As New Dictionary(Of String, clsEPSXB)

            For Each objDTL In dicPODTL.Values

              If gMain.objHandling.O_GetDB_dicEPSXBByXB001_XB002_XB003(objPO.PO_KEY1, objPO.PO_KEY2, objDTL.PO_SERIAL_NO, dicEPSXB) = False Then
                ret_strResultMsg = $"ERP單號 :{ERP_PO_ID}，無法取得 EPSXB 資料"

                Return False
              Else
                For Each objEPSXB In dicEPSXB.Values
                  Dim obj = objEPSXB.Clone

                  'XD010 是更新碼
                  obj.XB010 = "4"
                  obj.XB015 = Now_Time
                  If dicUpdate_EPSXB.ContainsKey(obj.gid) = False Then
                    dicUpdate_EPSXB.Add(obj.gid, obj)
                  End If
                Next
              End If
            Next
          Case enuPOType_2.transfer_in, enuPOType_2.transfer_out, enuPOType_2.Change_Stock, enuPOType_2.Change_Out '轉播單
            Dim dicINVXF As New Dictionary(Of String, clsINVXF)

            If gMain.objHandling.O_GetDB_dicINVXFByXF001_XF002(objPO.PO_KEY1, objPO.PO_KEY2, dicINVXF) = False Then
              ret_strResultMsg = $"ERP單號 :{ERP_PO_ID}，無法取得 INVXF 資料"

              Return False
            Else
              For Each objINVXF In dicINVXF.Values
                Dim obj = objINVXF.Clone

                'XF009 是更新碼
                obj.XF009 = "4"

                If dicUpdate_INVXF.ContainsKey(obj.gid) = False Then
                  dicUpdate_INVXF.Add(obj.gid, obj)
                End If
              Next
            End If
        End Select
      Next

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef lstSql As List(Of String),
                           ByRef lstSql_ERP As List(Of String),
                           ByVal dicUpdate_EPSXB As Dictionary(Of String, clsEPSXB),
                           ByVal dicUpdate_MOCXD As Dictionary(Of String, clsMOCXD),
                           ByVal dicUpdate_INVXF As Dictionary(Of String, clsINVXF)) As Boolean
    Try
      '取得要新增的SQL語句
      For Each obj In dicUpdate_EPSXB.Values
        If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Update EPSXB Command SQL Failed"
          Return False
        End If
      Next

      For Each obj In dicUpdate_MOCXD.Values
        If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Update MOCXD Command SQL Failed"
          Return False
        End If
      Next

      For Each obj In dicUpdate_INVXF.Values
        If obj.O_Add_Update_SQLString(lstSql_ERP) = False Then
          Result_Message = "Get Update INVXF Command SQL Failed"
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

  '執行SQL語句
  Private Function Execute_DataUpdate(ByRef Result_Message As String,
                                      ByRef lstSql As List(Of String),
                                      ByRef lstSql_ERP As List(Of String)) As Boolean
    Try
      If lstSql_ERP.Any Then
        If ERP_DBManagement.BatchUpdate(lstSql_ERP) = False Then
          Result_Message = "WMS Update ERP DB Failed"

          Return False
        End If
      End If

      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module
