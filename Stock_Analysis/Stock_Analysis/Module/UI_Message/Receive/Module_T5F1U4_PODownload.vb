'20180629
'V1.0.0
'Jerry

'提單(強制或不強制)

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T5F1U4_PODownload
  Public Function O_T5F1U4_PODownload(ByVal Receive_Msg As MSG_T5F1U4_PODownload,
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
  Private Function Check_Data(ByVal Receive_Msg As MSG_T5F1U4_PODownload,
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
  Private Function Process_Data(ByVal Receive_Msg As MSG_T5F1U4_PODownload,
                              ByRef ret_strResultMsg As String, ByRef ret_Wait_UUID As String) As Boolean
    Try
      '先進行資料邏輯檢查
      For Each objPOInfo In Receive_Msg.Body.POList.POInfo
        '資料檢查
        Dim PO_ID As String = objPOInfo.PO_ID
        Dim H_PO_ORDER_TYPE As String = objPOInfo.H_PO_ORDER_TYPE
        Dim FORCED_UPDATE As String = objPOInfo.FORCED_UPDATE
        Dim FORCED_FLAG As Boolean = False
        Dim dicPO_ID As New Dictionary(Of String, String)
        Dim tmp_PO_ID = PO_ID
        Dim ERP_ORDER_TYPE = objPOInfo.ERP_ORDER_TYPE

        If ERP_ORDER_TYPE = "VC" Then
          '環鴻客制
          'If Mod_WCFHost.ASRS_getSingleVC(PO_ID, objPOInfo.COMMON01, ret_strResultMsg) = True Then
          '  Return True
          'End If
        Else
          SendMessageToLog("Start Process Data T5F1U4_PODownload", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)     'Vito_12b30
          SyncLock gMain.objHandling.objCT_PO_DTLLock
            SendMessageToLog("T5F1U4_PODownload objCT_PO_DTLLock Locked", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)  'Vito_12b30
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
          SendMessageToLog("Start Process Data T5F1U4_PODownload", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)     'Vito_12b30
          Dim Str_USER = ""

        End If
      Next
      Return False
    Catch ex As Exception
      ret_strResultMsg = ex.InnerException.Message
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function



End Module
