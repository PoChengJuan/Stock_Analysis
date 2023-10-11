Imports eCA_TransactionMessage
Imports eCA_HostObject

''' <summary>
''' 20180612
''' V1.0.0
''' Mark
''' 處理使用DB當成交握介面
''' </summary>
Module Module_DBHandling
  '處理使用DB和GUI交握的部份
  Public Sub O_thr_GUIDBHandling()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        If Count < 10 Then
          Count = Count + 1
        Else
          Count = 0
          If gMain.int_tGUIDBHandle > 99 Then
            gMain.int_tGUIDBHandle = 0
          Else
            gMain.int_tGUIDBHandle = gMain.int_tGUIDBHandle + 1
          End If
        End If
        I_ProcessFromGUIMessage()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("thrGUIDBHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub
  Private Function I_ProcessFromGUIMessage() As Boolean
    Try
      Dim dicFromGUICommand = GUI_T_CommandManagement.GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(enuSystemType.HostHandler)
      While dicFromGUICommand.Any = True
        Dim Function_ID As String = ""
        Dim XmlMessage As String = ""
        Dim Result As String = ""
        Dim ResultMessage As String = ""
        Dim Wait_UUID As String = ""
        Dim UUID As String = ""
        '用來暫存要處理的GUICommand
        Dim dicProcessFromGUICommand As New Dictionary(Of String, clsFromGUICommand)
        For Each objFromGUICommand As clsFromGUICommand In dicFromGUICommand.Values
          If UUID = "" Then
            UUID = objFromGUICommand.UUID
          End If
          '不相等表示是下一筆了，下一次再處理
          If UUID <> objFromGUICommand.UUID Then
            Exit For
          End If
          '儲存GUICommand的資訊
          If dicProcessFromGUICommand.ContainsKey(objFromGUICommand.gid) = False Then
            dicProcessFromGUICommand.Add(objFromGUICommand.gid, objFromGUICommand)
          Else
            SendMessageToLog(String.Format("GUICommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objFromGUICommand.UUID, objFromGUICommand.Function_ID, objFromGUICommand.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
          Function_ID = objFromGUICommand.Function_ID
          XmlMessage = XmlMessage & objFromGUICommand.Message
        Next
        '把取得的GUICommand送出去進行處理
        If O_ProcessCommand(Function_ID, XmlMessage, ResultMessage, enuHTTPContentType.XML, Wait_UUID) = True Then
          '執行成功
          '如果Wait_UUID不為空時才把Wait_UUID填入
          If Wait_UUID = "" Then
            Result = "0"
            'ResultMessage = ""
          End If
        Else
          '執行失敗
          Result = "1"
        End If
        '把GUICommand的執行結果寫入DB
        Dim lstSQL As New List(Of String)
        For Each objFromGUICommand As clsFromGUICommand In dicProcessFromGUICommand.Values
          objFromGUICommand.Result = Result
          objFromGUICommand.Result_Message = StrConv(ResultMessage, VbStrConv.TraditionalChinese, 2052) ' ResultMessage
          objFromGUICommand.Wait_UUID = Wait_UUID
          SendMessageToLog("Set GUI Command Data Report, UUID=" & objFromGUICommand.UUID & ", Function_ID=" & objFromGUICommand.Function_ID & ", SEQ=" & objFromGUICommand.SEQ & ", Result=" & objFromGUICommand.Result & ", Result_Message=" & objFromGUICommand.Result_Message & ", Wait_UUID=" & objFromGUICommand.Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          If objFromGUICommand.O_Add_Update_SQLString(lstSQL) = False Then
            SendMessageToLog("Get SQL Faile, UUID=" & objFromGUICommand.UUID & ", Function_ID=" & objFromGUICommand.Function_ID & ", SEQ=" & objFromGUICommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
          End If
        Next
        If Common_DBManagement.BatchUpdate(lstSQL) = False Then
          SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
        '刪除已經處理過的GUICommand
        For Each objFromGUICommand As clsFromGUICommand In dicProcessFromGUICommand.Values
          If dicFromGUICommand.ContainsKey(objFromGUICommand.gid) = True Then
            dicFromGUICommand.Remove(objFromGUICommand.gid)
          End If
        Next
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '處理使用DB和MCS交握的部份
  Public Sub O_thr_MCSDBHandling()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        If Count < 10 Then
          Count = Count + 1
        Else
          Count = 0
          If gMain.int_tMCSDBHandle > 99 Then
            gMain.int_tMCSDBHandle = 0
          Else
            gMain.int_tMCSDBHandle = gMain.int_tMCSDBHandle + 1
          End If
        End If
        I_ProcessFromMCSMessage()
        I_ProcessToMCSMessageResult()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("thrMCSDBHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub
  Private Function I_ProcessFromMCSMessage() As Boolean
    Try
      Dim dicFromMCSCommand = MCS_T_CommandManagement.GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(enuSystemType.HostHandler)
      If dicFromMCSCommand Is Nothing Then Return False
      While dicFromMCSCommand.Any = True
        Dim Function_ID As String = ""
        Dim XmlMessage As String = ""
        Dim Result As String = ""
        Dim ResultMessage As String = ""
        Dim Wait_UUID As String = ""
        Dim UUID As String = ""
        '用來暫存要處理的MCSCommand
        Dim dicProcessFromMCSCommand As New Dictionary(Of String, clsFromMCSCommand)
        SendMessageToLog(String.Format("Get Process MCS Command"), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        For Each objFromMCSCommand As clsFromMCSCommand In dicFromMCSCommand.Values
          If UUID = "" Then
            UUID = objFromMCSCommand.UUID
            Function_ID = objFromMCSCommand.Function_ID '因為WCS的UUID會填重覆，所以要再加上Function_ID的控管
          End If
          '不相等表示是下一筆了，下一次再處理
          If UUID <> objFromMCSCommand.UUID OrElse Function_ID <> objFromMCSCommand.Function_ID Then
            Exit For
          End If
          '儲存GUICommand的資訊
          If dicProcessFromMCSCommand.ContainsKey(objFromMCSCommand.gid) = False Then
            dicProcessFromMCSCommand.Add(objFromMCSCommand.gid, objFromMCSCommand)
          Else
            SendMessageToLog(String.Format("MCSCommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objFromMCSCommand.UUID, objFromMCSCommand.Function_ID, objFromMCSCommand.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
          Function_ID = objFromMCSCommand.Function_ID
          XmlMessage = XmlMessage & objFromMCSCommand.Message
        Next
        SendMessageToLog(String.Format("Start Process MCS Command"), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        '把取得的GUICommand送出去進行處理
        If O_ProcessCommand(Function_ID, XmlMessage, ResultMessage, enuHTTPContentType.XML, Wait_UUID) = True Then
          '執行成功
          Result = "0"
          ResultMessage = ""
        Else
          '執行失敗
          Result = "1"
        End If
        SendMessageToLog(String.Format("End Process MCS Command"), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        '把MCSCommand的執行結果寫入DB
        Dim lstSQL As New List(Of String)
        For Each objFromMCSCommand As clsFromMCSCommand In dicProcessFromMCSCommand.Values
          objFromMCSCommand.Result = Result
          objFromMCSCommand.Result_Message = ResultMessage
          If objFromMCSCommand.O_Add_Update_SQLString(lstSQL) = False Then
            SendMessageToLog("Get SQL Faile, UUID=" & objFromMCSCommand.UUID & ", Function_ID=" & objFromMCSCommand.Function_ID & ", SEQ=" & objFromMCSCommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
          End If
        Next
        If MCS_T_CommandManagement.BatchUpdate_DynamicConnection(lstSQL) = False Then
          SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
        '把處理過的MCSCommand加到歷史記錄，不刪除原Table的資料
        If ProcessFromMCSCommand_MoveToHist(dicProcessFromMCSCommand) = True Then

        End If
        SendMessageToLog(String.Format("Start Delete Memory MCS Command"), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        '清除已經處理過的MCSCommand
        For Each objFromMCSCommand As clsFromMCSCommand In dicProcessFromMCSCommand.Values
          If dicFromMCSCommand.ContainsKey(objFromMCSCommand.gid) = True Then
            dicFromMCSCommand.Remove(objFromMCSCommand.gid)
          End If
        Next
        SendMessageToLog(String.Format("End Delete Memory MCS Command"), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function




  '處理使用DB和WMS交握的部份
  Public Sub O_thr_WMSDBHandling()
    Const SleepTime As Integer = 200
    While True
      Try
        I_ProcessFromWMSCommand()
        'Vito_19b18  I_ProcessToWMSResult()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("thrWMSDBHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub
  '處理使用DB和其他系統交握的部份(回覆)
  Public Sub O_thr_ToOtherDBHandling_Result()
    Const SleepTime As Integer = 200
    While True
      Try
        I_ProcessToOtherSystemResult()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_ToOtherDBHandling_Result End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub
  Private Function I_ProcessFromWMSCommand() As Boolean
    Try
      Dim dicWMSCommand = WMS_T_HOST_CommandManagement.GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(enuSystemType.HostHandler)
      If dicWMSCommand IsNot Nothing Then

        While dicWMSCommand.Any = True
          Dim Function_ID As String = ""
          Dim XmlMessage As String = ""
          Dim Result As String = ""
          Dim ResultMessage As String = ""
          Dim Wait_UUID As String = ""
          Dim UUID As String = ""
          '用來暫存要處理的WMSCommand
          Dim dicProcessWMSCommand As New Dictionary(Of String, clsToHostCommand)
          For Each objWMSCommand As clsToHostCommand In dicWMSCommand.Values
            If UUID = "" Then
              UUID = objWMSCommand.UUID
            End If
            '不相等表示是下一筆了，下一次再處理
            If UUID <> objWMSCommand.UUID Then
              Exit For
            End If
            '儲存WMSCommand的資訊
            If dicProcessWMSCommand.ContainsKey(objWMSCommand.gid) = False Then
              dicProcessWMSCommand.Add(objWMSCommand.gid, objWMSCommand)
            Else
              SendMessageToLog(String.Format("WMSCommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objWMSCommand.UUID, objWMSCommand.Function_ID, objWMSCommand.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            End If
            Function_ID = objWMSCommand.Function_ID
            XmlMessage = XmlMessage & objWMSCommand.Message
          Next
          '把取得的WMSCommand送出去進行處理
          If O_ProcessWMSCommand(Function_ID, XmlMessage, ResultMessage, Wait_UUID) = True Then
            '執行成功
            If Wait_UUID = "" Then
              Result = "0"
              ResultMessage = ResultMessage
            End If
          Else
            '執行失敗
            Result = "1"
          End If
          '把WMSCommand的執行結果寫入DB
          Dim lstSQL As New List(Of String)
          For Each objWMSCommand As clsToHostCommand In dicProcessWMSCommand.Values
            objWMSCommand.Result = Result
            objWMSCommand.Result_Message = ResultMessage
            objWMSCommand.Wait_UUID = Wait_UUID
            If objWMSCommand.O_Add_Update_SQLString(lstSQL) = False Then
              SendMessageToLog("Get SQL Faile, UUID=" & objWMSCommand.UUID & ", Function_ID=" & objWMSCommand.Function_ID & ", SEQ=" & objWMSCommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
            End If
          Next
          If Common_DBManagement.BatchUpdate(lstSQL) = False Then
            SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
          End If
          ''刪除已經處理過的WMSCommand
          For Each objWMSCommand As clsToHostCommand In dicProcessWMSCommand.Values
            If objWMSCommand.Result <> "" Then
              If dicWMSCommand.ContainsKey(objWMSCommand.gid) = True Then
                dicWMSCommand.Remove(objWMSCommand.gid)
              End If
            End If
          Next
          '所有command 已處理完畢
          'Exit While

        End While
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function I_ProcessToOtherSystemResult() As Boolean
    Try
      Dim dicCommandResult = HOST_T_CommandManagement.GetCommandDictionaryBySendSystem_ResultIsNotNULL(enuSystemType.HostHandler)
      Dim ret_Msg As String = ""
      While dicCommandResult.Any = True
        Dim Function_ID As String = ""
        Dim XmlMessage As String = ""
        Dim Result As Boolean = True
        Dim ResultMessage As String = ""
        Dim UUID As String = ""
        '用來暫存要處理的HostHandlerCommand
        Dim dicProcessCommandResult As New Dictionary(Of String, clsFromHostCommand)
        For Each objCommandResult As clsFromHostCommand In dicCommandResult.Values
          If UUID = "" Then
            UUID = objCommandResult.UUID
          End If
          '不相等表示是下一筆了，下一次再處理
          If UUID <> objCommandResult.UUID Then
            Exit For
          End If
          '儲存HostHandlerCommand的資訊
          If dicProcessCommandResult.ContainsKey(objCommandResult.gid) = False Then
            dicProcessCommandResult.Add(objCommandResult.gid, objCommandResult)
          Else
            SendMessageToLog(String.Format("CommandResult exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objCommandResult.UUID, objCommandResult.Function_ID, objCommandResult.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
          Function_ID = objCommandResult.Function_ID
          XmlMessage = XmlMessage & objCommandResult.Message
          Result = IIf(objCommandResult.Result = "0", True, False)
          ResultMessage = objCommandResult.Result_Message
        Next
        If Result = True Then
          SendMessageToLog("Get Host Command Data Report, UUID=" & UUID & ", Function_ID=" & Function_ID & ", Result=" & Result & ", Result_Message=" & ResultMessage, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog("Get Host Command Data Report, UUID=" & UUID & ", Function_ID=" & Function_ID & ", Result=" & Result & ", Result_Message=" & ResultMessage, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        '把取得的HostHandlerCommand送出去進行處理
        If O_ProcessCommandResult(UUID, Function_ID, XmlMessage, ResultMessage, Result, ret_Msg) = True Then
          '  '執行成功

        End If
        '根據Function來決定是否需要進行事件交握，進行相對應的處理，目前的處理方式固定使用一種模式
        '目前使用的法式是用UUID查詢GUICommand相對應的Wait_UUID並把結果寫入GUICommand
        Select Case Function_ID
          Case enuHostCommandFunctionID.T5F3U23_POToWO.ToString
            Dim dicCommandSetResult = GUI_T_CommandManagement.GetCommandDictionaryByReceiveSystem_WaitUUID(enuSystemType.HostHandler, UUID)
            If dicCommandSetResult.Any = True Then
              '把GUICommand的執行結果寫入DB
              Dim lstSQL As New List(Of String)
              For Each objCommand As clsFromGUICommand In dicCommandSetResult.Values
                objCommand.Result = IIf(Result = True, "0", "1")
                objCommand.Result_Message = ResultMessage
                SendMessageToLog("Set WMS Command Data Report, UUID=" & objCommand.UUID & ", Function_ID=" & objCommand.Function_ID & ", SEQ=" & objCommand.SEQ & ", Result=" & objCommand.Result & ", Result_Message=" & objCommand.Result_Message & ", Wait_UUID=" & objCommand.Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                If objCommand.O_Add_Update_SQLString(lstSQL) = False Then
                  SendMessageToLog("Get SQL Faile, UUID=" & objCommand.UUID & ", Function_ID=" & objCommand.Function_ID & ", SEQ=" & objCommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
                End If
              Next
              If Common_DBManagement.BatchUpdate(lstSQL) = False Then
                SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
              End If
            End If
          Case enuHostCommandFunctionID.T6F5U1_ItemLabelManagement.ToString
            Dim dicCommandSetResult = GUI_T_CommandManagement.GetCommandDictionaryByReceiveSystem_WaitUUID(enuSystemType.HostHandler, UUID)
            'For Each objCommandResult In dicCommandResult.Values
            '  With objCommandResult
            '    .Wait_UUID = ""
            '  End With
            'Next
          Case enuHostCommandFunctionID.T6F5U2_ItemLabelPrint.ToString
            Dim dicCommandSetResult = GUI_T_CommandManagement.GetCommandDictionaryByReceiveSystem_WaitUUID(enuSystemType.HostHandler, UUID)
            For Each objCommandResult In dicCommandResult.Values
              With objCommandResult
                .Wait_UUID = ""
              End With
            Next
            'Case Else
            '  Dim dicCommandSetResult = GUI_T_CommandManagement.GetCommandDictionaryByReceiveSystem_WaitUUID(enuSystemType.HostHandler, UUID)
            '  If dicCommandSetResult.Any = True Then
            '    '把GUICommand的執行結果寫入DB
            '    Dim lstSQL As New List(Of String)
            '    For Each objCommand As clsFromGUICommand In dicCommandSetResult.Values
            '      objCommand.Result = IIf(Result = True, "0", "1")
            '      objCommand.Result_Message = ResultMessage
            '      SendMessageToLog("Set WMS Command Data Report, UUID=" & objCommand.UUID & ", Function_ID=" & objCommand.Function_ID & ", SEQ=" & objCommand.SEQ & ", Result=" & objCommand.Result & ", Result_Message=" & objCommand.Result_Message & ", Wait_UUID=" & objCommand.Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
            '      If objCommand.O_Add_Update_SQLString(lstSQL) = False Then
            '        SendMessageToLog("Get SQL Faile, UUID=" & objCommand.UUID & ", Function_ID=" & objCommand.Function_ID & ", SEQ=" & objCommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
            '      End If
            '    Next
            '    If Common_DBManagement.BatchUpdate(lstSQL) = False Then
            '      SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            '    End If
            '  End If
          Case enuHostCommandFunctionID.T5F1U1_POManagement.ToString
            Dim dicCommandSetResult = GUI_T_CommandManagement.GetCommandDictionaryByReceiveSystem_WaitUUID(enuSystemType.HostHandler, UUID)
            '把GUICommand的執行結果寫入DB
            Dim lstSQL As New List(Of String)
            For Each objCommand As clsFromGUICommand In dicCommandSetResult.Values
              objCommand.Result = IIf(Result = True, "0", "1")
              objCommand.Result_Message = ResultMessage
              SendMessageToLog("Set WMS Command Data Report, UUID=" & objCommand.UUID & ", Function_ID=" & objCommand.Function_ID & ", SEQ=" & objCommand.SEQ & ", Result=" & objCommand.Result & ", Result_Message=" & objCommand.Result_Message & ", Wait_UUID=" & objCommand.Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
              If objCommand.O_Add_Update_SQLString(lstSQL) = False Then
                SendMessageToLog("Get SQL Faile, UUID=" & objCommand.UUID & ", Function_ID=" & objCommand.Function_ID & ", SEQ=" & objCommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
              End If
            Next
            If Common_DBManagement.BatchUpdate(lstSQL) = False Then
              SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            End If
        End Select
        '把HostCommand写入历史
        ProcessHostCommandResult_Move_to_Hist(dicProcessCommandResult)
        '刪除已經處理過的HostHandlerCommand
        For Each objCommand As clsFromHostCommand In dicProcessCommandResult.Values
          If dicCommandResult.ContainsKey(objCommand.gid) = True Then
            dicCommandResult.Remove(objCommand.gid)
          End If
        Next
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '將結束的命令移入歷史
  Private Function ProcessHostCommandResult_Move_to_Hist(ByVal tmp_dicMCSCommand As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try
      Dim lstSQL As New List(Of String)
      Dim lstQueueSQL As New List(Of String)
      Dim Now_Time As String = GetNewTime_DBFormat()
      For Each objFromHostCommand As clsFromHostCommand In tmp_dicMCSCommand.Values
        With objFromHostCommand
          '將WMS_T_MCSCommand寫到History
          Dim objToMCSCommandHist As New clsFromHostCommandHist(.UUID, .Send_System, .Receive_System, .Function_ID, .SEQ, .User_ID, .Create_Time, .Message, .Result, .Result_Message, .Wait_UUID, Now_Time)
          objToMCSCommandHist.O_Add_Insert_SQLString(lstQueueSQL)
        End With
        objFromHostCommand.O_Add_Delete_SQLString(lstSQL)
      Next
      If WMS_T_HOST_CommandManagement.BatchUpdate(lstSQL) = False Then
        SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      End If
      Common_DBManagement.AddQueued_BatchUpdate(lstQueueSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Send_MessageToWMS_ByDB(ByVal strMessage As String, HeaderInfo As clsHeader, ByRef dicWMSCommand As Dictionary(Of String, clsFromHostCommand)) As Boolean
    Try
      'Dim dicWMSCommand As New Dictionary(Of String, clsHOST_T_WMS_Command)
      Dim DBMaxLength As Long = 3500
      If HeaderInfo.ClientInfo IsNot Nothing Then
        '如果strMessage超過DBMaxLength個字就要分成多個Message(如果只有一個Message則SEQ填0)
        If strMessage.Length < DBMaxLength Then
          Dim objWMSCommand As clsFromHostCommand = Nothing
          If HeaderInfo.ClientInfo IsNot Nothing Then
            objWMSCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.WMS, HeaderInfo.EventID, 0, HeaderInfo.ClientInfo.UserID, "", "", GetNewTime_DBFormat, strMessage, "", "", "")
          Else
            objWMSCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.WMS, HeaderInfo.EventID, 0, "", "", "", GetNewTime_DBFormat, strMessage, "", "", "")
          End If
          If dicWMSCommand.ContainsKey(objWMSCommand.gid) = False Then
            dicWMSCommand.Add(objWMSCommand.gid, objWMSCommand)
          End If
        Else
          Dim Seq As Long = 1
          Do
            Dim NewMessage As String = ""
            If strMessage.Length > DBMaxLength Then
              NewMessage = strMessage.Substring(0, 3500)
              strMessage = strMessage.Substring(3500)
            Else
              NewMessage = strMessage.Substring(0, strMessage.Length)
              strMessage = ""
            End If

            Dim objWMSCommand As clsFromHostCommand = Nothing
            If HeaderInfo.ClientInfo IsNot Nothing Then
              objWMSCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.WMS, HeaderInfo.EventID, Seq, HeaderInfo.ClientInfo.UserID, "", "", GetNewTime_DBFormat, NewMessage, "", "", "")
            Else
              objWMSCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.WMS, HeaderInfo.EventID, Seq, "", "", "", GetNewTime_DBFormat, NewMessage, "", "", "")
            End If
            If dicWMSCommand.ContainsKey(objWMSCommand.gid) = False Then
              dicWMSCommand.Add(objWMSCommand.gid, objWMSCommand)
            End If
            Seq = Seq + 1
          Loop While (strMessage.Length > 0)
        End If
      End If

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Send_MessageToOther_ByDB(ByVal strMessage As String,
                                             ByVal HeaderInfo As clsHeader,
                                             ByVal sendTo As enuSystemType) As Boolean
    Try
      Dim dicFromHostCommand As New Dictionary(Of String, clsFromHostCommand)
      Dim DBMaxLength As Long = 3500
      If HeaderInfo.ClientInfo IsNot Nothing Then
        '如果strMessage超過DBMaxLength個字就要分成多個Message(如果只有一個Message則SEQ填0)
        If strMessage.Length < DBMaxLength Then
          Dim objFromHostCommand As clsFromHostCommand = Nothing
          If HeaderInfo.ClientInfo IsNot Nothing Then
            objFromHostCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, 0, HeaderInfo.ClientInfo.UserID, "", "", GetNewTime_DBFormat, strMessage, "", "", "")
          Else
            objFromHostCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, 0, "", "", "", GetNewTime_DBFormat, strMessage, "", "", "")
          End If
          If dicFromHostCommand.ContainsKey(objFromHostCommand.gid) = False Then
            dicFromHostCommand.Add(objFromHostCommand.gid, objFromHostCommand)
          End If
        Else
          Dim Seq As Long = 1
          Do
            Dim NewMessage As String = ""
            If strMessage.Length > DBMaxLength Then
              NewMessage = strMessage.Substring(0, 3500)
              strMessage = strMessage.Substring(3500)
            Else
              NewMessage = strMessage.Substring(0, strMessage.Length)
              strMessage = ""
            End If
            Dim objFromHostCommand As clsFromHostCommand = Nothing
            If HeaderInfo.ClientInfo IsNot Nothing Then
              objFromHostCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, Seq, HeaderInfo.ClientInfo.UserID, "", "", GetNewTime_DBFormat, NewMessage, "", "", "")
            Else
              objFromHostCommand = New clsFromHostCommand(HeaderInfo.UUID, enuSystemType.HostHandler, sendTo, HeaderInfo.EventID, Seq, "", "", "", GetNewTime_DBFormat, NewMessage, "", "", "")
            End If
            If dicFromHostCommand.ContainsKey(objFromHostCommand.gid) = False Then
              dicFromHostCommand.Add(objFromHostCommand.gid, objFromHostCommand)
            End If
            Seq = Seq + 1
          Loop While (strMessage.Length > 0)
        End If
      End If

      '取得HostCommand的Insert SQL
      Dim lstSQL As New List(Of String)
      For Each objFromHostCommand As clsFromHostCommand In dicFromHostCommand.Values
        If objFromHostCommand.O_Add_Insert_SQLString(lstSQL) = False Then
          SendMessageToLog("Get Insert HOST_T_WMS_Command SQL Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      '寫入DB
      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        SendMessageToLog("eHOST 更新資料庫失敗", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '送出Messge給HostHandler，使用DB當Intrface
  Public Function O_Send_MessageToMCS_ByDB(ByVal strMessage As String,
                                          ByVal HeaderInfo As clsHeader) As Boolean
    Try
      Dim dicToMCSCommand As New Dictionary(Of String, clsToMCSCommand)
      Dim DBMaxLength As Long = 3500
      If HeaderInfo.ClientInfo IsNot Nothing Then
        '如果strMessage超過DBMaxLength個字就要分成多個Message(如果只有一個Message則SEQ填0)
        If strMessage.Length < DBMaxLength Then
          Dim objToMCSCommand As clsToMCSCommand = Nothing
          If HeaderInfo.ClientInfo IsNot Nothing Then
            objToMCSCommand = New clsToMCSCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.MCS, HeaderInfo.EventID, 0, HeaderInfo.ClientInfo.UserID, GetNewTime_DBFormat, strMessage, "", "", "")
          Else
            objToMCSCommand = New clsToMCSCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.MCS, HeaderInfo.EventID, 0, "", GetNewTime_DBFormat, strMessage, "", "", "")
          End If
          If dicToMCSCommand.ContainsKey(objToMCSCommand.gid) = False Then
            dicToMCSCommand.Add(objToMCSCommand.gid, objToMCSCommand)
          End If
        Else
          Dim Seq As Long = 1
          Do
            Dim NewMessage As String = ""
            If strMessage.Length > DBMaxLength Then
              NewMessage = strMessage.Substring(0, 3500)
              strMessage = strMessage.Substring(3500)
            Else
              NewMessage = strMessage.Substring(0, strMessage.Length)
              strMessage = ""
            End If
            Dim objToMCSCommand As clsToMCSCommand = Nothing
            If HeaderInfo.ClientInfo IsNot Nothing Then
              objToMCSCommand = New clsToMCSCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.MCS, HeaderInfo.EventID, Seq, HeaderInfo.ClientInfo.UserID, GetNewTime_DBFormat, NewMessage, "", "", "")
            Else
              objToMCSCommand = New clsToMCSCommand(HeaderInfo.UUID, enuSystemType.HostHandler, enuSystemType.MCS, HeaderInfo.EventID, Seq, "", GetNewTime_DBFormat, NewMessage, "", "", "")
            End If
            If dicToMCSCommand.ContainsKey(objToMCSCommand.gid) = False Then
              dicToMCSCommand.Add(objToMCSCommand.gid, objToMCSCommand)
            End If
            Seq = Seq + 1
          Loop While (strMessage.Length > 0)
        End If
      End If
      '取得HostCommand的Insert SQL
      Dim lstSQL As New List(Of String)
      For Each objToMCSCommand As clsToMCSCommand In dicToMCSCommand.Values
        objToMCSCommand.O_Add_Insert_SQLString(lstSQL)
      Next
      '寫入DB
      If WMS_T_MCS_CommandManagement.BatchUpdate_DynamicConnection(lstSQL) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '處理使用DB接收MCS的回覆
  Private Function I_ProcessToMCSMessageResult() As Boolean
    Try
      Dim dicToMCSCommandResult = WMS_T_MCS_CommandManagement.GetCommandDictionaryBySendSystem_ResultIsNotNULL(enuSystemType.HostHandler)
      While dicToMCSCommandResult.Any = True
        Dim Function_ID As String = ""
        Dim XmlMessage As String = ""
        Dim ReplyResult As Boolean = True
        Dim RejectReason As String = ""
        Dim str_ResultMsg As String = ""
        Dim UUID As String = ""
        '用來暫存要處理的ToMCSCommand
        Dim dicProcessToMCSCommandResult As New Dictionary(Of String, clsToMCSCommand)
        For Each objToMCSCommandResult As clsToMCSCommand In dicToMCSCommandResult.Values
          If UUID = "" Then
            UUID = objToMCSCommandResult.UUID
          End If
          '不相等表示是下一筆了，下一次再處理
          If UUID <> objToMCSCommandResult.UUID Then
            Exit For
          End If
          '儲存HostHandlerCommand的資訊
          If dicProcessToMCSCommandResult.ContainsKey(objToMCSCommandResult.gid) = False Then
            dicProcessToMCSCommandResult.Add(objToMCSCommandResult.gid, objToMCSCommandResult)
          Else
            SendMessageToLog(String.Format("WMS_T_MCSCommand Result exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objToMCSCommandResult.UUID, objToMCSCommandResult.Function_ID, objToMCSCommandResult.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
          Function_ID = objToMCSCommandResult.Function_ID
          XmlMessage = XmlMessage & objToMCSCommandResult.Message
          ReplyResult = IIf(objToMCSCommandResult.Result = "0", True, False)
          RejectReason = objToMCSCommandResult.Result_Message
        Next
        '記錄發生給MCS後回覆的給結果
        If ReplyResult = True Then
          SendMessageToLog("Get Send To MCS Command Data Report, UUID=" & UUID & ", Function_ID=" & Function_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog("Get Send To MCS Command Data Report, UUID=" & UUID & ", Function_ID=" & Function_ID & ", Result=" & ReplyResult & ", Result_Message=" & RejectReason, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        '把取得的ToMCSCommand送出去進行處理
        If O_ProcessCommandResult(UUID, Function_ID, XmlMessage, RejectReason, ReplyResult, str_ResultMsg) = True Then
          '執行成功

        End If
        ''根據Function來決定是否需要進行事件交握，進行相對應的處理，目前的處理方式固定使用一種模式
        ''目前使用的法式是用UUID查詢GUICommand相對應的Wait_UUID並把結果寫入GUICommand
        'Select Case Function_ID
        '  Case enuHostMessageFunctionID.T11F1S1_POClose.ToString, enuHostMessageFunctionID.T11F1S2_POSimulationClose.ToString
        '    Dim dicGUICommandSetResult = GUI_T_CommandManagement.GetGUICommandDictionaryByReceiveSystem_WaitUUID(enuSystemType.WMS, UUID)
        '    If dicGUICommandSetResult.Any = True Then
        '      '把GUICommand的執行結果寫入DB
        '      Dim lstSQL As New List(Of String)
        '      For Each objGUICommand As clsFromGUICommand In dicGUICommandSetResult.Values
        '        objGUICommand.Result = IIf(Result = True, "0", "1")
        '        objGUICommand.Result_Message = ResultMessage
        '        SendMessageToLog("Set GUI Command Data Report, UUID=" & objGUICommand.UUID & ", Function_ID=" & objGUICommand.Function_ID & ", SEQ=" & objGUICommand.SEQ & ", Result=" & objGUICommand.Result & ", Result_Message=" & objGUICommand.Result_Message & ", Wait_UUID=" & objGUICommand.Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        '        If objGUICommand.O_Add_Update_SQLString(lstSQL) = False Then
        '          SendMessageToLog("Get SQL Faile, UUID=" & objGUICommand.UUID & ", Function_ID=" & objGUICommand.Function_ID & ", SEQ=" & objGUICommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
        '        End If
        '      Next
        '      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        '        SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        '      End If
        '    End If
        'End Select
        '把ToMCSCommand写入历史
        If ProcessToMCSCommandResult_MoveToHist(dicProcessToMCSCommandResult) = True Then

        End If
        '清除已經處理過的HostHandlerCommand
        For Each objToMCSCommand As clsToMCSCommand In dicProcessToMCSCommandResult.Values
          If dicToMCSCommandResult.ContainsKey(objToMCSCommand.gid) = True Then
            dicToMCSCommandResult.Remove(objToMCSCommand.gid)
          End If
        Next
      End While
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '將MCS傳入的Message寫入History(不刪除原來的Message，只有在和MCS不同DB時才進行此操作)
  Private Function ProcessFromMCSCommand_MoveToHist(ByRef tmp_dicFromMCSCommand As Dictionary(Of String, clsFromMCSCommand)) As Boolean
    Try
      Dim lstQueueSQL As New List(Of String)
      Dim Now_Time As String = GetNewTime_DBFormat()
      For Each objFromMCSCommand As clsFromMCSCommand In tmp_dicFromMCSCommand.Values
        With objFromMCSCommand
          '將WMS_T_MCSCommand寫到History
          Dim objFromMCSCommandHist As New clsFromMCSCommandHist(.UUID, .Send_System, .Receive_System, .Function_ID, .SEQ, .User_ID, .Create_Time, .Message, .Result, .Result_Message, .Wait_UUID, Now_Time)
          objFromMCSCommandHist.O_Add_Insert_SQLString(lstQueueSQL)
        End With
      Next
      Common_DBManagement.AddQueued(lstQueueSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '將給MCS的Message移入History
  Private Function ProcessToMCSCommandResult_MoveToHist(ByRef tmp_dicToMCSCommand As Dictionary(Of String, clsToMCSCommand)) As Boolean
    Try
      Dim lstSQL As New List(Of String)
      Dim lstQueueSQL As New List(Of String)
      Dim Now_Time As String = GetNewTime_DBFormat()
      For Each objToMCSCommand As clsToMCSCommand In tmp_dicToMCSCommand.Values
        With objToMCSCommand
          '將WMS_T_MCSCommand寫到History
          Dim objToMCSCommandHist As New clsToMCSCommandHist(.UUID, .Send_System, .Receive_System, .Function_ID, .SEQ, .User_ID, .Create_Time, .Message, .Result, .Result_Message, .Wait_UUID, Now_Time)
          objToMCSCommandHist.O_Add_Insert_SQLString(lstQueueSQL)
        End With
        objToMCSCommand.O_Add_Delete_SQLString(lstSQL)
      Next
      If WMS_T_MCS_CommandManagement.BatchUpdate_DynamicConnection(lstSQL) = False Then
        SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      End If
      Common_DBManagement.AddQueued(lstQueueSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Sub O_thr_NSDBHandling()
    Const SleepTime As Integer = 200
    Dim Count As Integer = 0
    While True
      Try
        'If Count < 10 Then
        '  Count = Count + 1
        'Else
        '  Count = 0
        '  If gMain.int_tNSDBHandle > 99 Then
        '    gMain.int_tNSDBHandle = 0
        '  Else
        '    gMain.int_tNSDBHandle = gMain.int_tNSDBHandle + 1
        '  End If
        'End If
        I_ProcessFromNSMessage()
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("thrNSDBHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub
  Private Function I_ProcessFromNSMessage() As Boolean
    Try
      '  Dim dicFromNSCommand = NS_T_CommandManagement.GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(enuSystemType.HostHandler)
      '  While dicFromNSCommand.Any = True
      '    Dim Function_ID As String = ""
      '    Dim XmlMessage As String = ""
      '    Dim Result As String = ""
      '    Dim ResultMessage As String = ""
      '    Dim Wait_UUID As String = ""
      '    Dim UUID As String = ""
      '    '用來暫存要處理的NSCommand
      '    Dim dicProcessFromNSCommand As New Dictionary(Of String, clsFromNSCommand)
      '    For Each objFromNSCommand As clsFromNSCommand In dicFromNSCommand.Values
      '      If UUID = "" Then
      '        UUID = objFromNSCommand.UUID
      '      End If
      '      '不相等表示是下一筆了，下一次再處理
      '      If UUID <> objFromNSCommand.UUID Then
      '        Exit For
      '      End If
      '      '儲存NSCommand的資訊
      '      If dicProcessFromNSCommand.ContainsKey(objFromNSCommand.gid) = False Then
      '        dicProcessFromNSCommand.Add(objFromNSCommand.gid, objFromNSCommand)
      '      Else
      '        SendMessageToLog(String.Format("NSCommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objFromNSCommand.UUID, objFromNSCommand.Function_ID, objFromNSCommand.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '      End If
      '      Function_ID = objFromNSCommand.Function_ID
      '      XmlMessage = XmlMessage & objFromNSCommand.Message
      '    Next
      '    '把取得的NSCommand送出去進行處理
      '    If O_ProcessCommand(Function_ID, XmlMessage, ResultMessage, Wait_UUID) = True Then
      '      '執行成功
      '      '如果Wait_UUID不為空時才把Wait_UUID填入
      '      If Wait_UUID = "" Then
      '        Result = "0"
      '        'ResultMessage = ""
      '      End If
      '    Else
      '      '執行失敗
      '      Result = "1"
      '    End If
      '    '把NSCommand的執行結果寫入DB
      '    Dim lstSQL As New List(Of String)
      '    For Each objFromNSCommand As clsFromNSCommand In dicProcessFromNSCommand.Values
      '      objFromNSCommand.Result = Result
      '      objFromNSCommand.Result_Message = StrConv(ResultMessage, VbStrConv.TraditionalChinese, 2052) ' ResultMessage
      '      objFromNSCommand.Wait_UUID = Wait_UUID
      '      SendMessageToLog("Set NS Command Data Report, UUID=" & objFromNSCommand.UUID & ", Function_ID=" & objFromNSCommand.Function_ID & ", SEQ=" & objFromNSCommand.SEQ & ", Result=" & objFromNSCommand.Result & ", Result_Message=" & objFromNSCommand.Result_Message & ", Wait_UUID=" & objFromNSCommand.Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      '      If objFromNSCommand.O_Add_Update_SQLString(lstSQL) = False Then
      '        SendMessageToLog("Get SQL Faile, UUID=" & objFromNSCommand.UUID & ", Function_ID=" & objFromNSCommand.Function_ID & ", SEQ=" & objFromNSCommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '      End If
      '    Next
      '    If Common_DBManagement.BatchUpdate(lstSQL) = False Then
      '      SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '    End If
      '    '刪除已經處理過的NSCommand
      '    For Each objFromNSCommand As clsFromNSCommand In dicProcessFromNSCommand.Values
      '      If dicFromNSCommand.ContainsKey(objFromNSCommand.gid) = True Then
      '        dicFromNSCommand.Remove(objFromNSCommand.gid)
      '      End If
      '    Next
      '  End While
      Return True
    Catch ex As Exception
      '  SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
