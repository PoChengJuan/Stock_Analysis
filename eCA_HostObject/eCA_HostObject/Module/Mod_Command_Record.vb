Public Module Mod_Command_Record

  '執行後會直接進行資料庫資料更新(Handle)

  ''' <summary>
  ''' 將Host發給其他系統的Command記錄到DB
  ''' </summary>
  ''' <param name="FUNCTION_ID"></param>
  ''' <param name="CONNECTION_TYPE"></param>
  ''' <param name="RECEIVE_SYSTEM"></param>
  ''' <param name="UUID">若空則給GUID</param>
  ''' <param name="CREATE_TIME">若空則給當前時間</param>
  ''' <param name="MESSAGE">內部自動切割最大長度，並區分SEQ</param>
  ''' <param name="RESULT">長度10若超過10會自動切割</param>
  ''' <param name="RESULT_MESSAGE">長度超過3500會自動切割</param>
  ''' <param name="WAIT_UUID">長度超過100會自動切割</param>
  ''' <param name="USER_ID">長度超過50會自動切割</param>
  ''' <param name="CLIENT_ID">長度超過50會自動切割</param>
  ''' <param name="IP">長度超過50會自動切割</param>
  Public Sub O_Handle_Host_To_HS_Command_Hist(ByVal FUNCTION_ID As String, ByVal CONNECTION_TYPE As enuConnectionType, ByVal RECEIVE_SYSTEM As enuSystemType, ByVal UUID As String,
                                              ByVal CREATE_TIME As String, ByVal MESSAGE As String, ByVal RESULT As String, ByVal RESULT_MESSAGE As String, ByVal WAIT_UUID As String,
                                              ByVal USER_ID As String, ByVal CLIENT_ID As String, ByVal IP As String)
    Try
      Dim DBMaxLength As Long = 3500
      If FUNCTION_ID.Length > 50 Then FUNCTION_ID = FUNCTION_ID.Substring(0, 50)
      If USER_ID.Length > 50 Then USER_ID = USER_ID.Substring(0, 50)
      If CLIENT_ID.Length > 50 Then CLIENT_ID = CLIENT_ID.Substring(0, 50)
      If IP.Length > 50 Then IP = IP.Substring(0, 50)
      If CREATE_TIME.Length > 20 Then CREATE_TIME = CREATE_TIME.Substring(0, 20)
      If RESULT.Length > 10 Then RESULT = RESULT.Substring(0, 10)
      If WAIT_UUID.Length > 100 Then WAIT_UUID = WAIT_UUID.Substring(0, 100)
      If UUID.Length > 100 Then UUID = UUID.Substring(0, 100)
      If RESULT_MESSAGE.Length > DBMaxLength Then RESULT_MESSAGE = RESULT_MESSAGE.Substring(0, DBMaxLength)

      If UUID = "" Then UUID = Get_System_GUID() '如果UUID為空 則給系統編號
      If CREATE_TIME = "" Then CREATE_TIME = GetNewTime_DBFormat() '如果CREATE_TIME為空 則給當前時間

      Dim dicHostToHSCommand As New Dictionary(Of String, clsHostToHSCommand)
      Dim Now_Time = GetNewTime_DBFormat()
      '如果strMessage超過DBMaxLength個字就要分成多個Message(如果只有一個Message則SEQ填0)
      MESSAGE = MESSAGE.Replace("'", "''")
      If MESSAGE.Length < DBMaxLength Then
        Dim objHostToHSCommand = New clsHostToHSCommand(UUID, CONNECTION_TYPE, RECEIVE_SYSTEM, FUNCTION_ID, 0, USER_ID, CLIENT_ID, IP, CREATE_TIME, MESSAGE, RESULT, RESULT_MESSAGE, WAIT_UUID, Now_Time)
        If dicHostToHSCommand.ContainsKey(objHostToHSCommand.gid) = False Then dicHostToHSCommand.Add(objHostToHSCommand.gid, objHostToHSCommand)
      Else
        Dim Seq As Long = 1
        Do
          Dim NewMessage As String = ""
          If MESSAGE.Length > DBMaxLength Then
            NewMessage = MESSAGE.Substring(0, DBMaxLength)
            MESSAGE = MESSAGE.Substring(DBMaxLength)
          Else
            NewMessage = MESSAGE.Substring(0, MESSAGE.Length)
            MESSAGE = ""
          End If

          Dim objHostToHSCommand = New clsHostToHSCommand(UUID, CONNECTION_TYPE, RECEIVE_SYSTEM, FUNCTION_ID, Seq, USER_ID, CLIENT_ID, IP, CREATE_TIME, NewMessage, RESULT, RESULT_MESSAGE, WAIT_UUID, Now_Time)
          If dicHostToHSCommand.ContainsKey(objHostToHSCommand.gid) = False Then dicHostToHSCommand.Add(objHostToHSCommand.gid, objHostToHSCommand)

          Seq = Seq + 1
        Loop While (MESSAGE.Length > 0)
      End If

      If dicHostToHSCommand.Any Then
        Dim lstQueueSql As New List(Of String)
        For Each obj In dicHostToHSCommand.Values
          obj.O_Add_Insert_SQLString(lstQueueSql)
        Next
        Common_DBManagement.AddQueued(lstQueueSql)

      End If

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  ''' <summary>
  ''' 將其他系統發給Host的Command記錄到DB
  ''' </summary>
  ''' <param name="FUNCTION_ID"></param>
  ''' <param name="CONNECTION_TYPE"></param>
  ''' <param name="SEND_SYSTEM"></param>
  ''' <param name="UUID">若空則給GUID</param>
  ''' <param name="CREATE_TIME">若空則給當前時間</param>
  ''' <param name="MESSAGE">內部自動切割最大長度，並區分SEQ</param>
  ''' <param name="RESULT">長度10若超過10會自動切割</param>
  ''' <param name="RESULT_MESSAGE">長度超過3500會自動切割</param>
  ''' <param name="WAIT_UUID">長度超過100會自動切割</param>
  ''' <param name="USER_ID">長度超過50會自動切割</param>
  ''' <param name="CLIENT_ID">長度超過50會自動切割</param>
  ''' <param name="IP">長度超過50會自動切割</param>
  Public Sub O_Handle_HS_To_Host_Command_Hist(ByVal FUNCTION_ID As String, ByVal CONNECTION_TYPE As enuConnectionType, ByVal SEND_SYSTEM As enuSystemType, ByVal UUID As String,
                                              ByVal CREATE_TIME As String, ByVal MESSAGE As String, ByVal RESULT As String, ByVal RESULT_MESSAGE As String, ByVal WAIT_UUID As String,
                                              ByVal USER_ID As String, ByVal CLIENT_ID As String, ByVal IP As String)
    Try
      Dim DBMaxLength As Long = 4000
      If FUNCTION_ID.Length > 50 Then FUNCTION_ID = FUNCTION_ID.Substring(0, 50)
      If USER_ID.Length > 50 Then USER_ID = USER_ID.Substring(0, 50)
      If CLIENT_ID.Length > 50 Then CLIENT_ID = CLIENT_ID.Substring(0, 50)
      If IP.Length > 50 Then IP = IP.Substring(0, 50)
      If CREATE_TIME.Length > 20 Then CREATE_TIME = CREATE_TIME.Substring(0, 20)
      If RESULT.Length > 10 Then RESULT = RESULT.Substring(0, 10)
      If WAIT_UUID.Length > 100 Then WAIT_UUID = WAIT_UUID.Substring(0, 100)
      If UUID.Length > 100 Then UUID = UUID.Substring(0, 100)
      If RESULT_MESSAGE.Length > DBMaxLength Then RESULT_MESSAGE = RESULT_MESSAGE.Substring(0, DBMaxLength)

      If UUID = "" Then UUID = Get_System_GUID() '如果UUID為空 則給系統編號
      If CREATE_TIME = "" Then CREATE_TIME = GetNewTime_DBFormat() '如果CREATE_TIME為空 則給當前時間

      Dim dicHSToHostCommand As New Dictionary(Of String, clsHSToHostCommand)
      Dim Now_Time = GetNewTime_DBFormat()
      '如果strMessage超過DBMaxLength個字就要分成多個Message(如果只有一個Message則SEQ填0)
      If MESSAGE.Length < DBMaxLength Then
        Dim objHostToHSCommand = New clsHSToHostCommand(UUID, CONNECTION_TYPE, SEND_SYSTEM, FUNCTION_ID, 0, USER_ID, CLIENT_ID, IP, CREATE_TIME, MESSAGE, RESULT, RESULT_MESSAGE, WAIT_UUID, Now_Time)
        If dicHSToHostCommand.ContainsKey(objHostToHSCommand.gid) = False Then dicHSToHostCommand.Add(objHostToHSCommand.gid, objHostToHSCommand)
      Else
        Dim Seq As Long = 1
        Do
          Dim NewMessage As String = ""
          If MESSAGE.Length > DBMaxLength Then
            NewMessage = MESSAGE.Substring(0, DBMaxLength)
            MESSAGE = MESSAGE.Substring(DBMaxLength)
          Else
            NewMessage = MESSAGE.Substring(0, MESSAGE.Length)
            MESSAGE = ""
          End If

          Dim objHostToHSCommand = New clsHSToHostCommand(UUID, CONNECTION_TYPE, SEND_SYSTEM, FUNCTION_ID, Seq, USER_ID, CLIENT_ID, IP, CREATE_TIME, NewMessage, RESULT, RESULT_MESSAGE, WAIT_UUID, Now_Time)
          If dicHSToHostCommand.ContainsKey(objHostToHSCommand.gid) = False Then dicHSToHostCommand.Add(objHostToHSCommand.gid, objHostToHSCommand)

          Seq = Seq + 1
        Loop While (MESSAGE.Length > 0)
      End If

      Dim lstQueueSql As New List(Of String)
      If dicHSToHostCommand.Any Then
        For Each obj In dicHSToHostCommand.Values
          If obj.O_Add_Insert_SQLString(lstQueueSql) = True Then
          End If
        Next
        Common_DBManagement.AddQueued(lstQueueSql)

      End If
      'If O_Get_Insert_SQLString(dicHSToHostCommand, lstQueueSql) Then
      'End If

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub







  '執行後會回傳需要進行更新的資料(Process)



End Module



