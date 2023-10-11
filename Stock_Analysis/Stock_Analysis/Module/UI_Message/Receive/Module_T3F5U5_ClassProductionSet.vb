'2019/1/16 上午 11:58:24
'V1.0.0
'Tool
'狀態:Checked

Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T3F5U5_ClassProductionSet

	Public Function O_Process_Message(ByRef Receive_Msg As MSG_T3F5U5_ClassProductionSet, ByRef Result_Message As String) As Boolean
		Try
      '異動的資料
      Dim dicUpdateClassAssignation As New Dictionary(Of String, clsCLASS_ASSIGNATION)
      Dim dicUpdateClassAttendance As New Dictionary(Of String, clsCLASS_ATTENDANCE)
      Dim lstAddClassAssignationHist As New List(Of clsCLASS_ASSIGNATION_HIST)
      Dim lstAddClassAttendanceHist As New List(Of clsCLASS_ATTENDANCE_HIST)

      Dim lstSql As New List(Of String) '儲存要更新的SQL，進行一次性更新
      Dim lstQueueSql As New List(Of String) '儲存要更新的SQL，進行一次性更新

      '先進行資料邏輯檢查
      If Check_UpdateData(Receive_Msg, Result_Message) = False Then
        Return False
      End If
      '邏輯
      If Get_UpdateData(Receive_Msg, Result_Message, dicUpdateClassAssignation, dicUpdateClassAttendance, lstAddClassAssignationHist, lstAddClassAttendanceHist) = False Then
        Return False
      End If
      '取得SQL
      If Get_SQL(Result_Message, dicUpdateClassAssignation, dicUpdateClassAttendance, lstAddClassAssignationHist, lstAddClassAttendanceHist, lstSql, lstQueueSql) = False Then
        Return False
      End If
      '執行資料更新
      If Execute_DataUpdate(Result_Message, dicUpdateClassAssignation, dicUpdateClassAttendance, lstSql, lstQueueSql) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Check_UpdateData(ByRef Receive_Msg As MSG_T3F5U5_ClassProductionSet,
                                    ByRef Result_Message As String) As Boolean
    Try
      For Each objWrokDataInfo In Receive_Msg.Body.ClassList.ClassInfo
        If objWrokDataInfo.CLASS_NO = "" Then
          Result_Message = "CLASS_NO is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If IsNumeric(objWrokDataInfo.ATTENDANCE_COUNT) = False Then
          Result_Message = "ATTENDANCE_COUNT is empty"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        'dtl
        For Each item In objWrokDataInfo.AssignationList.AssignationInfo
          If item.FACTORY_NO = "" Then
            Result_Message = "FACTORY_NO is empty"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If item.AREA_NO = "" Then
            Result_Message = "AREA_NO is empty"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If IsNumeric(item.ASSIGNATION_RATE) = False Then
            Result_Message = "ASSIGNATION_RATE is empty"
            SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
        Next
      Next
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Get_UpdateData(ByRef Receive_Msg As MSG_T3F5U5_ClassProductionSet,
                                  ByRef ret_strResultMsg As String,
                                  ByRef ret_dicUpdateClassAssignation As Dictionary(Of String, clsCLASS_ASSIGNATION),
                                  ByRef ret_dicUpdateClassAttendance As Dictionary(Of String, clsCLASS_ATTENDANCE),
                                  ByRef ret_lstAddClassAssignationHist As List(Of clsCLASS_ASSIGNATION_HIST),
                                  ByRef ret_lstAddClasssAttendanceHist As List(Of clsCLASS_ATTENDANCE_HIST)) As Boolean
    Try
      'logic
      Dim UserID = Receive_Msg.Header.ClientInfo.UserID
      Dim Now_Time = GetNewTime_DBFormat()
      For Each ClassInfo In Receive_Msg.Body.ClassList.ClassInfo
        Dim CLASS_NO = ClassInfo.CLASS_NO
        Dim ATTENDANCE_COUNT = ClassInfo.ATTENDANCE_COUNT
        Dim tmp_ClassAttendance As New Dictionary(Of String, clsCLASS_ATTENDANCE)
        If gMain.objHandling.O_Get_dicClassAttendanceByClassNo(CLASS_NO, tmp_ClassAttendance) = False Then
          ret_strResultMsg = "Cant Get Class Attendance Data"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        If tmp_ClassAttendance.Any = False Then
          ret_strResultMsg = "Cant Get Class Attendance Data"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
        For Each objClassAttendance As clsCLASS_ATTENDANCE In tmp_ClassAttendance.Values
          If ret_dicUpdateClassAttendance.ContainsKey(objClassAttendance.gid) = False Then
            Dim objNewClassAttendance As clsCLASS_ATTENDANCE = objClassAttendance.Clone()
            objNewClassAttendance.ATTENDANCE_COUNT = ATTENDANCE_COUNT
            objNewClassAttendance.UPDATE_USER = UserID
            objNewClassAttendance.UPDATE_TIME = Now_Time
            ret_dicUpdateClassAttendance.Add(objNewClassAttendance.gid, objNewClassAttendance)
            Dim objClassAttendanceHist As New clsCLASS_ATTENDANCE_HIST(CLASS_NO, ATTENDANCE_COUNT, UserID, Now_Time)
            ret_lstAddClasssAttendanceHist.Add(objClassAttendanceHist)
          End If
        Next
        For Each AssignationInfo In ClassInfo.AssignationList.AssignationInfo
          Dim FACTORY_NO = AssignationInfo.FACTORY_NO
          Dim AREA_NO = AssignationInfo.AREA_NO
          Dim ASSIGNATION_RATE = AssignationInfo.ASSIGNATION_RATE
          Dim tmp_dicClassAssignation As New Dictionary(Of String, clsCLASS_ASSIGNATION)
          If gMain.objHandling.O_Get_DicClassAssignationByFactoryNo_AreaNo_ClassNo(FACTORY_NO, AREA_NO, CLASS_NO, tmp_dicClassAssignation) = False Then
            ret_strResultMsg = "Cant Get Class Assignation Data"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          If tmp_dicClassAssignation.Any = False Then
            ret_strResultMsg = "Cant Get Class Assignation Data"
            SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          For Each objClassAssignation As clsCLASS_ASSIGNATION In tmp_dicClassAssignation.Values
            If ret_dicUpdateClassAssignation.ContainsKey(objClassAssignation.gid) = False Then
              Dim objNewClassAssignation As clsCLASS_ASSIGNATION = objClassAssignation.Clone()
              objNewClassAssignation.ASSIGNATION_RATE = ASSIGNATION_RATE
              objNewClassAssignation.UPDATE_USER = UserID
              objNewClassAssignation.UPDATE_TIME = Now_Time
              ret_dicUpdateClassAssignation.Add(objNewClassAssignation.gid, objNewClassAssignation)
              Dim objClassAssignationHist As New clsCLASS_ASSIGNATION_HIST(FACTORY_NO, AREA_NO, CLASS_NO, ASSIGNATION_RATE, UserID, Now_Time)
              ret_lstAddClassAssignationHist.Add(objClassAssignationHist)
            End If
          Next
        Next
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Function Get_SQL(ByRef Result_Message As String,
                           ByRef ret_dicUpdateClassAssignation As Dictionary(Of String, clsCLASS_ASSIGNATION),
                           ByRef ret_dicUpdateClassAttendance As Dictionary(Of String, clsCLASS_ATTENDANCE),
                           ByRef ret_lstAddClassAssignationHist As List(Of clsCLASS_ASSIGNATION_HIST),
                           ByRef ret_lstAddClasssAttendanceHist As List(Of clsCLASS_ATTENDANCE_HIST),
                           ByRef lstSql As List(Of String),
                           ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      For Each obj In ret_dicUpdateClassAssignation.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          Result_Message = "Get Update WMS_CM_CLASS_ASSIGNATION SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj In ret_dicUpdateClassAttendance.Values
        If obj.O_Add_Update_SQLString(lstSql) = False Then
          Result_Message = "Get Update WMS_CM_CLASS_ATTENDANCE SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj In ret_lstAddClassAssignationHist
        If obj.O_Add_Insert_SQLString(lstQueueSql) = False Then
          Result_Message = "Get Update WMS_CH_CLASS_ASSIGNATION_HIST SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj In ret_lstAddClasssAttendanceHist
        If obj.O_Add_Insert_SQLString(lstQueueSql) = False Then
          Result_Message = "Get Update WMS_CH_CLASS_ATTENDANCE_HIST SQL Failed"
          SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Function Execute_DataUpdate(ByRef Result_Message As String,
                                      ByRef ret_dicUpdateClassAssignation As Dictionary(Of String, clsCLASS_ASSIGNATION),
                                      ByRef ret_dicUpdateClassAttendance As Dictionary(Of String, clsCLASS_ATTENDANCE),
                                      ByRef lstSql As List(Of String),
                                      ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSql.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
        '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
        Result_Message = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '更新所有的SQL
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        Result_Message = " Update DB Failed"
        SendMessageToLog(Result_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If

      '執行Queue
      Common_DBManagement.AddQueued_BatchUpdate(lstQueueSql)
      '修改記憶體資料
      For Each objNew As clsCLASS_ASSIGNATION In ret_dicUpdateClassAssignation.Values
        '移除關聯
        Dim obj As clsCLASS_ASSIGNATION = Nothing
        If gMain.objHandling.gdicClassAssignation.TryGetValue(objNew.gid, obj) = True Then
          obj.Update_To_Memory(objNew)
        End If
      Next
      For Each objNew As clsCLASS_ATTENDANCE In ret_dicUpdateClassAttendance.Values
        '移除關聯
        Dim obj As clsCLASS_ATTENDANCE = Nothing
        If gMain.objHandling.gdicClassAttendance.TryGetValue(objNew.gid, obj) = True Then
          obj.Update_To_Memory(objNew)
        End If
      Next
      Return True
    Catch ex As Exception
      Result_Message = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function





End Module
