'20180921
'V1.0.0
'Mark
'狀態:Open

'設備上報Alarm的情況，上報是否有新的異常產生或是清除



Imports eCA_HostObject
Imports eCA_TransactionMessage

Module Module_T3F4R2_DeviceAlarmReport
  Public Function O_Process_Message(ByVal Receive_Msg As MSG_T3F4R2_DeviceAlarmReport,
                                    ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim dicAddAlarm As New Dictionary(Of String, clsALARM)
      Dim dicDeleteAlarm As New Dictionary(Of String, clsALARM)
      Dim dicAddAlarm_HIST As New Dictionary(Of String, clsALARM_HIST)

      Dim lstSQL As New List(Of String)
      Dim lstHistroySQL As New List(Of String)

      '檢查資料
      If Check_Data(Receive_Msg, ret_strResultMsg) = False Then
        Return False
      End If
      ''取得更新資料
      If Get_Data(Receive_Msg, ret_strResultMsg, dicAddAlarm, dicDeleteAlarm, dicAddAlarm_HIST) = False Then
        Return False
      End If
      ''取得要更新到DB的SQL
      If Get_SQL(ret_strResultMsg, dicAddAlarm, dicDeleteAlarm, dicAddAlarm_HIST, lstSQL, lstHistroySQL) = False Then
        Return False
      End If
      ''執行資料更新
      If Execute_DataUpdate(ret_strResultMsg, dicAddAlarm, dicDeleteAlarm, lstSQL, lstHistroySQL) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''檢查資料
  Private Function Check_Data(ByVal Receive_Msg As MSG_T3F4R2_DeviceAlarmReport,
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      For Each AlarmInfo In Receive_Msg.Body.AlarmList.AlarmInfo
        Dim Alarm_Code As String = AlarmInfo.ALARM_CODE
        Dim Alarm_Desc As String = AlarmInfo.ALARM_DESC
        Dim COMMAND_ID As String = AlarmInfo.COMMAND_ID
        Dim Factory_No As String = AlarmInfo.FACTORY_NO
        Dim Device_No As String = AlarmInfo.DEVICE_NO
        Dim Unit_ID As String = AlarmInfo.UNIT_ID
        Dim Time As String = AlarmInfo.TIME
        Dim Set_Flag As String = AlarmInfo.SET_FLAG
        '檢查是否有Device的資料
        If Module_DataFilter.O_Get_dicCLineProductionInfoByFacotryNo_AreaNo_DeviceNo_UnitID(gMain.objHandling.gdicLineProduction_Info, Factory_No, "", Device_No, Unit_ID) = False Then
          ret_strResultMsg += "Unit_ID is not Exist, Factory_No = <" & Factory_No & ">" & ", Device_No = <" & Device_No & ">" & ", Unit_ID = <" & Unit_ID & ">" & vbNewLine
        End If

      Next
      If ret_strResultMsg.Length > 0 Then
        Return False
      Else
        Return True
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''取得更新資料
  Private Function Get_Data(ByRef Receive_Msg As MSG_T3F4R2_DeviceAlarmReport,
                            ByRef ret_strResultMsg As String,
                            ByRef ret_dicAddAlarm As Dictionary(Of String, clsALARM),
                            ByRef ret_dicDeleteAlarm As Dictionary(Of String, clsALARM),
                            ByRef ret_dicAddAlarm_HIST As Dictionary(Of String, clsALARM_HIST)) As Boolean
    Try
      Dim Now_Time As String = GetNewTime_DBFormat()
      For Each AlarmInfo In Receive_Msg.Body.AlarmList.AlarmInfo
        Dim Alarm_Code As String = AlarmInfo.ALARM_CODE
        Dim Alarm_Desc As String = AlarmInfo.ALARM_DESC
        Dim COMMAND_ID As String = AlarmInfo.COMMAND_ID
        Dim Factory_No As String = AlarmInfo.FACTORY_NO
        Dim Device_No As String = AlarmInfo.DEVICE_NO
        Dim Unit_ID As String = AlarmInfo.UNIT_ID
        Dim OCCUR_TIME As String = AlarmInfo.TIME
        Dim Clear_Time As String = AlarmInfo.TIME
        Dim Set_Flag As String = AlarmInfo.SET_FLAG

        Dim Area_NO As String = "" '暂且先填空
        '取得Factory_No和Device_No對應的Area_No
        Dim tmp_dicLineProductionInfo As New Dictionary(Of String, clsLineProduction_Info)
        If Module_DataFilter.O_Get_dicCLineProductionInfoByFacotryNo_AreaNo_DeviceNo_UnitID(gMain.objHandling.gdicLineProduction_Info, Factory_No, "", Device_No, Unit_ID, tmp_dicLineProductionInfo) = True Then
          Area_NO = tmp_dicLineProductionInfo.First.Value.Area_No
        End If
        '检查是否存在
        Dim objAlarm As clsALARM = Nothing
        gMain.objHandling.O_Get_Alarm(Factory_No, Area_NO, Device_No, Unit_ID, Alarm_Code, objAlarm)

        '根据flag检查是否存在或不存在
        Select Case Set_Flag
          Case 1 '异常发生
            '如果有重复的异常
            If objAlarm IsNot Nothing Then
              ret_strResultMsg += "此异常已存在, Factory_No = <" & Factory_No & ">" & ", Device_No = <" & Device_No & ">" &
                                 ", Alarm_Code = <" & Alarm_Code & ">" & ", Unit_ID = <" & Unit_ID & ">" & vbNewLine
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If
            '建立Alarm
            Dim ALARM_TYPE = enuAlarmType.None '先不填
            Dim SEND_STATUS = enuSend_Status.UnChecked
            Dim New_objAlarm As New clsALARM(Factory_No, Area_NO, Device_No, Unit_ID, OCCUR_TIME, Alarm_Code, ALARM_TYPE, COMMAND_ID, SEND_STATUS)
            If ret_dicAddAlarm.ContainsKey(New_objAlarm.gid) = False Then
              ret_dicAddAlarm.Add(New_objAlarm.gid, New_objAlarm)
            End If

          Case 0 '异常结束
            '如果无此异常
            If objAlarm Is Nothing Then
              ret_strResultMsg += "此异常不存在, Factory_No = <" & Factory_No & ">" & ", Device_No = <" & Device_No & ">" &
                                 ", Alarm_Code = <" & Alarm_Code & ">" & ", Unit_ID = <" & Unit_ID & ">" & vbNewLine
              SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If
            '加入删除行列
            If ret_dicDeleteAlarm.ContainsKey(objAlarm.gid) = False Then
              ret_dicDeleteAlarm.Add(objAlarm.gid, objAlarm)
              Dim New_objAlarm_HIST As New clsALARM_HIST(objAlarm.FACTORY_NO, objAlarm.AREA_NO, objAlarm.DEVICE_NO, objAlarm.UNIT_ID,
                                                         objAlarm.OCCUR_TIME, objAlarm.ALARM_CODE, objAlarm.ALARM_TYPE, objAlarm.CMD_ID,
                                                         objAlarm.SEND_STATUS, Clear_Time, Now_Time)
              If ret_dicAddAlarm_HIST.ContainsKey(New_objAlarm_HIST.gid) = False Then
                ret_dicAddAlarm_HIST.Add(New_objAlarm_HIST.gid, New_objAlarm_HIST)
              End If
            End If
        End Select
      Next
      If ret_strResultMsg.Length > 0 Then
        Return False
      Else
        Return True
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef ret_dicAddAlarm As Dictionary(Of String, clsALARM),
                           ByRef ret_dicDeleteAlarm As Dictionary(Of String, clsALARM),
                           ByRef ret_dicAddAlarm_HIST As Dictionary(Of String, clsALARM_HIST),
                           ByRef lstSql As List(Of String),
                           ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '取得新增SQL
      For Each obj As clsALARM In ret_dicAddAlarm.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert WMS_T_ALARM SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      '取得删除SQL
      For Each obj As clsALARM In ret_dicDeleteAlarm.Values
        If obj.O_Add_Delete_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Delete WMS_T_ALARM SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      '取得建立历史资料的QSL
      'Module_Process_History.Get_ALARM_HIST_SQL(lstQueueSql, ret_dicDeleteAlarm)
      For Each obj As clsALARM_HIST In ret_dicAddAlarm_HIST.Values
        If obj.O_Add_Insert_SQLString(lstQueueSql) = False Then
          ret_strResultMsg = "Get Insert WMS_H_ALARM_HIST SQL Failed"
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
  '執行SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                      ByRef ret_dicAddAlarm As Dictionary(Of String, clsALARM),
                                      ByRef ret_dicDeleteAlarm As Dictionary(Of String, clsALARM),
                                      ByRef lstSql As List(Of String),
                                      ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      If lstSql.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
        '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
        ret_strResultMsg = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        'ret_strResultMsg = "WMS Update DB Failed"
        ret_strResultMsg = "WMS 更新资料库失败"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Common_DBManagement.AddQueued_BatchUpdate(lstQueueSql) '歷史紀錄或無須理會的SQL

      '修改記憶體資料
      '1.新增
      For Each objNew As clsALARM In ret_dicAddAlarm.Values
        objNew.Add_Relationship(gMain.objHandling)
      Next
      '2.删除
      For Each objALARM As clsALARM In ret_dicDeleteAlarm.Values
        objALARM.Remove_Relationship()
      Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module
