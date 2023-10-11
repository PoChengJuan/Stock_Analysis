Imports eCA_HostObject


''' <summary>
''' 20181117
''' V1.0.0
''' Benny
''' 处理上報
''' </summary>
Module Module_Auto_Report
  Public Sub O_thr_Auto_Report()
    Const SleepTime As Integer = 60000 '改為一分鐘一次
    '載入設定 若更新設定則須重新載入
    Dim objBusiness As clsBusiness_Rule = Nothing
    If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.Report_Time, objBusiness) Then
      For Each _ReportTime In objBusiness.Rule_Value.Split(",")
        ReportTime.Add(_ReportTime)
      Next
    End If

    While True
      Try
        '取得上次上報時間
        Dim LastReportTime = gMain.objHandling.gdicSystemStatus(enuSystemStatus.LastReportTime)
        '檢查當前時間
        Dim _DateTime = DateTime.Now.ToString(DBTimeFormat)
        '更新上報時間
        Dim lstSQL As New List(Of String)

        '檢查這個小時是否需上報
        'Dim bln_Check = False
        'For Each obj In ReportTime
        '  If Hour(obj) = Hour(_DateTime) Then '這個小時需上報
        '    bln_Check = True
        '    Exit For
        '  End If
        'Next
        '如果需上報 檢查上次是哪個時間
        'If bln_Check Then
        '  If Hour(LastReportTime.UPDATE_TIME) < Hour(_DateTime) Or DateAndTime.Day(LastReportTime.UPDATE_TIME) < DateAndTime.Day(_DateTime) Then
        '    '取得須上報的時間點list 最小單位為小時
        '    I_Auto_Report()

        '  End If
        'End If

      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_Auto_Report End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub
  Private Sub Hist_Report()
    Try


    Catch ex As Exception
      SendMessageToLog(ex.Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub UpdateDB(ByVal dicUpdateProductionInfo As Dictionary(Of String, clsProduce_Info), ByVal dicDeleteProductioninfo As Dictionary(Of String, clsProduce_Info),
                                           ByVal dicAddProductioninfohist As Dictionary(Of String, clsProduce_Hist))
    Try
      Dim lstSQL As New List(Of String)
      'update
      For Each obj As clsProduce_Info In dicUpdateProductionInfo.Values
        obj.O_Add_Update_SQLString(lstSQL)
      Next
      'delete
      For Each obj As clsProduce_Info In dicDeleteProductioninfo.Values
        obj.O_Add_Delete_SQLString(lstSQL)
      Next
      'Insert
      For Each obj As clsProduce_Hist In dicAddProductioninfohist.Values
        obj.O_Add_Insert_SQLString(lstSQL)
      Next

      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        '更新DB失敗則回傳False
        'ret_ResultMsg = "WMS Update DB Failed"
        Dim strLog = "WMS 更新资料库失败"
        SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      Else
        '更新記憶體資料									
        '1.更新
        For Each objNew As clsProduce_Info In dicUpdateProductionInfo.Values
          Dim obj As clsProduce_Info = Nothing
          If gMain.objHandling.gdicProduce_Info.TryGetValue(objNew.gid, obj) = True Then
            obj.Update_To_Memory(objNew)
          End If
        Next

        'delete
        For Each objNew As clsProduce_Info In dicDeleteProductioninfo.Values
          If gMain.objHandling.gdicProduce_Info.ContainsKey(objNew.gid) Then
            gMain.objHandling.gdicProduce_Info.Remove(objNew.gid)
          End If
        Next


      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub



  ''' 20190219 修改後的 
  ''' 根據 Area_Type1 判斷前後製程 
  ''' 根據 AREA_INDEX 找出下一個製程的資訊



  Public Sub I_Auto_Report()
    SyncLock gMain.objHandling.objLineProduction_InfoLock
      Try
        SendMessageToLog("I_Auto_Report Start", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        'Hist_Report() '從歷史資料取的資料
        'Current_Report_2()  '從Current 上報 '-

      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      End Try
    End SyncLock
  End Sub

  Public Function ClassResumeByFactory_AreaNo(ByVal factory As String, ByVal areano As String, ByVal startTime As String,
                                              ByVal finishTime As String, ByRef dic As Dictionary(Of String, Dictionary(Of String, clsResume))) As Boolean
    'dic(日期,(班別,資訊))
    Try
      Dim _dicClass = gMain.objHandling.gdicClass
      Dim _gdicClassAttendance = gMain.objHandling.gdicClassAttendance
      Dim _gdicClassAssignation = gMain.objHandling.gdicClassAssignation
      Dim _dicDayClass As New Dictionary(Of String, clsClass)
      Dim _dicNightClass As New Dictionary(Of String, clsClass)
      For Each item In _dicClass
        If item.Value.CLASS_START_TIME > item.Value.CLASS_END_TIME Then '-白班正常時間
          If _dicDayClass.ContainsKey(item.Key) = False Then
            _dicDayClass.Add(item.Key, item.Value)
          End If
        Else '-有跨夜班的
          If _dicNightClass.ContainsKey(item.Key) = False Then
            _dicNightClass.Add(item.Key, item.Value)
          End If
        End If
      Next

      '-要先知道經過幾個班別
      While finishTime >= startTime
        Dim _day = ParseTimeForDay(startTime, DBTimeFormat) '-取得日期
        Dim _Time = ParseTimeForTime(startTime, DBTimeFormat) '-取得時間
        '-判斷時間落在哪個班別
        Dim _check1 = From item In _dicDayClass Where item.Value.CLASS_START_TIME <= _Time AndAlso item.Value.CLASS_END_TIME >= _Time
        Dim _check2 = From item In _dicNightClass Where ((item.Value.CLASS_START_TIME <= _Time AndAlso _Time <= "23:59:59") OrElse (item.Value.CLASS_END_TIME >= _Time AndAlso _Time >= "00:00:00"))

        '-一定會落在一個班別上
        If (_check1.Count <> 0 And _check2.Count = 0) Or (_check1.Count = 0 And _check2.Count <> 0) Then
        Else
          SendMessageToLog("Class Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        End If
        '---
        If _check1.Count <> 0 Then '-落點在白天班
          Dim _class = _check1.First.Value
          If dic.ContainsKey(_day) Then
            If dic(_day).ContainsKey(_class.CLASS_NO) = False Then
              Dim _addResume As New clsResume
              '根據班別 找到 每個班別出席人數
              If _gdicClassAttendance.ContainsKey(_class.CLASS_NO) Then
                '再根據班別  factory and AreaNo 找到相對應的
                Dim _key = clsCLASS_ASSIGNATION.Get_Combination_Key(factory, areano, _class.CLASS_NO)
                '人時
                _addResume.Resume_Time = SubTractTime_Second(_class.CLASS_END_TIME, _class.CLASS_START_TIME) * _gdicClassAssignation(_key).ASSIGNATION_RATE '-這個班別經歷的時間(S)
                If _gdicClassAssignation.ContainsKey(_key) Then
                  '_addResume.Class_PerSon = _gdicClassAttendance(_class.CLASS_NO).ATTENDANCE_COUNT * _gdicClassAssignation(_key).ASSIGNATION_RATE
                  _addResume.Class_PerSon = _gdicClassAttendance(_class.CLASS_NO).ATTENDANCE_COUNT '* _gdicClassAssignation(_key).ASSIGNATION_RATE '先不看比例
                Else
                  SendMessageToLog("找無此班別,班別:" & _class.CLASS_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
                  _addResume.Class_PerSon = 0 '-當下班別的人數
                End If
              Else
                SendMessageToLog("找無此班別,班別:" & _class.CLASS_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
                _addResume.Class_PerSon = 0 '-當下班別的人數
              End If

              dic(_day).Add(_class.CLASS_NO, _addResume)
            End If

          Else
            Dim _newDic As New Dictionary(Of String, clsResume)
            Dim _addResume As New clsResume
            '根據班別 找到 每個班別出席人數
            If _gdicClassAttendance.ContainsKey(_class.CLASS_NO) Then
              '再根據班別  factory and AreaNo 找到相對應的
              Dim _key = clsCLASS_ASSIGNATION.Get_Combination_Key(factory, areano, _class.CLASS_NO)
              '人時
              _addResume.Resume_Time = SubTractTime_Second(_class.CLASS_END_TIME, _class.CLASS_START_TIME) * _gdicClassAssignation(_key).ASSIGNATION_RATE '-這個班別經歷的時間(S)
              If _gdicClassAssignation.ContainsKey(_key) Then
                '_addResume.Class_PerSon = _gdicClassAttendance(_class.CLASS_NO).ATTENDANCE_COUNT * _gdicClassAssignation(_key).ASSIGNATION_RATE
                _addResume.Class_PerSon = _gdicClassAttendance(_class.CLASS_NO).ATTENDANCE_COUNT '* _gdicClassAssignation(_key).ASSIGNATION_RATE
              Else
                SendMessageToLog("找無此班別,班別:" & _class.CLASS_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
                _addResume.Class_PerSon = 0 '-當下班別的人數
              End If
            Else
              SendMessageToLog("找無此班別,班別:" & _class.CLASS_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
              _addResume.Class_PerSon = 0 '-當下班別的人數
            End If
            _addResume.CLASS_MANAGER = _class.CLASS_MANAGER '-當下主管
            _newDic.Add(_class.CLASS_NO, _addResume)
            dic.Add(_day, _newDic)
          End If
        End If
        '----
        If _check2.Count <> 0 Then '-落點在跨日班
          Dim _class = _check2.First.Value
          If dic.ContainsKey(_day) Then
            If dic(_day).ContainsKey(_class.CLASS_NO) = False Then
              Dim _addResume As New clsResume
              _addResume.Resume_Time = OneDaySeconds - SubTractTime_Second(_class.CLASS_END_TIME, _class.CLASS_START_TIME) '-這個班別經歷的時間(S)							
              '根據班別 找到 每個班別出席人數
              If _gdicClassAttendance.ContainsKey(_class.CLASS_NO) Then
                Dim _key = clsCLASS_ASSIGNATION.Get_Combination_Key(factory, areano, _class.CLASS_NO) '再根據班別  factory and AreaNo 找到相對應的
                If _gdicClassAssignation.ContainsKey(_key) Then
                  _addResume.Class_PerSon = _gdicClassAttendance(_class.CLASS_NO).ATTENDANCE_COUNT * _gdicClassAssignation(_key).ASSIGNATION_RATE
                Else
                  SendMessageToLog("找無此班別,班別:" & _class.CLASS_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
                  _addResume.Class_PerSon = 0 '-當下班別的人數
                End If
              Else
                SendMessageToLog("找無此班別,班別:" & _class.CLASS_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
                _addResume.Class_PerSon = 0 '-當下班別的人數
              End If

              dic(_day).Add(_class.CLASS_NO, _addResume)
            End If

          Else
            Dim _newDic As New Dictionary(Of String, clsResume)
            Dim _addResume As New clsResume
            _addResume.Resume_Time = OneDaySeconds - SubTractTime_Second(_class.CLASS_END_TIME, _class.CLASS_START_TIME) '-這個班別經歷的時間(S)							
            If _gdicClassAttendance.ContainsKey(_class.CLASS_NO) Then '根據班別 找到 每個班別出席人數
              Dim _key = clsCLASS_ASSIGNATION.Get_Combination_Key(factory, areano, _class.CLASS_NO) '再根據班別  factory and AreaNo 找到相對應的
              If _gdicClassAssignation.ContainsKey(_key) Then
                _addResume.Class_PerSon = _gdicClassAttendance(_class.CLASS_NO).ATTENDANCE_COUNT * _gdicClassAssignation(_key).ASSIGNATION_RATE
              Else
                SendMessageToLog("找無此班別,班別:" & _class.CLASS_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
                _addResume.Class_PerSon = 0 '-當下班別的人數
              End If
            Else
              SendMessageToLog("找無此班別,班別:" & _class.CLASS_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
              _addResume.Class_PerSon = 0 '-當下班別的人數
            End If
            _addResume.CLASS_MANAGER = _class.CLASS_MANAGER '-當下主管
            _newDic.Add(_class.CLASS_NO, _addResume)
            dic.Add(_day, _newDic)
          End If
        End If

        '時間增加一個小時
        startTime = AddTractTime_Hour(startTime, 1)

      End While


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  Public Class clsResume
    Public Resume_Time As Double = 0 '-這個班別經歷的時間(S)
    Public Class_PerSon As Double = 0 '-當下班別的人數
    Public CLASS_MANAGER As String = "" '-當下主管
  End Class
End Module
