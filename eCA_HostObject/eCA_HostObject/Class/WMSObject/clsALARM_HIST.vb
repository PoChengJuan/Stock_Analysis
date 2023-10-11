Public Class clsALARM_HIST
  Private ShareName As String = "WMS_H_ALARM_HIST"
  Private ShareKey As String = ""
  Private _gid As String
  Private _FACTORY_NO As String '廠別

  Private _AREA_NO As String '區域編號

  Private _DEVICE_NO As String '設備編號

  Private _UNIT_ID As String '單元設備編號

  Private _OCCUR_TIME As String '發生時間

  Private _ALARM_CODE As String '異常代碼

  Private _ALARM_TYPE As enuAlarmType '異常類型

  Private _CMD_ID As String '命令編號

  Private _SEND_STATUS As enuSend_Status '是否發送給SendSystem0:確認1:不發送2:已發送3:發送成功4:發送失敗

  Private _CLEAR_TIME As String '清除時間

  Private _HIST_TIME As String '移入歷史的時間

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property FACTORY_NO() As String
    Get
      Return _FACTORY_NO
    End Get
    Set(ByVal value As String)
      _FACTORY_NO = value
    End Set
  End Property
  Public Property AREA_NO() As String
    Get
      Return _AREA_NO
    End Get
    Set(ByVal value As String)
      _AREA_NO = value
    End Set
  End Property
  Public Property DEVICE_NO() As String
    Get
      Return _DEVICE_NO
    End Get
    Set(ByVal value As String)
      _DEVICE_NO = value
    End Set
  End Property
  Public Property UNIT_ID() As String
    Get
      Return _UNIT_ID
    End Get
    Set(ByVal value As String)
      _UNIT_ID = value
    End Set
  End Property
  Public Property OCCUR_TIME() As String
    Get
      Return _OCCUR_TIME
    End Get
    Set(ByVal value As String)
      _OCCUR_TIME = value
    End Set
  End Property
  Public Property ALARM_CODE() As String
    Get
      Return _ALARM_CODE
    End Get
    Set(ByVal value As String)
      _ALARM_CODE = value
    End Set
  End Property
  Public Property ALARM_TYPE() As enuAlarmType
    Get
      Return _ALARM_TYPE
    End Get
    Set(ByVal value As enuAlarmType)
      _ALARM_TYPE = value
    End Set
  End Property
  Public Property CMD_ID() As String
    Get
      Return _CMD_ID
    End Get
    Set(ByVal value As String)
      _CMD_ID = value
    End Set
  End Property
  Public Property SEND_STATUS() As enuSend_Status
    Get
      Return _SEND_STATUS
    End Get
    Set(ByVal value As enuSend_Status)
      _SEND_STATUS = value
    End Set
  End Property
  Public Property CLEAR_TIME() As String
    Get
      Return _CLEAR_TIME
    End Get
    Set(ByVal value As String)
      _CLEAR_TIME = value
    End Set
  End Property
  Public Property HIST_TIME() As String
    Get
      Return _HIST_TIME
    End Get
    Set(ByVal value As String)
      _HIST_TIME = value
    End Set
  End Property

  Public Sub New(ByVal FACTORY_NO As String, ByVal AREA_NO As String, ByVal DEVICE_NO As String, ByVal UNIT_ID As String,
                 ByVal OCCUR_TIME As String, ByVal ALARM_CODE As String, ByVal ALARM_TYPE As enuAlarmType, ByVal CMD_ID As String,
                 ByVal SEND_STATUS As enuSend_Status, ByVal CLEAR_TIME As String, ByVal HIST_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key()
      _gid = key
      _FACTORY_NO = FACTORY_NO
      _AREA_NO = AREA_NO
      _DEVICE_NO = DEVICE_NO
      _UNIT_ID = UNIT_ID
      _OCCUR_TIME = OCCUR_TIME
      _ALARM_CODE = ALARM_CODE
      _ALARM_TYPE = ALARM_TYPE
      _CMD_ID = CMD_ID
      _SEND_STATUS = SEND_STATUS
      _CLEAR_TIME = CLEAR_TIME
      _HIST_TIME = HIST_TIME
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '物件結束時觸發的事件，用來清除物件的內容
  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
  Private Sub Class_Terminate_Renamed()
    '目的:結束物件
  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Function Clone() As clsALARM_HIST
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_H_ALARM_HISTManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_H_ALARM_HIST As clsALARM_HIST) As Boolean
    Try
      Dim key As String = objWMS_H_ALARM_HIST._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _FACTORY_NO = objWMS_H_ALARM_HIST.FACTORY_NO
      _AREA_NO = objWMS_H_ALARM_HIST.AREA_NO
      _DEVICE_NO = objWMS_H_ALARM_HIST.DEVICE_NO
      _UNIT_ID = objWMS_H_ALARM_HIST.UNIT_ID
      _OCCUR_TIME = objWMS_H_ALARM_HIST.OCCUR_TIME
      _ALARM_CODE = objWMS_H_ALARM_HIST.ALARM_CODE
      _ALARM_TYPE = objWMS_H_ALARM_HIST.ALARM_TYPE
      _CMD_ID = objWMS_H_ALARM_HIST.CMD_ID
      _SEND_STATUS = objWMS_H_ALARM_HIST.SEND_STATUS
      _CLEAR_TIME = objWMS_H_ALARM_HIST.CLEAR_TIME
      _HIST_TIME = objWMS_H_ALARM_HIST.HIST_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function Get_Combination_Key() As String
    Try
      Dim key As String = ""
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
End Class
