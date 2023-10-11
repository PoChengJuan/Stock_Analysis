Public Class clsMAINTENANCE_STATUS
	Private ShareName As String = "STATUS"
	Private ShareKey As String = ""
	Private _gid As String
	Private _FACTORY_NO As String '廠別

	Private _DEVICE_NO As String 'LCS的編號

	Private _AREA_NO As String 'Area編號

  Private _UNIT_ID As String '設備編號

  Private _MAINTENANCE_ID As String '設備中的哪一個Maintenance的保養設定

  Private _FUNCTION_ID As String '設備中的哪一個Function的保養設定，對應到底層上報的Function_ID(目前不由底層上報)

  Private _VALUE As String '當前的值(數值/時間)

  Private _UPDATE_TIME As String '更新的時間

  Private _MAINTENANCE_SET As Boolean '是否檢查過並發了保養資訊

  Private _MAINTENANCE_TIME As String '保養資訊發送時間

  Private _objHandling As clsHandlingObject

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
  Public Property DEVICE_NO() As String
    Get
      Return _DEVICE_NO
    End Get
    Set(ByVal value As String)
      _DEVICE_NO = value
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
  Public Property UNIT_ID() As String
    Get
      Return _UNIT_ID
    End Get
    Set(ByVal value As String)
      _UNIT_ID = value
    End Set
  End Property
  Public Property MAINTENANCE_ID() As String
    Get
      Return _MAINTENANCE_ID
    End Get
    Set(ByVal value As String)
      _MAINTENANCE_ID = value
    End Set
  End Property
  Public Property FUNCTION_ID() As String
    Get
      Return _FUNCTION_ID
    End Get
    Set(ByVal value As String)
      _FUNCTION_ID = value
    End Set
  End Property
  Public Property VALUE() As String
    Get
      Return _VALUE
    End Get
    Set(ByVal value As String)
      _VALUE = value
    End Set
  End Property
  Public Property UPDATE_TIME() As String
    Get
      Return _UPDATE_TIME
    End Get
    Set(ByVal value As String)
      _UPDATE_TIME = value
    End Set
  End Property
  Public Property MAINTENANCE_SET() As Boolean
    Get
      Return _MAINTENANCE_SET
    End Get
    Set(ByVal value As Boolean)
      _MAINTENANCE_SET = value
    End Set
  End Property
  Public Property MAINTENANCE_TIME() As String
    Get
      Return _MAINTENANCE_TIME
    End Get
    Set(ByVal value As String)
      _MAINTENANCE_TIME = value
    End Set
  End Property
  Public Property objHandling() As clsHandlingObject
    Get
      Return _objHandling
    End Get
    Set(ByVal value As clsHandlingObject)
      _objHandling = value
    End Set
  End Property

  Public Sub New(ByVal FACTORY_NO As String, ByVal DEVICE_NO As String, ByVal AREA_NO As String, ByVal UNIT_ID As String,
                 ByVal MAINTENANCE_ID As String, ByVal FUNCTION_ID As String, ByVal VALUE As String, ByVal UPDATE_TIME As String,
                 ByVal MAINTENANCE_SET As Boolean, ByVal MAINTENANCE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(FACTORY_NO, DEVICE_NO, AREA_NO, UNIT_ID, MAINTENANCE_ID, FUNCTION_ID)
      _gid = key
      _FACTORY_NO = FACTORY_NO
      _DEVICE_NO = DEVICE_NO
      _AREA_NO = AREA_NO
      _UNIT_ID = UNIT_ID
      _MAINTENANCE_ID = MAINTENANCE_ID
      _FUNCTION_ID = FUNCTION_ID
      _VALUE = VALUE
      _UPDATE_TIME = UPDATE_TIME
      _MAINTENANCE_SET = MAINTENANCE_SET
      _MAINTENANCE_TIME = MAINTENANCE_TIME
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
  Public Shared Function Get_Combination_Key(ByVal FACTORY_NO As String, ByVal DEVICE_NO As String, ByVal AREA_NO As String,
                        ByVal UNIT_ID As String, ByVal MAINTENANCE_ID As String, ByVal FUNCTION_ID As String) As String
    Try
      Dim key As String = FACTORY_NO & LinkKey & DEVICE_NO & LinkKey & AREA_NO & LinkKey & UNIT_ID & LinkKey & MAINTENANCE_ID & LinkKey & FUNCTION_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMAINTENANCE_STATUS
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  Public Sub Add_Relationship(ByRef objHandling As clsHandlingObject)
    Try
      If objHandling IsNot Nothing Then
        _objHandling = objHandling
        objHandling.O_Add_MaintenanceStatus(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_MaintenanceStatus(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_MAINTENANCE_STATUSManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_MAINTENANCE_STATUSManagement.GetUpdateSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Delete的SQL
  Public Function O_Add_Delete_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_MAINTENANCE_STATUSManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_T_MAINTENANCE_STATUS As clsMAINTENANCE_STATUS) As Boolean
    Try
      Dim key As String = objWMS_T_MAINTENANCE_STATUS._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _FACTORY_NO = FACTORY_NO
      _DEVICE_NO = DEVICE_NO
      _AREA_NO = AREA_NO
      _UNIT_ID = UNIT_ID
      _MAINTENANCE_ID = MAINTENANCE_ID
			_FUNCTION_ID = FUNCTION_ID
			_VALUE = VALUE
			_UPDATE_TIME = UPDATE_TIME
			_MAINTENANCE_SET = MAINTENANCE_SET
			_MAINTENANCE_TIME = MAINTENANCE_TIME
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
