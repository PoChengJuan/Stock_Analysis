Public Class clsMAINTENANCE
	Private ShareName As String = "MAINTENANCE"
	Private ShareKey As String = ""
	Private _gid As String
	Private _FACTORY_NO As String '廠別

	Private _DEVICE_NO As String 'LCS的編號

	Private _AREA_NO As String 'Area編號

  Private _UNIT_ID As String '設備編號

  Private _MAINTENANCE_ID As String '設備中的哪一個Maintenance的保養設定

  Private _MAINTENANCE_NAME As String 'Maintenance_Name

  Private _CONTINUE_SEND As Double '是否持續發送(0:只發送一次、1:間隔一定時間發送一次)

  Private _SEND_INTERVAL As Double '發送間隔時間(S)

  Private _SEND_TYPE As Double '發送類型(暫定不用)

  Private _UPDATE_TIME As String '修改的日期

  Private _ENABLE As Boolean '是否啟用0:禁用1:啟用

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
  Public Property MAINTENANCE_NAME() As String
    Get
      Return _MAINTENANCE_NAME
    End Get
    Set(ByVal value As String)
      _MAINTENANCE_NAME = value
    End Set
  End Property
  Public Property CONTINUE_SEND() As Double
    Get
      Return _CONTINUE_SEND
    End Get
    Set(ByVal value As Double)
      _CONTINUE_SEND = value
    End Set
  End Property
  Public Property SEND_INTERVAL() As Double
    Get
      Return _SEND_INTERVAL
    End Get
    Set(ByVal value As Double)
      _SEND_INTERVAL = value
    End Set
  End Property
  Public Property SEND_TYPE() As Double
    Get
      Return _SEND_TYPE
    End Get
    Set(ByVal value As Double)
      _SEND_TYPE = value
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
  Public Property ENABLE() As Boolean
    Get
      Return _ENABLE
    End Get
    Set(ByVal value As Boolean)
      _ENABLE = value
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
                 ByVal MAINTENANCE_ID As String, ByVal MAINTENANCE_NAME As String, ByVal CONTINUE_SEND As Double,
                 ByVal SEND_INTERVAL As Double, ByVal SEND_TYPE As Double, ByVal UPDATE_TIME As String, ByVal ENABLE As Boolean)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(FACTORY_NO, DEVICE_NO, AREA_NO, UNIT_ID, MAINTENANCE_ID)
      _gid = key
      _FACTORY_NO = FACTORY_NO
      _DEVICE_NO = DEVICE_NO
      _AREA_NO = AREA_NO
      _UNIT_ID = UNIT_ID
      _MAINTENANCE_ID = MAINTENANCE_ID
      _MAINTENANCE_NAME = MAINTENANCE_NAME
      _CONTINUE_SEND = CONTINUE_SEND
      _SEND_INTERVAL = SEND_INTERVAL
      _SEND_TYPE = SEND_TYPE
      _UPDATE_TIME = UPDATE_TIME
      _ENABLE = ENABLE
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
                                             ByVal UNIT_ID As String, ByVal MAINTENANCE_ID As String) As String
    Try
      Dim key As String = FACTORY_NO & LinkKey & DEVICE_NO & LinkKey & AREA_NO & LinkKey & UNIT_ID & LinkKey & MAINTENANCE_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMAINTENANCE
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
        objHandling.O_Add_Maintenance(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_Maintenance(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_M_MAINTENANCEManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_MAINTENANCEManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_MAINTENANCEManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_M_MAINTENANCE As clsMAINTENANCE) As Boolean
    Try
      Dim key As String = objWMS_M_MAINTENANCE._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _FACTORY_NO = FACTORY_NO
      _DEVICE_NO = DEVICE_NO
      _AREA_NO = AREA_NO
      _UNIT_ID = UNIT_ID
      _MAINTENANCE_ID = MAINTENANCE_ID
			_MAINTENANCE_NAME = MAINTENANCE_NAME
			_CONTINUE_SEND = CONTINUE_SEND
			_SEND_INTERVAL = SEND_INTERVAL
			_SEND_TYPE = SEND_TYPE
			_UPDATE_TIME = UPDATE_TIME
			_ENABLE = ENABLE
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
