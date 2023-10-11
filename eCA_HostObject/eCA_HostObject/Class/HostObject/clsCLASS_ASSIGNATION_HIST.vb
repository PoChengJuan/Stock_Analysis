Public Class clsCLASS_ASSIGNATION_HIST
	Private ShareName As String = "HIST"
	Private ShareKey As String = ""
	Private _gid As String
	Private _FACTORY_NO As String '廠別

	Private _AREA_NO As String '區域編號

	Private _CLASS_NO As String '班別編號

	Private _ASSIGNATION_RATE As Double '分配比率(%)

	Private _UPDATE_USER As String '更新人員

	Private _HIST_TIME As String '寫入歷史時間

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
	Public Property CLASS_NO() As String
		Get
			Return _CLASS_NO
		End Get
		Set(ByVal value As String)
			_CLASS_NO = value
		End Set
	End Property
	Public Property ASSIGNATION_RATE() As Double
		Get
			Return _ASSIGNATION_RATE
		End Get
		Set(ByVal value As Double)
			_ASSIGNATION_RATE = value
		End Set
	End Property
	Public Property UPDATE_USER() As String
		Get
			Return _UPDATE_USER
		End Get
		Set(ByVal value As String)
			_UPDATE_USER = value
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

	Public Sub New(ByVal FACTORY_NO As String, ByVal AREA_NO As String, ByVal CLASS_NO As String, ByVal ASSIGNATION_RATE As Double, ByVal UPDATE_USER As String, ByVal HIST_TIME As String)
		MyBase.New()
		Try
			_FACTORY_NO = FACTORY_NO
			_AREA_NO = AREA_NO
			_CLASS_NO = CLASS_NO
			_ASSIGNATION_RATE = ASSIGNATION_RATE
			_UPDATE_USER = UPDATE_USER
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
	Public Shared Function Get_Combination_Key(ByVal FACTORY_NO As String, ByVal AREA_NO As String, ByVal CLASS_NO As String) As String
		Try
			Dim key As String = FACTORY_NO & LinkKey & AREA_NO & LinkKey & CLASS_NO
			Return key
 Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return ""
		End Try
	End Function
	Public Function Clone() As clsCLASS_ASSIGNATION_HIST
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
Dim strSQL As String = WMS_CH_CLASS_ASSIGNATION_HISTManagement.GetInsertSQL(Me)
 lstSQL.Add(strSQL)
 Return True
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return False
 End Try
 End Function
	'取得要Update的SQL
	'Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
	' Try
	' Dim strSQL As String = WMS_CH_CLASS_ASSIGNATION_HISTManagement.GetUpdateSQL(Me)
	' lstSQL.Add(strSQL)
	' Return True
	' Catch ex As Exception
	' SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
	' Return False
	' End Try
	' End Function
	'取得要Delete的SQL
	'Public Function O_Add_Delete_SQLString(ByRef lstSQL As List(Of String)) As Boolean
	' Try
	' Dim strSQL As String = WMS_CH_CLASS_ASSIGNATION_HISTManagement.GetDeleteSQL(Me)
	' lstSQL.Add(strSQL)
	' Return True
	' Catch ex As Exception
	' SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
	' Return False
	' End Try
	' End Function
End Class
