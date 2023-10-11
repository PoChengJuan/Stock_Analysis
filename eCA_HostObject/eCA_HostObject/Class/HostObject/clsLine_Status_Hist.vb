Public Class clsLine_Status_Hist
	Private ShareName As String = "Line_Hist"
	Private ShareKey As String = ""

	Private _Factory_No As String
	Private _Area_No As String
	Private _Device_No As String
	Private _Unit_ID As String
	Private _From_Status As enuLineStatus
	Private _To_Status As enuLineStatus
	Private _Hist_Time As String

	Public Property Factory_No() As String
		Get
			Return _Factory_No
		End Get
		Set(ByVal value As String)
			_Factory_No = value
		End Set
	End Property
	Public Property Area_No() As String
		Get
			Return _Area_No
		End Get
		Set(ByVal value As String)
			_Area_No = value
		End Set
	End Property
	Public Property Device_No() As String
		Get
			Return _Device_No
		End Get
		Set(ByVal value As String)
			_Device_No = value
		End Set
	End Property
	Public Property Unit_ID() As String
		Get
			Return _Unit_ID
		End Get
		Set(ByVal value As String)
			_Unit_ID = value
		End Set
	End Property
	Public Property From_Status() As enuLineStatus
		Get
			Return _From_Status
		End Get
		Set(ByVal value As enuLineStatus)
			_From_Status = value
		End Set
	End Property
	Public Property To_Status() As enuLineStatus
		Get
			Return _To_Status
		End Get
		Set(ByVal value As enuLineStatus)
			_To_Status = value
		End Set
	End Property
	Public Property Hist_Time() As String
		Get
			Return _Hist_Time
		End Get
		Set(ByVal value As String)
			_Hist_Time = value
		End Set
	End Property

	'物件建立時執行的事件
	Public Sub New(ByVal Factory_No As String,
								 ByVal Area_No As String,
								 ByVal Device_No As String,
								 ByVal Unit_ID As String,
								 ByVal From_Status As enuLineStatus,
								 ByVal To_Status As enuLineStatus,
								 ByVal Hist_Time As String)
		MyBase.New()
		Try
			_Factory_No = Factory_No
			_Area_No = Area_No
			_Device_No = Device_No
			_Unit_ID = Unit_ID
			_From_Status = From_Status
			_To_Status = To_Status
			_Hist_Time = Hist_Time
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
		End Try
	End Sub
	'物件結束時觸發的事件，用來清除物件的內容
	Protected Overrides Sub Finalize()

		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	Private Sub Class_Terminate_Renamed()
		'目的:結束物件
	End Sub

	'取得要Insert的SQL
	Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
		Try
			Dim strSQL As String = WMS_CH_LINE_STATUS_HISTManagement.GetInsertSQL(Me)
			lstSQL.Add(strSQL)
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
