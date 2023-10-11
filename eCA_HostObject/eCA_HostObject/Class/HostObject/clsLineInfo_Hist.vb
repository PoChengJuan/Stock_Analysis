Public Class clsLineInfo_Hist
  Private ShareName As String = "LineInfo_Hist"
  Private ShareKey As String = ""

  Private _Factory_No As String
  Private _Area_No As String
  Private _Device_No As String
  Private _Unit_ID As String
  Private _Occur_Time As String
  Private _Maintenance_Message As String
  Private _Remove_User As String
	Private _Hist_Time As String

	Private _MAINTENANCE_ID As String
	Private _FUCTION_ID As String
	Private _OPERATOR_USER As String
	Private _COMMENTS As String


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
  Public Property Occur_Time() As String
    Get
      Return _Occur_Time
    End Get
    Set(ByVal value As String)
      _Occur_Time = value
    End Set
  End Property
  Public Property Maintenance_Message() As String
    Get
      Return _Maintenance_Message
    End Get
    Set(ByVal value As String)
      _Maintenance_Message = value
    End Set
  End Property
  Public Property Remove_User() As String
    Get
      Return _Remove_User
    End Get
    Set(ByVal value As String)
      _Remove_User = value
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

	Public Property MAINTENANCE_ID() As String
		Get
			Return _MAINTENANCE_ID
		End Get
		Set(ByVal value As String)
			_MAINTENANCE_ID = value
		End Set
	End Property
	Public Property FUCTION_ID() As String
		Get
			Return _FUCTION_ID
		End Get
		Set(ByVal value As String)
			_FUCTION_ID = value
		End Set
	End Property
	Public Property OPERATOR_USER() As String
		Get
			Return _OPERATOR_USER
		End Get
		Set(ByVal value As String)
			_OPERATOR_USER = value
		End Set
	End Property
	Public Property COMMENTS() As String
		Get
			Return _COMMENTS
		End Get
		Set(ByVal value As String)
			_COMMENTS = value
		End Set
	End Property

	'物件建立時執行的事件
	Public Sub New(ByVal Factory_No As String,
								 ByVal Area_No As String,
								 ByVal Device_No As String,
								 ByVal Unit_ID As String,
								 ByVal Occur_Time As String,
								 ByVal Maintenance_Message As String,
								 ByVal Remove_User As String,
								 ByVal Hist_Time As String,
								 ByVal MAINTENANCE_ID As String,
								 ByVal FUCTION_ID As String,
								 ByVal OPERATOR_USER As String,
								 ByVal COMMENTS As String)
		MyBase.New()
		Try
			_Factory_No = Factory_No
			_Area_No = Area_No
			_Device_No = Device_No
			_Unit_ID = Unit_ID
			_Occur_Time = Occur_Time
			_Maintenance_Message = Maintenance_Message
			_Remove_User = Remove_User
			_Hist_Time = Hist_Time

			_MAINTENANCE_ID = MAINTENANCE_ID
			_FUCTION_ID = FUCTION_ID
			_OPERATOR_USER = OPERATOR_USER
			_COMMENTS = COMMENTS

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
      Dim strSQL As String = WMS_CH_LINE_HISTManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
