Public Class clsLineInfo
  Private ShareName As String = "LineInfo"
  Private ShareKey As String = ""

  Private _gid As String
  Private _Factory_No As String
  Private _Area_No As String
  Private _Device_No As String
  Private _Unit_ID As String
  Private _Occur_Time As String
	Private _Maintenance_Message As String
	Private _MAINTENANCE_ID As String
	Private _FUCTION_ID As String



	Private _objHandling As clsHandlingObject

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
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


	Public Property objHandling() As clsHandlingObject
    Get
      Return _objHandling
    End Get
    Set(ByVal value As clsHandlingObject)
      _objHandling = value
    End Set
  End Property

	'物件建立時執行的事件
	Public Sub New(ByVal Factory_No As String,
								 ByVal Area_No As String,
								 ByVal Device_No As String,
								 ByVal Unit_ID As String,
								 ByVal Occur_Time As String,
								 ByVal Maintenance_Message As String,
								 ByVal MAINTENANCE_ID As String,
								 ByVal FUCTION_ID As String)
		MyBase.New()
		Try
			Dim key As String = Get_Combination_Key(Factory_No, Area_No, Device_No, Unit_ID, MAINTENANCE_ID, FUCTION_ID)
			_gid = key
			_Factory_No = Factory_No
			_Area_No = Area_No
			_Device_No = Device_No
			_Unit_ID = Unit_ID
			_Occur_Time = Occur_Time
			_Maintenance_Message = Maintenance_Message
			_MAINTENANCE_ID = MAINTENANCE_ID
			_FUCTION_ID = FUCTION_ID

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
    _objHandling = Nothing
  End Sub

	'=================Public Function=======================
	'傳入指定參數取得Key值
	Public Shared Function Get_Combination_Key(ByVal Factory_No As String,
																						 ByVal Area_No As String,
																						 ByVal Device_No As String,
																						 ByVal Unit_ID As String,
																						 ByVal MAINTENANCE_ID As String,
																						 ByVal FUCTION_ID As String) As String
		Try
			Dim key As String = Factory_No & LinkKey & Area_No & LinkKey & Device_No & LinkKey & Unit_ID & LinkKey & MAINTENANCE_ID & LinkKey & FUCTION_ID
			Return key
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return ""
		End Try
	End Function
	Public Function Clone() As clsLineInfo
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
        objHandling.O_Add_LineInfo(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_LineInfo(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_CT_LINE_INFOManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_CT_LINE_INFOManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_CT_LINE_INFOManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsLineInfo) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> gid Then
        SendMessageToLog("Key can not Update, old_Key=" & gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _Factory_No = obj.Factory_No
      _Area_No = obj.Area_No
      _Device_No = obj.Device_No
      _Unit_ID = obj.Unit_ID
      _Occur_Time = obj.Occur_Time
      _Maintenance_Message = obj.Maintenance_Message
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
