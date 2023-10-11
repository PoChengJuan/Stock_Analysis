Public Class clsDATA_REPORT_SET
	Private ShareName As String = "SET"
	Private ShareKey As String = ""
	Private _gid As String
	Private _ROLE_ID As String '底層定義的Alarm RoleID

	Private _ROLE_TYPE As Double '設定的類型(1:異常設定/2:保養參數設定)

	Private _FUNCTION_ID As String 'BigData定義的Function_ID

	Private _FUNCTION_NAME As String 'BigData定義的Function_Name

	Private _DEVICE_NO As String 'LCS的編號

	Private _AREA_NO As String 'Area編號

	Private _UNIT_ID As String '設備編號

	Private _HIGH_WATER_VALUE As Long '高標

	Private _LOW_WATER_VALUE As Long '低標

	Private _STANDARD_VALUE As Long '標準值

	Private _VALUE_RANGE As Long '警告值可容許範圍

	Private _NOTICE_TYPE As Double '警告模式(0:超過高低水位、1:低於低水位、2:高於高水位)

	Private _CONTINUE_SEND As Double '是否持續發送(0:只發送一次、1:間隔一定時間發送一次)

	Private _SEND_INTERVAL As Double '發送間隔時間(S)

	Private _ENABLE As Boolean '是否啟用0:禁用1:啟用

	Public Property gid() As String
		Get
			Return _gid
		End Get
		Set(ByVal value As String)
			_gid = value
		End Set
	End Property
	Public Property ROLE_ID() As String
		Get
			Return _ROLE_ID
		End Get
		Set(ByVal value As String)
			_ROLE_ID = value
		End Set
	End Property
	Public Property ROLE_TYPE() As Double
		Get
			Return _ROLE_TYPE
		End Get
		Set(ByVal value As Double)
			_ROLE_TYPE = value
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
	Public Property FUNCTION_NAME() As String
		Get
			Return _FUNCTION_NAME
		End Get
		Set(ByVal value As String)
			_FUNCTION_NAME = value
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
	Public Property HIGH_WATER_VALUE() As Double
		Get
			Return _HIGH_WATER_VALUE
		End Get
		Set(ByVal value As Double)
			_HIGH_WATER_VALUE = value
		End Set
	End Property
	Public Property LOW_WATER_VALUE() As Double
		Get
			Return _LOW_WATER_VALUE
		End Get
		Set(ByVal value As Double)
			_LOW_WATER_VALUE = value
		End Set
	End Property
	Public Property STANDARD_VALUE() As Double
		Get
			Return _STANDARD_VALUE
		End Get
		Set(ByVal value As Double)
			_STANDARD_VALUE = value
		End Set
	End Property
	Public Property VALUE_RANGE() As Double
		Get
			Return _VALUE_RANGE
		End Get
		Set(ByVal value As Double)
			_VALUE_RANGE = value
		End Set
	End Property
	Public Property NOTICE_TYPE() As Double
		Get
			Return _NOTICE_TYPE
		End Get
		Set(ByVal value As Double)
			_NOTICE_TYPE = value
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
	Public Property ENABLE() As Double
		Get
			Return _ENABLE
		End Get
		Set(ByVal value As Double)
			_ENABLE = value
		End Set
	End Property

	Public Sub New(ByVal ROLE_ID As String, ByVal ROLE_TYPE As Double, ByVal FUNCTION_ID As String, ByVal FUNCTION_NAME As String, ByVal DEVICE_NO As String, ByVal AREA_NO As String, ByVal UNIT_NO As String, ByVal HIGH_WATER_VALUE As Long, ByVal LOW_WATER_VALUE As Long, ByVal STANDARD_VALUE As Long, ByVal VALUE_RANGE As Long, ByVal NOTICE_TYPE As Double, ByVal CONTINUE_SEND As Double, ByVal SEND_INTERVAL As Double, ByVal ENABLE As Boolean)
		MyBase.New()
		Try
			Dim key As String = Get_Combination_Key(ROLE_ID, FUNCTION_ID, DEVICE_NO, AREA_NO, UNIT_NO)
			_gid = key
			_ROLE_ID = ROLE_ID
			_ROLE_TYPE = ROLE_TYPE
			_FUNCTION_ID = FUNCTION_ID
			_FUNCTION_NAME = FUNCTION_NAME
			_DEVICE_NO = DEVICE_NO
			_AREA_NO = AREA_NO
			_UNIT_ID = UNIT_NO
			_HIGH_WATER_VALUE = HIGH_WATER_VALUE
			_LOW_WATER_VALUE = LOW_WATER_VALUE
			_STANDARD_VALUE = STANDARD_VALUE
			_VALUE_RANGE = VALUE_RANGE
			_NOTICE_TYPE = NOTICE_TYPE
			_CONTINUE_SEND = CONTINUE_SEND
			_SEND_INTERVAL = SEND_INTERVAL
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
	Public Shared Function Get_Combination_Key(ByVal ROLE_ID As String, ByVal FUNCTION_ID As String, ByVal DEVICE_NO As String, ByVal AREA_NO As String, ByVal UNIT_NO As String) As String
		Try
			Dim key As String = ROLE_ID & LinkKey & FUNCTION_ID & LinkKey & DEVICE_NO & LinkKey & AREA_NO & LinkKey & UNIT_NO

			Return key
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return ""
		End Try
	End Function
	Public Function Clone() As clsDATA_REPORT_SET
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
			Dim strSQL As String = WMS_M_DATA_REPORT_SETManagement.GetInsertSQL(Me)
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
			Dim strSQL As String = WMS_M_DATA_REPORT_SETManagement.GetUpdateSQL(Me)
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
			Dim strSQL As String = WMS_M_DATA_REPORT_SETManagement.GetDeleteSQL(Me)
			lstSQL.Add(strSQL)
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
	Public Function Update_To_Memory(ByRef objWMS_M_DATA_REPORT_SET As clsDATA_REPORT_SET) As Boolean
		Try
			Dim key As String = objWMS_M_DATA_REPORT_SET._gid
			If key <> _gid Then
				SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
				Return False
			End If
			_ROLE_ID = ROLE_ID
			_ROLE_TYPE = ROLE_TYPE
			_FUNCTION_ID = FUNCTION_ID
			_FUNCTION_NAME = FUNCTION_NAME
			_DEVICE_NO = DEVICE_NO
			_AREA_NO = AREA_NO
			_UNIT_ID = UNIT_ID
			_HIGH_WATER_VALUE = HIGH_WATER_VALUE
			_LOW_WATER_VALUE = LOW_WATER_VALUE
			_STANDARD_VALUE = STANDARD_VALUE
			_VALUE_RANGE = VALUE_RANGE
			_NOTICE_TYPE = NOTICE_TYPE
			_CONTINUE_SEND = CONTINUE_SEND
			_SEND_INTERVAL = SEND_INTERVAL
			_ENABLE = ENABLE
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
