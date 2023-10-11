Public Class clsMAINTENANCE_DTL
	Private ShareName As String = "DTL"
	Private ShareKey As String = ""
	Private _gid As String
	Private _FACTORY_NO As String '廠別

	Private _DEVICE_NO As String 'LCS的編號

	Private _AREA_NO As String 'Area編號

  Private _UNIT_ID As String '設備編號

  Private _MAINTENANCE_ID As String '設備中的哪一個Maintenance的保養設定

  Private _FUNCTION_ID As String '設備中的哪一個Function的保養設定

  Private _VALUE_TYPE As Double '輸入資料的類型(1:數值、2:日期...)

  Private _NOTICE_TYPE As Double '警告模式(0:超過高低水位、1:低於低水位、2:高於高水位)

  Private _HIGH_WATER_VALUE As String '高標

  Private _LOW_WATER_VALUE As String '低標

  Private _STANDARD_VALUE As String '標準值(暫時用不到)

  Private _VALUE_RANGE As String '警告值可容許範圍

  Private _MAINTENANCE_MESSAGE As String '警告訊息

  Private _VALUE_SOURCE As Double '保養資料來源(暫時不用)

  Private _VALUE_UPDATE_TYPE As Double '資料更新方式(1:更新成新值/2:累加)

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
  Public Property VALUE_TYPE() As Double
    Get
      Return _VALUE_TYPE
    End Get
    Set(ByVal value As Double)
      _VALUE_TYPE = value
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
  Public Property HIGH_WATER_VALUE() As String
    Get
      Return _HIGH_WATER_VALUE
    End Get
    Set(ByVal value As String)
      _HIGH_WATER_VALUE = value
    End Set
  End Property
  Public Property LOW_WATER_VALUE() As String
    Get
      Return _LOW_WATER_VALUE
    End Get
    Set(ByVal value As String)
      _LOW_WATER_VALUE = value
    End Set
  End Property
  Public Property STANDARD_VALUE() As String
    Get
      Return _STANDARD_VALUE
    End Get
    Set(ByVal value As String)
      _STANDARD_VALUE = value
    End Set
  End Property
  Public Property VALUE_RANGE() As String
    Get
      Return _VALUE_RANGE
    End Get
    Set(ByVal value As String)
      _VALUE_RANGE = value
    End Set
  End Property
  Public Property MAINTENANCE_MESSAGE() As String
    Get
      Return _MAINTENANCE_MESSAGE
    End Get
    Set(ByVal value As String)
      _MAINTENANCE_MESSAGE = value
    End Set
  End Property
  Public Property VALUE_SOURCE() As Double
    Get
      Return _VALUE_SOURCE
    End Get
    Set(ByVal value As Double)
      _VALUE_SOURCE = value
    End Set
  End Property
  Public Property VALUE_UPDATE_TYPE() As Double
    Get
      Return _VALUE_UPDATE_TYPE
    End Get
    Set(ByVal value As Double)
      _VALUE_UPDATE_TYPE = value
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

  Public Sub New(ByVal FACTORY_NO As String, ByVal DEVICE_NO As String, ByVal AREA_NO As String, ByVal UNIT_ID As String, ByVal MAINTENANCE_ID As String, ByVal FUNCTION_ID As String, ByVal VALUE_TYPE As Double, ByVal NOTICE_TYPE As Double, ByVal HIGH_WATER_VALUE As String, ByVal LOW_WATER_VALUE As String, ByVal STANDARD_VALUE As String, ByVal VALUE_RANGE As String, ByVal MAINTENANCE_MESSAGE As String, ByVal VALUE_SOURCE As Double, ByVal VALUE_UPDATE_TYPE As Double)
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
      _VALUE_TYPE = VALUE_TYPE
      _NOTICE_TYPE = NOTICE_TYPE
      _HIGH_WATER_VALUE = HIGH_WATER_VALUE
      _LOW_WATER_VALUE = LOW_WATER_VALUE
      _STANDARD_VALUE = STANDARD_VALUE
      _VALUE_RANGE = VALUE_RANGE
      _MAINTENANCE_MESSAGE = MAINTENANCE_MESSAGE
      _VALUE_SOURCE = VALUE_SOURCE
      _VALUE_UPDATE_TYPE = VALUE_UPDATE_TYPE
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
  Public Shared Function Get_Combination_Key(ByVal FACTORY_NO As String, ByVal DEVICE_NO As String, ByVal AREA_NO As String, ByVal UNIT_ID As String, ByVal MAINTENANCE_ID As String, ByVal FUNCTION_ID As String) As String
    Try
      Dim key As String = FACTORY_NO & LinkKey & DEVICE_NO & LinkKey & AREA_NO & LinkKey & UNIT_ID & LinkKey & MAINTENANCE_ID & LinkKey & FUNCTION_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMAINTENANCE_DTL
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
        objHandling.O_Add_MaintenanceDTL(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_MaintenanceDTL(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_M_MAINTENANCE_DTLManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_MAINTENANCE_DTLManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_MAINTENANCE_DTLManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_M_MAINTENANCE_DTL As clsMAINTENANCE_DTL) As Boolean
    Try
      Dim key As String = objWMS_M_MAINTENANCE_DTL._gid
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
			_VALUE_TYPE = VALUE_TYPE
			_NOTICE_TYPE = NOTICE_TYPE
			_HIGH_WATER_VALUE = HIGH_WATER_VALUE
			_LOW_WATER_VALUE = LOW_WATER_VALUE
			_STANDARD_VALUE = STANDARD_VALUE
			_VALUE_RANGE = VALUE_RANGE
			_MAINTENANCE_MESSAGE = MAINTENANCE_MESSAGE
			_VALUE_SOURCE = VALUE_SOURCE
			_VALUE_UPDATE_TYPE = VALUE_UPDATE_TYPE
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
