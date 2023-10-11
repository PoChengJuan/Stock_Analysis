Public Class clsLineProduction_Info
  Private ShareName As String = "LineProduction_Info"
  Private ShareKey As String = ""

  Private _gid As String
  Private _Factory_No As String
  Private _Area_No As String
  Private _Device_No As String
  Private _Unit_ID As String
  Private _Qty_Process As Double
  Private _Previous_Qty_Process As Double
  Private _Reset_Qty_Process As Double
  Private _Qty_Modify As Double
  Private _Previous_Qty_Modify As Double
  Private _Reset_Qty_Modify As Double
  Private _Qty_NG As Double
  Private _Previous_Qty_NG As Double
  Private _Reset_Qty_NG As Double
	Private _Update_Time As String
	Private _QTY_TOTAL As Double

	Private _objHandling As clsHandlingObject
  Private _objLine_Area As clsLine_Area

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
  Public Property Qty_Process() As Double
    Get
      Return _Qty_Process
    End Get
    Set(ByVal value As Double)
      _Qty_Process = value
    End Set
  End Property
  Public Property Previous_Qty_Process() As Double
    Get
      Return _Previous_Qty_Process
    End Get
    Set(ByVal value As Double)
      _Previous_Qty_Process = value
    End Set
  End Property
  Public Property Reset_Qty_Process() As Double
    Get
      Return _Reset_Qty_Process
    End Get
    Set(ByVal value As Double)
      _Reset_Qty_Process = value
    End Set
  End Property
  Public Property Qty_Modify() As Double
    Get
      Return _Qty_Modify
    End Get
    Set(ByVal value As Double)
      _Qty_Modify = value
    End Set
  End Property
  Public Property Previous_Qty_Modify() As Double
    Get
      Return _Previous_Qty_Modify
    End Get
    Set(ByVal value As Double)
      _Previous_Qty_Modify = value
    End Set
  End Property
  Public Property Reset_Qty_Modify() As Double
    Get
      Return _Reset_Qty_Modify
    End Get
    Set(ByVal value As Double)
      _Reset_Qty_Modify = value
    End Set
  End Property
  Public Property Qty_NG() As Double
    Get
      Return _Qty_NG
    End Get
    Set(ByVal value As Double)
      _Qty_NG = value
    End Set
  End Property
  Public Property Previous_Qty_NG() As Double
    Get
      Return _Previous_Qty_NG
    End Get
    Set(ByVal value As Double)
      _Previous_Qty_NG = value
    End Set
  End Property
  Public Property Reset_Qty_NG() As Double
    Get
      Return _Reset_Qty_NG
    End Get
    Set(ByVal value As Double)
      _Reset_Qty_NG = value
    End Set
  End Property
  Public Property Update_Time() As String
    Get
      Return _Update_Time
    End Get
    Set(ByVal value As String)
      _Update_Time = value
    End Set
  End Property
  Public Property objLine_Area() As clsLine_Area
    Get
      Return _objLine_Area
    End Get
    Set(ByVal value As clsLine_Area)
      _objLine_Area = value
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
	Public Property QTY_TOTAL() As Double
		Get
			Return _QTY_TOTAL
		End Get
		Set(ByVal value As Double)
			_QTY_TOTAL = value
		End Set
	End Property

  '物件建立時執行的事件
  Public Sub New(ByVal Factory_No As String,
                 ByVal Area_No As String,
                 ByVal Device_No As String,
                 ByVal Unit_ID As String,
                 ByVal Qty_Process As Double,
                 ByVal Previous_Qty_Process As Double,
                 ByVal Reset_Qty_Process As Double,
                 ByVal Qty_Modify As Double,
                 ByVal Previous_Qty_Modify As Double,
                 ByVal Reset_Qty_Modify As Double,
                 ByVal Qty_NG As Double,
                 ByVal Previous_Qty_NG As Double,
                 ByVal Reset_Qty_NG As Double,
                 ByVal Update_Time As String,
                 ByVal QTY_TOTAL As Double)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(Factory_No, Area_No, Device_No, Unit_ID)
      _gid = key
      _Factory_No = Factory_No
      _Area_No = Area_No
      _Device_No = Device_No
      _Unit_ID = Unit_ID
      _Qty_Process = Qty_Process
      _Previous_Qty_Process = Previous_Qty_Process
      _Reset_Qty_Process = Reset_Qty_Process
      _Qty_Modify = Qty_Modify
      _Previous_Qty_Modify = Previous_Qty_Modify
      _Reset_Qty_Modify = Reset_Qty_Modify
      _Qty_NG = Qty_NG
      _Previous_Qty_NG = Previous_Qty_NG
      _Reset_Qty_NG = Reset_Qty_NG
      _Update_Time = Update_Time
      _QTY_TOTAL = QTY_TOTAL


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
    _objLine_Area = Nothing
  End Sub

  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal Factory_No As String,
                                             ByVal Area_No As String,
                                             ByVal Device_No As String,
                                             ByVal Unit_ID As String) As String
    Try
      Dim key As String = Factory_No & LinkKey & Area_No & LinkKey & Device_No & LinkKey & Unit_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsLineProduction_Info
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
        objHandling.O_Add_LineProduction_Info(Me)
        Dim objLine_Area As clsLine_Area = Nothing
        If objHandling.O_Get_Line_Area(Factory_No, Area_No, objLine_Area) Then
          If objLine_Area IsNot Nothing Then
            _objLine_Area = objLine_Area
            objLine_Area.O_Add_LineProduction_Info(Me)
          End If
        End If
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_CLineProduction(Me)
      End If
      If _objLine_Area IsNot Nothing Then
        _objLine_Area.O_Remove_LineProduction_Info(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_CT_LINE_PRODUCTION_INFOManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_CT_LINE_PRODUCTION_INFOManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_CT_LINE_PRODUCTION_INFOManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '非標準的Function
  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsLineProduction_Info) As Boolean
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
      _Qty_Process = obj.Qty_Process
      _Previous_Qty_Process = obj.Previous_Qty_Process
      _Reset_Qty_Process = obj.Reset_Qty_Process
      _Qty_Modify = obj.Qty_Modify
      _Previous_Qty_Modify = obj.Previous_Qty_Modify
      _Reset_Qty_Modify = obj.Reset_Qty_Modify
      _Qty_NG = obj.Qty_NG
      _Previous_Qty_NG = obj.Previous_Qty_NG
      _Reset_Qty_NG = obj.Reset_Qty_NG
      _Update_Time = obj.Update_Time
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
