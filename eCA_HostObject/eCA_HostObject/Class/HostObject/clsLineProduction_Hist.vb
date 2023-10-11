Public Class clsLineProduction_Hist
  Private ShareName As String = "LineProduction_Hist"
  Private ShareKey As String = ""
  Private _Factory_No As String
  Private _Area_No As String
  Private _Device_No As String
  Private _Unit_ID As String
  Private _Qty_Process As Double
  Private _Qty_Modify As Double
  Private _Qty_NG As Double
  Private _Hist_Time As String
  Private _QTY_TOTAL As Double


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
  Public Property Qty_Modify() As Double
    Get
      Return _Qty_Modify
    End Get
    Set(ByVal value As Double)
      _Qty_Modify = value
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
  Public Property Hist_Time() As String
    Get
      Return _Hist_Time
    End Get
    Set(ByVal value As String)
      _Hist_Time = value
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
                               ByVal Qty_Modify As Double,
                               ByVal Qty_NG As Double,
                               ByVal Hist_Time As String,
                               ByVal QTY_TOTAL As Double)
    MyBase.New()
    Try
      _Factory_No = Factory_No
      _Area_No = Area_No
      _Device_No = Device_No
      _Unit_ID = Unit_ID
      _Qty_Process = Qty_Process
      _Qty_Modify = Qty_Modify
      _Qty_NG = Qty_NG
      _Hist_Time = Hist_Time
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
  End Sub

  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_CH_LINE_PRODUCTION_HISTManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
