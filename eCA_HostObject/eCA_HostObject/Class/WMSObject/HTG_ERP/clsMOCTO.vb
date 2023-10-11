Public Class clsMOCTO
  Private ShareName As String = "MOCTO"
  Private ShareKey As String = ""
  Private _gid As String
  Private _TO001 As String '製令單別 
  Private _TO002 As String '製令單號 
  Private _TO003 As String '變更版次 
  Private _TO004 As String '變更日期 
  Private _TO005 As String '變更原因 
  Private _TO006 As String '新開單日期 
  Private _TO007 As String '新BOM日期 
  Private _TO008 As String '新BOM版次 
  Private _TO009 As String '新產品品號 
  Private _TO010 As String '新單位 
  Private _TO017 As String '新數量
  Private _TO034 As String '新計畫批號


  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property TO001() As String
    Get
      Return _TO001
    End Get
    Set(ByVal value As String)
      _TO001 = value
    End Set
  End Property
  Public Property TO002() As String
    Get
      Return _TO002
    End Get
    Set(ByVal value As String)
      _TO002 = value
    End Set
  End Property
  Public Property TO003() As String
    Get
      Return _TO003
    End Get
    Set(ByVal value As String)
      _TO003 = value
    End Set
  End Property
  Public Property TO004() As String
    Get
      Return _TO004
    End Get
    Set(ByVal value As String)
      _TO004 = value
    End Set
  End Property
  Public Property TO005() As String
    Get
      Return _TO005
    End Get
    Set(ByVal value As String)
      _TO005 = value
    End Set
  End Property
  Public Property TO006() As String
    Get
      Return _TO006
    End Get
    Set(ByVal value As String)
      _TO006 = value
    End Set
  End Property
  Public Property TO007() As String
    Get
      Return _TO007
    End Get
    Set(ByVal value As String)
      _TO007 = value
    End Set
  End Property
  Public Property TO008() As String
    Get
      Return _TO008
    End Get
    Set(ByVal value As String)
      _TO008 = value
    End Set
  End Property
  Public Property TO009() As String
    Get
      Return _TO009
    End Get
    Set(ByVal value As String)
      _TO009 = value
    End Set
  End Property
  Public Property TO010() As String
    Get
      Return _TO010
    End Get
    Set(ByVal value As String)
      _TO010 = value
    End Set
  End Property
  Public Property TO017() As String
    Get
      Return _TO017
    End Get
    Set(ByVal value As String)
      _TO017 = value
    End Set
  End Property
  Public Property TO034() As String '新計畫批號
    Get
      Return _TO034
    End Get
    Set(ByVal value As String) '新計畫批號
      _TO034 = value
    End Set
  End Property



  '物件建立時執行的事件
  Public Sub New(ByVal TO001 As String, ByVal TO002 As String, ByVal TO003 As String, ByVal TO004 As String, ByVal TO005 As String, ByVal TO006 As String, ByVal TO007 As String, ByVal TO008 As String, ByVal TO009 As String, ByVal TO010 As String, ByVal TO017 As Decimal, ByVal TO034 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(TO001, TO002, TO003)
      _gid = key
      _TO001 = TO001
      _TO002 = TO002
      _TO003 = TO003
      _TO004 = TO004
      _TO005 = TO005
      _TO006 = TO006
      _TO007 = TO007
      _TO008 = TO008
      _TO009 = TO009
      _TO010 = TO010
      _TO017 = TO017
      _TO034 = TO034
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

  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal TO001 As String, ByVal TO002 As String, ByVal TO003 As String) As String
    Try
      Dim key As String = TO001 & LinkKey & TO002 & LinkKey & TO003
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMOCTO
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
      Dim strSQL As String = MOCTOManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = MOCTOManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = MOCTOManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
