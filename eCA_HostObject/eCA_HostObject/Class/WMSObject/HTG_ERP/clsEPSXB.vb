Public Class clsEPSXB
  Private ShareName As String = "EPSXB"
  Private ShareKey As String = ""
  Private _gid As String
  Private _XB001 As String '單別 
  Private _XB002 As String '單號 
  Private _XB003 As String '序號 
  Private _XB004 As String '品號 
  Private _XB005 As Decimal '預計出貨數量 
  Private _XB006 As Decimal '實際出貨數量 
  Private _XB007 As String '訂單單別 
  Private _XB008 As String '訂單單號 
  Private _XB009 As String '訂單序號 
  Private _XB010 As String '更新碼 
  Private _XB011 As String '出通單日期 
  Private _XB012 As String '建立者 
  Private _XB013 As String '確認者 
  Private _XB014 As String '首拋時間 
  Private _XB015 As String '更新時間 


  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property XB001() As String
    Get
      Return _XB001
    End Get
    Set(ByVal value As String)
      _XB001 = value
    End Set
  End Property
  Public Property XB002() As String
    Get
      Return _XB002
    End Get
    Set(ByVal value As String)
      _XB002 = value
    End Set
  End Property
  Public Property XB003() As String
    Get
      Return _XB003
    End Get
    Set(ByVal value As String)
      _XB003 = value
    End Set
  End Property
  Public Property XB004() As String
    Get
      Return _XB004
    End Get
    Set(ByVal value As String)
      _XB004 = value
    End Set
  End Property
  Public Property XB005() As Decimal
    Get
      Return _XB005
    End Get
    Set(ByVal value As Decimal)
      _XB005 = value
    End Set
  End Property
  Public Property XB006() As Decimal
    Get
      Return _XB006
    End Get
    Set(ByVal value As Decimal)
      _XB006 = value
    End Set
  End Property
  Public Property XB007() As String
    Get
      Return _XB007
    End Get
    Set(ByVal value As String)
      _XB007 = value
    End Set
  End Property
  Public Property XB008() As String
    Get
      Return _XB008
    End Get
    Set(ByVal value As String)
      _XB008 = value
    End Set
  End Property
  Public Property XB009() As String
    Get
      Return _XB009
    End Get
    Set(ByVal value As String)
      _XB009 = value
    End Set
  End Property
  Public Property XB010() As String
    Get
      Return _XB010
    End Get
    Set(ByVal value As String)
      _XB010 = value
    End Set
  End Property
  Public Property XB011() As String
    Get
      Return _XB011
    End Get
    Set(ByVal value As String)
      _XB011 = value
    End Set
  End Property
  Public Property XB012() As String
    Get
      Return _XB012
    End Get
    Set(ByVal value As String)
      _XB012 = value
    End Set
  End Property
  Public Property XB013() As String
    Get
      Return _XB013
    End Get
    Set(ByVal value As String)
      _XB013 = value
    End Set
  End Property
  Public Property XB014() As String
    Get
      Return _XB014
    End Get
    Set(ByVal value As String)
      _XB014 = value
    End Set
  End Property
  Public Property XB015() As String
    Get
      Return _XB015
    End Get
    Set(ByVal value As String)
      _XB015 = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal XB001 As String, ByVal XB002 As String, ByVal XB003 As String, ByVal XB004 As String, ByVal XB005 As Decimal, ByVal XB006 As Decimal, ByVal XB007 As String, ByVal XB008 As String, ByVal XB009 As String, ByVal XB010 As String, ByVal XB011 As String, ByVal XB012 As String, ByVal XB013 As String, ByVal XB014 As String, ByVal XB015 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(XB001, XB002, XB003)
      _gid = key
      _XB001 = XB001
      _XB002 = XB002
      _XB003 = XB003
      _XB004 = XB004
      _XB005 = XB005
      _XB006 = XB006
      _XB007 = XB007
      _XB008 = XB008
      _XB009 = XB009
      _XB010 = XB010
      _XB011 = XB011
      _XB012 = XB012
      _XB013 = XB013
      _XB014 = XB014
      _XB015 = XB015
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
  Public Shared Function Get_Combination_Key(ByVal XB001 As String, ByVal XB002 As String, ByVal XB003 As String) As String
    Try
      Dim key As String = XB001 & LinkKey & XB002 & LinkKey & XB003
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsEPSXB
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
      Dim strSQL As String = EPSXBManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = EPSXBManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = EPSXBManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
