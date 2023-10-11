Public Class clsPURXC
  Private ShareName As String = "PURXC"
  Private ShareKey As String = ""
  Private _gid As String
  Private _XC001 As String '進貨單別 
  Private _XC002 As String '進貨單號 
  Private _XC003 As String '進貨序號 
  Private _XC004 As String '品號 
  Private _XC005 As Decimal '進貨數量 
  Private _XC006 As String '單位 
  Private _XC007 As String '庫別代號 
  Private _XC008 As String '採購單別 
  Private _XC009 As String '採購單號 
  Private _XC010 As String '採購序號 
  Private _XC011 As Decimal '驗收數量 
  Private _XC012 As Decimal '計價數量 
  Private _XC013 As Decimal '驗退數量 
  Private _XC014 As String '更新碼 
  Private _XC015 As String '採購入庫單號 
  Private _XC016 As String 'WMS進貨序號 
  Private _XC017 As String '建立者 


  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property XC001() As String
    Get
      Return _XC001
    End Get
    Set(ByVal value As String)
      _XC001 = value
    End Set
  End Property
  Public Property XC002() As String
    Get
      Return _XC002
    End Get
    Set(ByVal value As String)
      _XC002 = value
    End Set
  End Property
  Public Property XC003() As String
    Get
      Return _XC003
    End Get
    Set(ByVal value As String)
      _XC003 = value
    End Set
  End Property
  Public Property XC004() As String
    Get
      Return _XC004
    End Get
    Set(ByVal value As String)
      _XC004 = value
    End Set
  End Property
  Public Property XC005() As Decimal
    Get
      Return _XC005
    End Get
    Set(ByVal value As Decimal)
      _XC005 = value
    End Set
  End Property
  Public Property XC006() As String
    Get
      Return _XC006
    End Get
    Set(ByVal value As String)
      _XC006 = value
    End Set
  End Property
  Public Property XC007() As String
    Get
      Return _XC007
    End Get
    Set(ByVal value As String)
      _XC007 = value
    End Set
  End Property
  Public Property XC008() As String
    Get
      Return _XC008
    End Get
    Set(ByVal value As String)
      _XC008 = value
    End Set
  End Property
  Public Property XC009() As String
    Get
      Return _XC009
    End Get
    Set(ByVal value As String)
      _XC009 = value
    End Set
  End Property
  Public Property XC010() As String
    Get
      Return _XC010
    End Get
    Set(ByVal value As String)
      _XC010 = value
    End Set
  End Property
  Public Property XC011() As Decimal
    Get
      Return _XC011
    End Get
    Set(ByVal value As Decimal)
      _XC011 = value
    End Set
  End Property
  Public Property XC012() As Decimal
    Get
      Return _XC012
    End Get
    Set(ByVal value As Decimal)
      _XC012 = value
    End Set
  End Property
  Public Property XC013() As Decimal
    Get
      Return _XC013
    End Get
    Set(ByVal value As Decimal)
      _XC013 = value
    End Set
  End Property
  Public Property XC014() As String
    Get
      Return _XC014
    End Get
    Set(ByVal value As String)
      _XC014 = value
    End Set
  End Property
  Public Property XC015() As String
    Get
      Return _XC015
    End Get
    Set(ByVal value As String)
      _XC015 = value
    End Set
  End Property
  Public Property XC016() As String
    Get
      Return _XC016
    End Get
    Set(ByVal value As String)
      _XC016 = value
    End Set
  End Property
  Public Property XC017() As String
    Get
      Return _XC017
    End Get
    Set(ByVal value As String)
      _XC017 = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal XC001 As String, ByVal XC002 As String, ByVal XC003 As String, ByVal XC004 As String, ByVal XC005 As Decimal, ByVal XC006 As String, ByVal XC007 As String, ByVal XC008 As String, ByVal XC009 As String, ByVal XC010 As String, ByVal XC011 As Decimal, ByVal XC012 As Decimal, ByVal XC013 As Decimal, ByVal XC014 As String, ByVal XC015 As String, ByVal XC016 As String, ByVal XC017 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(XC008, XC009, XC010, XC016)
      _gid = key
      _XC001 = XC001
      _XC002 = XC002
      _XC003 = XC003
      _XC004 = XC004
      _XC005 = XC005
      _XC006 = XC006
      _XC007 = XC007
      _XC008 = XC008
      _XC009 = XC009
      _XC010 = XC010
      _XC011 = XC011
      _XC012 = XC012
      _XC013 = XC013
      _XC014 = XC014
      _XC015 = XC015
      _XC016 = XC016
      _XC017 = XC017
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
  Public Shared Function Get_Combination_Key(ByVal XC008 As String, ByVal XC009 As String, ByVal XC010 As String, ByVal XC016 As String) As String
    Try
      Dim key As String = XC008 & LinkKey & XC009 & LinkKey & XC010 & XC016
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPURXC
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
      Dim strSQL As String = PURXCManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = PURXCManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = PURXCManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
