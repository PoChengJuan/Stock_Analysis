Public Class clsPURTD
  Private ShareName As String = "PURTD"
  Private ShareKey As String = ""
  Private _gid As String
  Private _TD001 As String '採購單別 
  Private _TD002 As String '採購單號 
  Private _TD003 As String '序號 
  Private _TD004 As String '品號 
  Private _TD005 As String '品名 
  Private _TD006 As String '規格 
  Private _TD007 As String '交貨庫別 
  Private _TD008 As Decimal '採購數量 
  Private _TD009 As String '單位 
  Private _TD010 As Decimal '採購單價 
  Private _TD011 As Decimal '採購金額 
  Private _TD012 As String '預交日 
  Private _TD021 As String  '若TD024沒填，則改用21+22+23的字串當計畫批號
  Private _TD022 As String
  Private _TD023 As String
  Private _TD024 As String '計畫批號



  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property TD001() As String
    Get
      Return _TD001
    End Get
    Set(ByVal value As String)
      _TD001 = value
    End Set
  End Property
  Public Property TD002() As String
    Get
      Return _TD002
    End Get
    Set(ByVal value As String)
      _TD002 = value
    End Set
  End Property
  Public Property TD003() As String
    Get
      Return _TD003
    End Get
    Set(ByVal value As String)
      _TD003 = value
    End Set
  End Property
  Public Property TD004() As String
    Get
      Return _TD004
    End Get
    Set(ByVal value As String)
      _TD004 = value
    End Set
  End Property
  Public Property TD005() As String
    Get
      Return _TD005
    End Get
    Set(ByVal value As String)
      _TD005 = value
    End Set
  End Property
  Public Property TD006() As String
    Get
      Return _TD006
    End Get
    Set(ByVal value As String)
      _TD006 = value
    End Set
  End Property
  Public Property TD007() As String
    Get
      Return _TD007
    End Get
    Set(ByVal value As String)
      _TD007 = value
    End Set
  End Property
  Public Property TD008() As Decimal
    Get
      Return _TD008
    End Get
    Set(ByVal value As Decimal)
      _TD008 = value
    End Set
  End Property
  Public Property TD009() As String
    Get
      Return _TD009
    End Get
    Set(ByVal value As String)
      _TD009 = value
    End Set
  End Property
  Public Property TD010() As Decimal
    Get
      Return _TD010
    End Get
    Set(ByVal value As Decimal)
      _TD010 = value
    End Set
  End Property
  Public Property TD011() As Decimal
    Get
      Return _TD011
    End Get
    Set(ByVal value As Decimal)
      _TD011 = value
    End Set
  End Property
  Public Property TD012() As String
    Get
      Return _TD012
    End Get
    Set(ByVal value As String)
      _TD012 = value
    End Set
  End Property
  Public Property TD021() As String
    Get
      Return _TD021
    End Get
    Set(ByVal value As String)
      _TD021 = value
    End Set
  End Property
  Public Property TD022() As String
    Get
      Return _TD022
    End Get
    Set(ByVal value As String)
      _TD022 = value
    End Set
  End Property
  Public Property TD023() As String
    Get
      Return _TD023
    End Get
    Set(ByVal value As String)
      _TD023 = value
    End Set
  End Property
  Public Property TD024() As String
    Get
      Return _TD024
    End Get
    Set(ByVal value As String)
      _TD024 = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal TD001 As String, ByVal TD002 As String, ByVal TD003 As String, ByVal TD004 As String, ByVal TD005 As String, ByVal TD006 As String, ByVal TD007 As String, ByVal TD008 As Decimal, ByVal TD009 As String, ByVal TD010 As Decimal, ByVal TD011 As Decimal, ByVal TD012 As String, ByVal TD021 As String, ByVal TD022 As String, ByVal TD023 As String, ByVal TD024 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(TD001, TD002, TD003)
      _gid = key
      _TD001 = TD001
      _TD002 = TD002
      _TD003 = TD003
      _TD004 = TD004
      _TD005 = TD005
      _TD006 = TD006
      _TD007 = TD007
      _TD008 = TD008
      _TD009 = TD009
      _TD010 = TD010
      _TD011 = TD011
      _TD012 = TD012
      _TD021 = TD021
      _TD022 = TD022
      _TD023 = TD023
      _TD024 = TD024
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
  Public Shared Function Get_Combination_Key(ByVal TD001 As String, ByVal TD002 As String, ByVal TD003 As String) As String
    Try
      Dim key As String = TD001 & LinkKey & TD002 & LinkKey & TD003
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPURTD
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
      Dim strSQL As String = PURTDManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = PURTDManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = PURTDManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
