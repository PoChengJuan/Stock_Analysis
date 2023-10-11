Public Class clsINVXD
  Private ShareName As String = "INVXD"
  Private ShareKey As String = ""
  Private _gid As String
  Private _XD001 As String '單別 
  Private _XD002 As String '單號 
  Private _XD003 As String '單據性質碼 
  Private _XD004 As String '序號 
  Private _XD005 As String '品號 
  Private _XD006 As Decimal '數量 
  Private _XD007 As String '轉出庫 
  Private _XD008 As String '轉入庫 
  Private _XD009 As String '更新碼 
  Private _XD010 As String '公司別 
  Private _XD011 As String '建立者 
  Private _XD012 As String '確認者 
  Private _XD013 As String '已完成數量



  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property XD001() As String
    Get
      Return _XD001
    End Get
    Set(ByVal value As String)
      _XD001 = value
    End Set
  End Property
  Public Property XD002() As String
    Get
      Return _XD002
    End Get
    Set(ByVal value As String)
      _XD002 = value
    End Set
  End Property
  Public Property XD003() As String
    Get
      Return _XD003
    End Get
    Set(ByVal value As String)
      _XD003 = value
    End Set
  End Property
  Public Property XD004() As String
    Get
      Return _XD004
    End Get
    Set(ByVal value As String)
      _XD004 = value
    End Set
  End Property
  Public Property XD005() As String
    Get
      Return _XD005
    End Get
    Set(ByVal value As String)
      _XD005 = value
    End Set
  End Property
  Public Property XD006() As Decimal
    Get
      Return _XD006
    End Get
    Set(ByVal value As Decimal)
      _XD006 = value
    End Set
  End Property
  Public Property XD007() As String
    Get
      Return _XD007
    End Get
    Set(ByVal value As String)
      _XD007 = value
    End Set
  End Property
  Public Property XD008() As String
    Get
      Return _XD008
    End Get
    Set(ByVal value As String)
      _XD008 = value
    End Set
  End Property
  Public Property XD009() As String
    Get
      Return _XD009
    End Get
    Set(ByVal value As String)
      _XD009 = value
    End Set
  End Property
  Public Property XD010() As String
    Get
      Return _XD010
    End Get
    Set(ByVal value As String)
      _XD010 = value
    End Set
  End Property
  Public Property XD011() As String
    Get
      Return _XD011
    End Get
    Set(ByVal value As String)
      _XD011 = value
    End Set
  End Property
  Public Property XD012() As String
    Get
      Return _XD012
    End Get
    Set(ByVal value As String)
      _XD012 = value
    End Set
  End Property
  Public Property XD013() As String
    Get
      Return _XD013
    End Get
    Set(ByVal value As String)
      _XD013 = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal XD001 As String, ByVal XD002 As String, ByVal XD003 As String, ByVal XD004 As String, ByVal XD005 As String, ByVal XD006 As Decimal, ByVal XD007 As String, ByVal XD008 As String, ByVal XD009 As String, ByVal XD010 As String, ByVal XD011 As String, ByVal XD012 As String, ByVal XD013 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(XD001, XD002, XD004)
      _gid = key
      _XD001 = XD001
      _XD002 = XD002
      _XD003 = XD003
      _XD004 = XD004
      _XD005 = XD005
      _XD006 = XD006
      _XD007 = XD007
      _XD008 = XD008
      _XD009 = XD009
      _XD010 = XD010
      _XD011 = XD011
      _XD012 = XD012
      _XD013 = XD013
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
  Public Shared Function Get_Combination_Key(ByVal XD001 As String, ByVal XD002 As String, ByVal XD004 As String) As String
    Try
      Dim key As String = XD001 & LinkKey & XD002 & LinkKey & XD004
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsINVXD
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
      Dim strSQL As String = INVXDManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = INVXDManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = INVXDManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
