Public Class clsINVXF
  Private ShareName As String = "INVXF"
  Private ShareKey As String = ""
  Private _gid As String
  Private _XF001 As String '單別 
  Private _XF002 As String '單號 
  Private _XF003 As String '單據性質碼 
  Private _XF004 As String '序號 
  Private _XF005 As String '品號 
  Private _XF006 As Decimal '數量 
  Private _XF007 As String '轉出庫 
  Private _XF008 As String '轉入庫 
  Private _XF009 As String '更新碼 
  Private _XF010 As String '公司別 
  Private _XF011 As String '建立者 
  Private _XF012 As String '確認者 
  Private _XF013 As String '已完成數量



  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property XF001() As String
    Get
      Return _XF001
    End Get
    Set(ByVal value As String)
      _XF001 = value
    End Set
  End Property
  Public Property XF002() As String
    Get
      Return _XF002
    End Get
    Set(ByVal value As String)
      _XF002 = value
    End Set
  End Property
  Public Property XF003() As String
    Get
      Return _XF003
    End Get
    Set(ByVal value As String)
      _XF003 = value
    End Set
  End Property
  Public Property XF004() As String
    Get
      Return _XF004
    End Get
    Set(ByVal value As String)
      _XF004 = value
    End Set
  End Property
  Public Property XF005() As String
    Get
      Return _XF005
    End Get
    Set(ByVal value As String)
      _XF005 = value
    End Set
  End Property
  Public Property XF006() As Decimal
    Get
      Return _XF006
    End Get
    Set(ByVal value As Decimal)
      _XF006 = value
    End Set
  End Property
  Public Property XF007() As String
    Get
      Return _XF007
    End Get
    Set(ByVal value As String)
      _XF007 = value
    End Set
  End Property
  Public Property XF008() As String
    Get
      Return _XF008
    End Get
    Set(ByVal value As String)
      _XF008 = value
    End Set
  End Property
  Public Property XF009() As String
    Get
      Return _XF009
    End Get
    Set(ByVal value As String)
      _XF009 = value
    End Set
  End Property
  Public Property XF010() As String
    Get
      Return _XF010
    End Get
    Set(ByVal value As String)
      _XF010 = value
    End Set
  End Property
  Public Property XF011() As String
    Get
      Return _XF011
    End Get
    Set(ByVal value As String)
      _XF011 = value
    End Set
  End Property
  Public Property XF012() As String
    Get
      Return _XF012
    End Get
    Set(ByVal value As String)
      _XF012 = value
    End Set
  End Property
  Public Property XF013() As String
    Get
      Return _XF013
    End Get
    Set(ByVal value As String)
      _XF013 = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal XF001 As String, ByVal XF002 As String, ByVal XF003 As String, ByVal XF004 As String, ByVal XF005 As String, ByVal XF006 As Decimal, ByVal XF007 As String, ByVal XF008 As String, ByVal XF009 As String, ByVal XF010 As String, ByVal XF011 As String, ByVal XF012 As String, ByVal XF013 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(XF001, XF002, XF004)
      _gid = key
      _XF001 = XF001
      _XF002 = XF002
      _XF003 = XF003
      _XF004 = XF004
      _XF005 = XF005
      _XF006 = XF006
      _XF007 = XF007
      _XF008 = XF008
      _XF009 = XF009
      _XF010 = XF010
      _XF011 = XF011
      _XF012 = XF012
      _XF013 = XF013
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
  Public Shared Function Get_Combination_Key(ByVal XF001 As String, ByVal XF002 As String, ByVal XF004 As String) As String
    Try
      Dim key As String = XF001 & LinkKey & XF002 & LinkKey & XF004
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsINVXF
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
      Dim strSQL As String = INVXFManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = INVXFManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = INVXFManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
