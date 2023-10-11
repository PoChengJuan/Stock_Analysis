Public Class clsINVXB
  Private ShareName As String = "INVXB"
  Private ShareKey As String = ""
  Private _gid As String
  Private _XB001 As String '品號 
  Private _XB002 As String '品名 
  Private _XB003 As String '規格 
  Private _XB004 As String '庫存單位 
  Private _XB005 As String '主要庫別 
  Private _XB007 As String '料品高度註記(高的只能放高欄位，低的高低都能放)
  Private _XB008 As String '更新碼 
  Private _XB009 As String '料品長度註記(Y = 1.6，N/NULL = 1.2)


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
  Public Property XB005() As String
    Get
      Return _XB005
    End Get
    Set(ByVal value As String)
      _XB005 = value
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
  Public Property XB007() As String
    Get
      Return _XB007
    End Get
    Set(ByVal value As String)
      _XB007 = value
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


  '物件建立時執行的事件
  Public Sub New(ByVal XB001 As String, ByVal XB002 As String, ByVal XB003 As String, ByVal XB004 As String, ByVal XB005 As String, ByVal XB007 As String, ByVal XB008 As String, ByVal XB009 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(XB001)
      _gid = key
      _XB001 = XB001
      _XB002 = XB002
      _XB003 = XB003
      _XB004 = XB004
      _XB005 = XB005
      _XB007 = XB007
      _XB008 = XB008
      _XB009 = XB009
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
  Public Shared Function Get_Combination_Key(ByVal XB001 As String) As String
    Try
      Dim key As String = XB001
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsINVXB
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
      Dim strSQL As String = INVXBManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = INVXBManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = INVXBManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
