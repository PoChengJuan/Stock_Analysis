Public Class clsMOCTP
  Private ShareName As String = "MOCTP"
  Private ShareKey As String = ""
  Private _gid As String
  Private _TP001 As String '製令單別 
  Private _TP002 As String '製令單號 
  Private _TP003 As String '變更版次 
  Private _TP004 As String '新材料品號 
  Private _TP005 As Long '新需領用量 
  Private _TP006 As Long '新已領用量 
  Private _TP007 As String '新製程代號 
  Private _TP008 As String '新單位 


  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property TP001() As String
    Get
      Return _TP001
    End Get
    Set(ByVal value As String)
      _TP001 = value
    End Set
  End Property
  Public Property TP002() As String
    Get
      Return _TP002
    End Get
    Set(ByVal value As String)
      _TP002 = value
    End Set
  End Property
  Public Property TP003() As String
    Get
      Return _TP003
    End Get
    Set(ByVal value As String)
      _TP003 = value
    End Set
  End Property
  Public Property TP004() As String
    Get
      Return _TP004
    End Get
    Set(ByVal value As String)
      _TP004 = value
    End Set
  End Property
  Public Property TP005() As Long
    Get
      Return _TP005
    End Get
    Set(ByVal value As Long)
      _TP005 = value
    End Set
  End Property
  Public Property TP006() As Long
    Get
      Return _TP006
    End Get
    Set(ByVal value As Long)
      _TP006 = value
    End Set
  End Property
  Public Property TP007() As String
    Get
      Return _TP007
    End Get
    Set(ByVal value As String)
      _TP007 = value
    End Set
  End Property
  Public Property TP008() As String
    Get
      Return _TP008
    End Get
    Set(ByVal value As String)
      _TP008 = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal TP001 As String, ByVal TP002 As String, ByVal TP003 As String, ByVal TP004 As String, ByVal TP005 As Long, ByVal TP006 As Long, ByVal TP007 As String, ByVal TP008 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(TP001, TP002, TP003)
      _gid = key
      _TP001 = TP001
      _TP002 = TP002
      _TP003 = TP003
      _TP004 = TP004
      _TP005 = TP005
      _TP006 = TP006
      _TP007 = TP007
      _TP008 = TP008
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
  Public Shared Function Get_Combination_Key(ByVal TP001 As String, ByVal TP002 As String, ByVal TP003 As String) As String
    Try
      Dim key As String = TP001 & LinkKey & TP002 & LinkKey & TP003
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMOCTP
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
      Dim strSQL As String = MOCTPManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = MOCTPManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = MOCTPManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
