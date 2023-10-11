Public Class clsPURTC
  Private ShareName As String = "PURTC"
  Private ShareKey As String = ""
  Private _gid As String
  Private _TC001 As String '採購單別 
  Private _TC002 As String '採購單號 
  Private _TC003 As String '採購日期 
  Private _TC014 As String '確認碼

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property TC001() As String
    Get
      Return _TC001
    End Get
    Set(ByVal value As String)
      _TC001 = value
    End Set
  End Property
  Public Property TC002() As String
    Get
      Return _TC002
    End Get
    Set(ByVal value As String)
      _TC002 = value
    End Set
  End Property
  Public Property TC003() As String
    Get
      Return _TC003
    End Get
    Set(ByVal value As String)
      _TC003 = value
    End Set
  End Property
  Public Property TC014() As String
    Get
      Return _TC014
    End Get
    Set(ByVal value As String)
      _TC014 = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal TC001 As String, ByVal TC002 As String, ByVal TC003 As String, ByVal TC014 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(TC001, TC002)
      _gid = key
      _TC001 = TC001
      _TC002 = TC002
      _TC003 = TC003
      _TC014 = TC014
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
  Public Shared Function Get_Combination_Key(ByVal TC001 As String, ByVal TC002 As String) As String
    Try
      Dim key As String = TC001 & LinkKey & TC002
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPURTC
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
      Dim strSQL As String = PURTCManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = PURTCManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = PURTCManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
