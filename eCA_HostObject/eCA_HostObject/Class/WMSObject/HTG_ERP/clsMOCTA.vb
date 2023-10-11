Public Class clsMOCTA
  Private ShareName As String = "MOCTA"
  Private ShareKey As String = ""
  Private _gid As String
  Private _TA001 As String '製令單別 
  Private _TA002 As String '製令單號 
  Private _TA003 As String '開單日期 
  Private _TA004 As String 'BOM日期 
  Private _TA005 As String 'BOM版次 
  Private _TA006 As String '產品品號 
  Private _TA007 As String '單位 
  Private _TA015 As String '數量
  Private _TA020 As String '庫別 
  Private _TA033 As String '計劃批號


  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property TA001() As String
    Get
      Return _TA001
    End Get
    Set(ByVal value As String)
      _TA001 = value
    End Set
  End Property
  Public Property TA002() As String
    Get
      Return _TA002
    End Get
    Set(ByVal value As String)
      _TA002 = value
    End Set
  End Property
  Public Property TA003() As String
    Get
      Return _TA003
    End Get
    Set(ByVal value As String)
      _TA003 = value
    End Set
  End Property
  Public Property TA004() As String
    Get
      Return _TA004
    End Get
    Set(ByVal value As String)
      _TA004 = value
    End Set
  End Property
  Public Property TA005() As String
    Get
      Return _TA005
    End Get
    Set(ByVal value As String)
      _TA005 = value
    End Set
  End Property
  Public Property TA006() As String
    Get
      Return _TA006
    End Get
    Set(ByVal value As String)
      _TA006 = value
    End Set
  End Property
  Public Property TA007() As String
    Get
      Return _TA007
    End Get
    Set(ByVal value As String)
      _TA007 = value
    End Set
  End Property
  Public Property TA015() As String
    Get
      Return _TA015
    End Get
    Set(ByVal value As String)
      _TA015 = value
    End Set
  End Property
  Public Property TA020() As String  '庫別
    Get
      Return _TA020
    End Get
    Set(ByVal value As String) '庫別
      _TA020 = value
    End Set
  End Property
  Public Property TA033() As String  '計劃批號
    Get
      Return _TA033
    End Get
    Set(ByVal value As String) '計劃批號
      _TA033 = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal TA001 As String, ByVal TA002 As String, ByVal TA003 As String, ByVal TA004 As String, ByVal TA005 As String, ByVal TA006 As String, ByVal TA007 As String, ByVal TA015 As Decimal, ByVal TA020 As String, ByVal TA033 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(TA001, TA002)
      _gid = key
      _TA001 = TA001
      _TA002 = TA002
      _TA003 = TA003
      _TA004 = TA004
      _TA005 = TA005
      _TA006 = TA006
      _TA007 = TA007
      _TA015 = TA015
      _TA020 = TA020
      _TA033 = TA033
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
  Public Shared Function Get_Combination_Key(ByVal TA001 As String, ByVal TA002 As String) As String
    Try
      Dim key As String = TA001 & LinkKey & TA002
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMOCTA
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
      Dim strSQL As String = MOCTAManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = MOCTAManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = MOCTAManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
