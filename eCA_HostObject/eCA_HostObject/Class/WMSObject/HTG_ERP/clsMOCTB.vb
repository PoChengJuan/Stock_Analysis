 Public Class clsMOCTB
 Private ShareName As String = "MOCTB"
 Private ShareKey As String = ""
 Private _gid As String
 Private _TB001 As String '製令單別 
 Private _TB002 As String '製令單號 
 Private _TB003 As String '材料品號 
 Private _TB004 As Decimal '需領用量 
 Private _TB005 As Decimal '已領用量 
 Private _TB006 As String '製程代號 
 Private _TB007 As String '單位 
  
  
 Public Property gid() As String
 Get
 Return _gid
 End Get
 Set(ByVal value As String)
 _gid = value
 End Set
 End Property
 Public Property TB001() As String
 Get
 Return _TB001
 End Get
 Set(ByVal value As String)
 _TB001 = value
 End Set
 End Property
 Public Property TB002() As String
 Get
 Return _TB002
 End Get
 Set(ByVal value As String)
 _TB002 = value
 End Set
 End Property
 Public Property TB003() As String
 Get
 Return _TB003
 End Get
 Set(ByVal value As String)
 _TB003 = value
 End Set
 End Property
 Public Property TB004() As Decimal
 Get
 Return _TB004
 End Get
 Set(ByVal value As Decimal)
 _TB004 = value
 End Set
 End Property
 Public Property TB005() As Decimal
 Get
 Return _TB005
 End Get
 Set(ByVal value As Decimal)
 _TB005 = value
 End Set
 End Property
 Public Property TB006() As String
 Get
 Return _TB006
 End Get
 Set(ByVal value As String)
 _TB006 = value
 End Set
 End Property
 Public Property TB007() As String
 Get
 Return _TB007
 End Get
 Set(ByVal value As String)
 _TB007 = value
 End Set
 End Property
  
  
 '物件建立時執行的事件
 Public Sub New( ByVal TB001 As String, ByVal TB002 As String, ByVal TB003 As String, ByVal TB004 As Decimal, ByVal TB005 As Decimal, ByVal TB006 As String, ByVal TB007 As String)
 MyBase.New()
 Try
      Dim key As String = Get_Combination_Key(TB001, TB002, TB003)
      _gid = key
 _TB001 = TB001
 _TB002 = TB002
 _TB003 = TB003
 _TB004 = TB004
 _TB005 = TB005
 _TB006 = TB006
 _TB007 = TB007
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
  Public Shared Function Get_Combination_Key(ByVal TB001 As String, ByVal TB002 As String, ByVal TB003 As String) As String
    Try
      Dim key As String = TB001 & LinkKey & TB002 & LinkKey & TB003
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMOCTB
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
 Dim strSQL As String = MOCTBManagement.GetInsertSQL(Me)
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
 Dim strSQL As String = MOCTBManagement.GetUpdateSQL(Me)
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
 Dim strSQL As String = MOCTBManagement.GetDeleteSQL(Me)
 lstSQL.Add(strSQL)
 Return True
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return False
 End Try
 End Function
  
  
 End Class
