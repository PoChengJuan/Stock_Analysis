 Public Class clsINVMB
 Private ShareName As String = "INVMB"
 Private ShareKey As String = ""
 Private _gid As String
 Private _MB001 As String '品號 
 Private _MB002 As String '品名 
 Private _MB003 As String '規格 
 Private _MB004 As String '庫存單位 
 Private _MB005 As String '品號分類一 
 Private _MB006 As String '品號分類二 
 Private _MB007 As String '品號分類三 
 Private _MB008 As String '品號分類四 
 Private _MB009 As String '商品描述 
 Private _MB014 As Decimal '單位淨重 
  
  
 Public Property gid() As String
 Get
 Return _gid
 End Get
 Set(ByVal value As String)
 _gid = value
 End Set
 End Property
 Public Property MB001() As String
 Get
 Return _MB001
 End Get
 Set(ByVal value As String)
 _MB001 = value
 End Set
 End Property
 Public Property MB002() As String
 Get
 Return _MB002
 End Get
 Set(ByVal value As String)
 _MB002 = value
 End Set
 End Property
 Public Property MB003() As String
 Get
 Return _MB003
 End Get
 Set(ByVal value As String)
 _MB003 = value
 End Set
 End Property
 Public Property MB004() As String
 Get
 Return _MB004
 End Get
 Set(ByVal value As String)
 _MB004 = value
 End Set
 End Property
 Public Property MB005() As String
 Get
 Return _MB005
 End Get
 Set(ByVal value As String)
 _MB005 = value
 End Set
 End Property
 Public Property MB006() As String
 Get
 Return _MB006
 End Get
 Set(ByVal value As String)
 _MB006 = value
 End Set
 End Property
 Public Property MB007() As String
 Get
 Return _MB007
 End Get
 Set(ByVal value As String)
 _MB007 = value
 End Set
 End Property
 Public Property MB008() As String
 Get
 Return _MB008
 End Get
 Set(ByVal value As String)
 _MB008 = value
 End Set
 End Property
 Public Property MB009() As String
 Get
 Return _MB009
 End Get
 Set(ByVal value As String)
 _MB009 = value
 End Set
 End Property
 Public Property MB014() As Decimal
 Get
 Return _MB014
 End Get
 Set(ByVal value As Decimal)
 _MB014 = value
 End Set
 End Property
  
  
 '物件建立時執行的事件
 Public Sub New( ByVal MB001 As String, ByVal MB002 As String, ByVal MB003 As String, ByVal MB004 As String, ByVal MB005 As String, ByVal MB006 As String, ByVal MB007 As String, ByVal MB008 As String, ByVal MB009 As String, ByVal MB014 As Decimal)
 MyBase.New()
 Try
 Dim key As String = Get_Combination_Key(MB001)
 _gid = key
 _MB001 = MB001
 _MB002 = MB002
 _MB003 = MB003
 _MB004 = MB004
 _MB005 = MB005
 _MB006 = MB006
 _MB007 = MB007
 _MB008 = MB008
 _MB009 = MB009
 _MB014 = MB014
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
 Public Shared Function Get_Combination_Key( ByVal MB001 As String) As String
 Try
 Dim key As String = MB001
 Return key
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return ""
 End Try
 End Function
 Public Function Clone() As clsINVMB
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
 Dim strSQL As String = INVMBManagement.GetInsertSQL(Me)
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
 Dim strSQL As String = INVMBManagement.GetUpdateSQL(Me)
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
 Dim strSQL As String = INVMBManagement.GetDeleteSQL(Me)
 lstSQL.Add(strSQL)
 Return True
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return False
 End Try
 End Function
  
  
 End Class
