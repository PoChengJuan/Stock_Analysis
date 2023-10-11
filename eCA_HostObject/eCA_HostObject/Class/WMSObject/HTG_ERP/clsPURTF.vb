Public Class clsPURTF
    Private ShareName As String = "PURTF"
    Private ShareKey As String = ""
    Private _gid As String
    Private _TF001 As String '採購變更單別 
    Private _TF002 As String '採購變更單號 
    Private _TF003 As String '版次 
    Private _TF004 As String '序號 
    Private _TF005 As String '品號 
    Private _TF006 As String '品名 
    Private _TF007 As String '規格 
    Private _TF008 As String '交貨庫別 
    Private _TF009 As Decimal '採購數量 
    Private _TF010 As String '單位 
    Private _TF011 As Decimal '採購單價 
    Private _TF012 As Decimal '採購金額 


    Public Property gid() As String
        Get
            Return _gid
        End Get
        Set(ByVal value As String)
            _gid = value
        End Set
    End Property
    Public Property TF001() As String
        Get
            Return _TF001
        End Get
        Set(ByVal value As String)
            _TF001 = value
        End Set
    End Property
    Public Property TF002() As String
        Get
            Return _TF002
        End Get
        Set(ByVal value As String)
            _TF002 = value
        End Set
    End Property
    Public Property TF003() As String
        Get
            Return _TF003
        End Get
        Set(ByVal value As String)
            _TF003 = value
        End Set
    End Property
    Public Property TF004() As String
        Get
            Return _TF004
        End Get
        Set(ByVal value As String)
            _TF004 = value
        End Set
    End Property
    Public Property TF005() As String
        Get
            Return _TF005
        End Get
        Set(ByVal value As String)
            _TF005 = value
        End Set
    End Property
    Public Property TF006() As String
        Get
            Return _TF006
        End Get
        Set(ByVal value As String)
            _TF006 = value
        End Set
    End Property
    Public Property TF007() As String
        Get
            Return _TF007
        End Get
        Set(ByVal value As String)
            _TF007 = value
        End Set
    End Property
    Public Property TF008() As String
        Get
            Return _TF008
        End Get
        Set(ByVal value As String)
            _TF008 = value
        End Set
    End Property
    Public Property TF009() As Decimal
        Get
            Return _TF009
        End Get
        Set(ByVal value As Decimal)
            _TF009 = value
        End Set
    End Property
    Public Property TF010() As String
        Get
            Return _TF010
        End Get
        Set(ByVal value As String)
            _TF010 = value
        End Set
    End Property
    Public Property TF011() As Decimal
        Get
            Return _TF011
        End Get
        Set(ByVal value As Decimal)
            _TF011 = value
        End Set
    End Property
    Public Property TF012() As Decimal
        Get
            Return _TF012
        End Get
        Set(ByVal value As Decimal)
            _TF012 = value
        End Set
    End Property


    '物件建立時執行的事件
    Public Sub New(ByVal TF001 As String, ByVal TF002 As String, ByVal TF003 As String, ByVal TF004 As String, ByVal TF005 As String, ByVal TF006 As String, ByVal TF007 As String, ByVal TF008 As String, ByVal TF009 As Decimal, ByVal TF010 As String, ByVal TF011 As Decimal, ByVal TF012 As Decimal)
        MyBase.New()
        Try
            Dim key As String = Get_Combination_Key(TF001, TF002, TF003)
            _gid = key
            _TF001 = TF001
            _TF002 = TF002
            _TF003 = TF003
            _TF004 = TF004
            _TF005 = TF005
            _TF006 = TF006
            _TF007 = TF007
            _TF008 = TF008
            _TF009 = TF009
            _TF010 = TF010
            _TF011 = TF011
            _TF012 = TF012
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
    Public Shared Function Get_Combination_Key(ByVal TF001 As String, ByVal TF002 As String, ByVal TF003 As String) As String
        Try
            Dim key As String = TF001 & LinkKey & TF002 & LinkKey & TF003
            Return key
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return ""
        End Try
    End Function
    Public Function Clone() As clsPURTF
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
            Dim strSQL As String = PURTFManagement.GetInsertSQL(Me)
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
            Dim strSQL As String = PURTFManagement.GetUpdateSQL(Me)
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
            Dim strSQL As String = PURTFManagement.GetDeleteSQL(Me)
            lstSQL.Add(strSQL)
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function


End Class
