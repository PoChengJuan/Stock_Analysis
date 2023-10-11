Public Class clsPURTE
    Private ShareName As String = "PURTE"
    Private ShareKey As String = ""
    Private _gid As String
    Private _TE001 As String '採購變更單別 
    Private _TE002 As String '採購變更單號 
    Private _TE003 As String '版次 
    Private _TE004 As String '變更日期 
    Private _TE005 As String '供應廠商 


    Public Property gid() As String
        Get
            Return _gid
        End Get
        Set(ByVal value As String)
            _gid = value
        End Set
    End Property
    Public Property TE001() As String
        Get
            Return _TE001
        End Get
        Set(ByVal value As String)
            _TE001 = value
        End Set
    End Property
    Public Property TE002() As String
        Get
            Return _TE002
        End Get
        Set(ByVal value As String)
            _TE002 = value
        End Set
    End Property
    Public Property TE003() As String
        Get
            Return _TE003
        End Get
        Set(ByVal value As String)
            _TE003 = value
        End Set
    End Property
    Public Property TE004() As String
        Get
            Return _TE004
        End Get
        Set(ByVal value As String)
            _TE004 = value
        End Set
    End Property
    Public Property TE005() As String
        Get
            Return _TE005
        End Get
        Set(ByVal value As String)
            _TE005 = value
        End Set
    End Property


    '物件建立時執行的事件
    Public Sub New(ByVal TE001 As String, ByVal TE002 As String, ByVal TE003 As String, ByVal TE004 As String, ByVal TE005 As String)
        MyBase.New()
        Try
            Dim key As String = Get_Combination_Key(TE001, TE002, TE003)
            _gid = key
            _TE001 = TE001
            _TE002 = TE002
            _TE003 = TE003
            _TE004 = TE004
            _TE005 = TE005
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
    Public Shared Function Get_Combination_Key(ByVal TE001 As String, ByVal TE002 As String, ByVal TE003 As String) As String
        Try
            Dim key As String = TE001 & LinkKey & TE002 & LinkKey & TE003
            Return key
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return ""
        End Try
    End Function
    Public Function Clone() As clsPURTE
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
            Dim strSQL As String = PURTEManagement.GetInsertSQL(Me)
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
            Dim strSQL As String = PURTEManagement.GetUpdateSQL(Me)
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
            Dim strSQL As String = PURTEManagement.GetDeleteSQL(Me)
            lstSQL.Add(strSQL)
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function


End Class
