Public Class clsPO_LINE
    Private ShareName As String = " & item.tableName & "
    Private ShareKey As String = " & "
    Private _gid As String
    Private _PO_ID As String '訂單編號

    Private _PO_LINE_NO As String '訂單明細編號(上傳時使用)

    Private _QTY As Double '需求量

    Private _QTY_FINISH As Double '已完成數量

    Private _H_QTY_PROCESS As Double '已上傳數量

    Private _H_POL1 As String

    Private _H_POL2 As String

    Private _H_POL3 As String

    Private _H_POL4 As String

    Private _H_POL5 As String


    Public Property H_POL5() As String
        Get
            Return _H_POL5
        End Get
        Set(ByVal value As String)
            _H_POL5 = value
        End Set
    End Property

    Public Property H_POL4() As String
        Get
            Return _H_POL4
        End Get
        Set(ByVal value As String)
            _H_POL4 = value
        End Set
    End Property

    Public Property H_POL3() As String
        Get
            Return _H_POL3
        End Get
        Set(ByVal value As String)
            _H_POL3 = value
        End Set
    End Property

    Public Property H_POL2() As String
        Get
            Return _H_POL2
        End Get
        Set(ByVal value As String)
            _H_POL2 = value
        End Set
    End Property

    Public Property H_POL1() As String
        Get
            Return _H_POL1
        End Get
        Set(ByVal value As String)
            _H_POL1 = value
        End Set
    End Property

    Public Property gid() As String
        Get
            Return _gid
        End Get
        Set(ByVal value As String)
            _gid = value
        End Set
    End Property
    Public Property PO_ID() As String
        Get
            Return _PO_ID
        End Get
        Set(ByVal value As String)
            _PO_ID = value
        End Set
    End Property
    Public Property PO_LINE_NO() As String
        Get
            Return _PO_LINE_NO
        End Get
        Set(ByVal value As String)
            _PO_LINE_NO = value
        End Set
    End Property
    Public Property QTY() As Double
        Get
            Return _QTY
        End Get
        Set(ByVal value As Double)
            _QTY = value
        End Set
    End Property
    Public Property QTY_FINISH() As Double
        Get
            Return _QTY_FINISH
        End Get
        Set(ByVal value As Double)
            _QTY_FINISH = value
        End Set
    End Property
    Public Property H_QTY_PROCESS() As Double
        Get
            Return _H_QTY_PROCESS
        End Get
        Set(ByVal value As Double)
            _H_QTY_PROCESS = value
        End Set
    End Property

    Public Sub New(ByVal PO_ID As String, ByVal PO_LINE_NO As String, ByVal QTY As Double, ByVal QTY_FINISH As Double, ByVal H_QTY_PROCESS As Double, ByVal H_POL1 As String, ByVal H_POL2 As String, ByVal H_POL3 As String, ByVal H_POL4 As String, ByVal H_POL5 As String)
        MyBase.New()
        Try
            Dim key As String = Get_Combination_Key(PO_ID, PO_LINE_NO)
            _gid = key
            _PO_ID = PO_ID
            _PO_LINE_NO = PO_LINE_NO
            _QTY = QTY
            _QTY_FINISH = QTY_FINISH
            _H_QTY_PROCESS = H_QTY_PROCESS
            _H_POL1 = H_POL1
            _H_POL2 = H_POL2
            _H_POL3 = H_POL3
            _H_POL4 = H_POL4
            _H_POL5 = H_POL5
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End Try
    End Sub
    '物件結束時觸發的事件，用來清除物件的內容
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Private Sub Class_Terminate_Renamed()
        '目的:結束物件
    End Sub
    '=================Public Function=======================
    '傳入指定參數取得Key值
    Public Shared Function Get_Combination_Key(ByVal PO_ID As String, ByVal PO_LINE_NO As String) As String
        Try
            Dim key As String = PO_ID & LinkKey & PO_LINE_NO
            Return key
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return ""
        End Try
    End Function
    Public Function Clone() As clsPO_LINE
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
            Dim strSQL As String = WMS_T_PO_LINEManagement.GetInsertSQL(Me)
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
            Dim strSQL As String = WMS_T_PO_LINEManagement.GetUpdateSQL(Me)
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
            Dim strSQL As String = WMS_T_PO_LINEManagement.GetDeleteSQL(Me)
            lstSQL.Add(strSQL)
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Public Function Update_To_Memory(ByRef objWMS_T_PO_LINE As clsPO_LINE) As Boolean
        Try
            Dim key As String = objWMS_T_PO_LINE._gid
            If key <> _gid Then
                SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
            End If
            _PO_ID = objWMS_T_PO_LINE.PO_ID
            _PO_LINE_NO = objWMS_T_PO_LINE.PO_LINE_NO
            _QTY = objWMS_T_PO_LINE.QTY
            _QTY_FINISH = objWMS_T_PO_LINE.QTY_FINISH
            _H_QTY_PROCESS = objWMS_T_PO_LINE.H_QTY_PROCESS
            _H_POL1 = objWMS_T_PO_LINE.H_POL1
            _H_POL2 = objWMS_T_PO_LINE.H_POL2
            _H_POL3 = objWMS_T_PO_LINE.H_POL3
            _H_POL4 = objWMS_T_PO_LINE.H_POL4
            _H_POL5 = objWMS_T_PO_LINE.H_POL5

            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function

    Public Function Update_Data_By_Forced(ByRef objWMS_T_PO_LINE As clsPO_LINE) As Boolean
        Try

            _QTY = objWMS_T_PO_LINE.QTY
            _H_POL1 = objWMS_T_PO_LINE.H_POL1
            _H_POL2 = objWMS_T_PO_LINE.H_POL2
            _H_POL3 = objWMS_T_PO_LINE.H_POL3
            _H_POL4 = objWMS_T_PO_LINE.H_POL4
            _H_POL5 = objWMS_T_PO_LINE.H_POL5

            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Public Function Update_Data_H_POL1() As Boolean
        Try

            _H_POL1 = ModuleHelpFunc.GetNewTime_ByDataTimeFormat(DBDateFormat)

            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Public Function Update_QTY(ByVal QTY As Long) As Boolean
        Try

            _QTY = QTY


            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function

End Class
