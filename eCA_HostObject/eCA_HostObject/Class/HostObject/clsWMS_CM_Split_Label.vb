Public Class clsWMS_CM_Split_Label
    Private ShareName As String = "WMS_CM_Split_Label"
    Private ShareKey As String = ""
    Private _gid As String
    Private _COMPANY_CODE As String '公司代碼

    Private _MATERIAL_NUMBER As String 'Material Number

    Private _PLANT As String 'Plant

    Private _BATCH As String 'Batch Number

    Private _LAST_CHANGE_DATE As String 'Date of Last Change

    Private _BIN As String 'BIN

    Private _GP As String 'GP

    Private _HF As String 'HF

    Private _EXP_DATE As String 'EXP_DATE

    Private _LOT As String 'LOT

    Private _DC As String 'DC

    Private _SAFETY As String 'SAFETY

    Private _ESD As String 'ESD

    Private _MFR As String 'manufacturer

    Private _MFRPN As String 'manufacturer part number

    Private _LABEL_COMMON1 As String '備用字段

    Private _LABEL_COMMON2 As String '備用字段

    Private _LABEL_COMMON3 As String '備用字段

    Private _LABEL_COMMON4 As String '備用字段

    Private _LABEL_COMMON5 As String '備用字段

    Private _UPDATE_TIME As String '更新時間

    Public Property gid() As String
        Get
            Return _gid
        End Get
        Set(ByVal value As String)
            _gid = value
        End Set
    End Property
    Public Property COMPANY_CODE() As String
        Get
            Return _COMPANY_CODE
        End Get
        Set(ByVal value As String)
            _COMPANY_CODE = value
        End Set
    End Property
    Public Property MATERIAL_NUMBER() As String
        Get
            Return _MATERIAL_NUMBER
        End Get
        Set(ByVal value As String)
            _MATERIAL_NUMBER = value
        End Set
    End Property
    Public Property PLANT() As String
        Get
            Return _PLANT
        End Get
        Set(ByVal value As String)
            _PLANT = value
        End Set
    End Property
    Public Property BATCH() As String
        Get
            Return _BATCH
        End Get
        Set(ByVal value As String)
            _BATCH = value
        End Set
    End Property
    Public Property LAST_CHANGE_DATE() As String
        Get
            Return _LAST_CHANGE_DATE
        End Get
        Set(ByVal value As String)
            _LAST_CHANGE_DATE = value
        End Set
    End Property
    Public Property BIN() As String
        Get
            Return _BIN
        End Get
        Set(ByVal value As String)
            _BIN = value
        End Set
    End Property
    Public Property GP() As String
        Get
            Return _GP
        End Get
        Set(ByVal value As String)
            _GP = value
        End Set
    End Property
    Public Property HF() As String
        Get
            Return _HF
        End Get
        Set(ByVal value As String)
            _HF = value
        End Set
    End Property
    Public Property EXP_DATE() As String
        Get
            Return _EXP_DATE
        End Get
        Set(ByVal value As String)
            _EXP_DATE = value
        End Set
    End Property
    Public Property LOT() As String
        Get
            Return _LOT
        End Get
        Set(ByVal value As String)
            _LOT = value
        End Set
    End Property
    Public Property DC() As String
        Get
            Return _DC
        End Get
        Set(ByVal value As String)
            _DC = value
        End Set
    End Property
    Public Property SAFETY() As String
        Get
            Return _SAFETY
        End Get
        Set(ByVal value As String)
            _SAFETY = value
        End Set
    End Property
    Public Property ESD() As String
        Get
            Return _ESD
        End Get
        Set(ByVal value As String)
            _ESD = value
        End Set
    End Property
    Public Property MFR() As String
        Get
            Return _MFR
        End Get
        Set(ByVal value As String)
            _MFR = value
        End Set
    End Property
    Public Property MFRPN() As String
        Get
            Return _MFRPN
        End Get
        Set(ByVal value As String)
            _MFRPN = value
        End Set
    End Property
    Public Property LABEL_COMMON1() As String
        Get
            Return _LABEL_COMMON1
        End Get
        Set(ByVal value As String)
            _LABEL_COMMON1 = value
        End Set
    End Property
    Public Property LABEL_COMMON2() As String
        Get
            Return _LABEL_COMMON2
        End Get
        Set(ByVal value As String)
            _LABEL_COMMON2 = value
        End Set
    End Property
    Public Property LABEL_COMMON3() As String
        Get
            Return _LABEL_COMMON3
        End Get
        Set(ByVal value As String)
            _LABEL_COMMON3 = value
        End Set
    End Property
    Public Property LABEL_COMMON4() As String
        Get
            Return _LABEL_COMMON4
        End Get
        Set(ByVal value As String)
            _LABEL_COMMON4 = value
        End Set
    End Property
    Public Property LABEL_COMMON5() As String
        Get
            Return _LABEL_COMMON5
        End Get
        Set(ByVal value As String)
            _LABEL_COMMON5 = value
        End Set
    End Property
    Public Property UPDATE_TIME() As String
        Get
            Return _UPDATE_TIME
        End Get
        Set(ByVal value As String)
            _UPDATE_TIME = value
        End Set
    End Property

    Public Sub New(ByVal COMPANY_CODE As String, ByVal MATERIAL_NUMBER As String, ByVal PLANT As String, ByVal BATCH As String, ByVal LAST_CHANGE_DATE As String, ByVal BIN As String, ByVal GP As String, ByVal HF As String, ByVal EXP_DATE As String, ByVal LOT As String, ByVal DC As String, ByVal SAFETY As String, ByVal ESD As String, ByVal MFR As String, ByVal MFRPN As String, ByVal LABEL_COMMON1 As String, ByVal LABEL_COMMON2 As String, ByVal LABEL_COMMON3 As String, ByVal LABEL_COMMON4 As String, ByVal LABEL_COMMON5 As String, ByVal UPDATE_TIME As String)
        MyBase.New()
        Try
            Dim key As String = Get_Combination_Key(COMPANY_CODE, MATERIAL_NUMBER, PLANT, BATCH)
            _gid = key
            _COMPANY_CODE = COMPANY_CODE
            _MATERIAL_NUMBER = MATERIAL_NUMBER
            _PLANT = PLANT
            _BATCH = BATCH
            _LAST_CHANGE_DATE = LAST_CHANGE_DATE
            _BIN = BIN
            _GP = GP
            _HF = HF
            _EXP_DATE = EXP_DATE
            _LOT = LOT
            _DC = DC
            _SAFETY = SAFETY
            _ESD = ESD
            _MFR = MFR
            _MFRPN = MFRPN
            _LABEL_COMMON1 = LABEL_COMMON1
            _LABEL_COMMON2 = LABEL_COMMON2
            _LABEL_COMMON3 = LABEL_COMMON3
            _LABEL_COMMON4 = LABEL_COMMON4
            _LABEL_COMMON5 = LABEL_COMMON5
            _UPDATE_TIME = UPDATE_TIME
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
    Public Shared Function Get_Combination_Key(ByVal COMPANY_CODE As String, ByVal MATERIAL_NUMBER As String, ByVal PLANT As String, ByVal BATCH As String) As String
        Try
            Dim key As String = COMPANY_CODE & LinkKey & MATERIAL_NUMBER & LinkKey & PLANT & LinkKey & BATCH
            Return key
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return ""
        End Try
    End Function
    Public Function Clone() As clsWMS_CM_Split_Label
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
            Dim strSQL As String = WMS_CM_Split_LabelManagement.GetInsertSQL(Me)
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
            Dim strSQL As String = WMS_CM_Split_LabelManagement.GetUpdateSQL(Me)
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
            Dim strSQL As String = WMS_CM_Split_LabelManagement.GetDeleteSQL(Me)
            lstSQL.Add(strSQL)
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Public Function Update_To_Memory(ByRef objWMS_CM_Split_Label As clsWMS_CM_Split_Label) As Boolean
        Try
            Dim key As String = objWMS_CM_Split_Label._gid
            If key <> _gid Then
                SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
            End If
            _COMPANY_CODE = COMPANY_CODE
            _MATERIAL_NUMBER = MATERIAL_NUMBER
            _PLANT = PLANT
            _BATCH = BATCH
            _LAST_CHANGE_DATE = LAST_CHANGE_DATE
            _BIN = BIN
            _GP = GP
            _HF = HF
            _EXP_DATE = EXP_DATE
            _LOT = LOT
            _DC = DC
            _SAFETY = SAFETY
            _ESD = ESD
            _MFR = MFR
            _MFRPN = MFRPN
            _LABEL_COMMON1 = LABEL_COMMON1
            _LABEL_COMMON2 = LABEL_COMMON2
            _LABEL_COMMON3 = LABEL_COMMON3
            _LABEL_COMMON4 = LABEL_COMMON4
            _LABEL_COMMON5 = LABEL_COMMON5
            _UPDATE_TIME = UPDATE_TIME
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
End Class
