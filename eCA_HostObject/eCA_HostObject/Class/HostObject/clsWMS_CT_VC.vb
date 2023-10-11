Public Class clsWMS_CT_VC
    Private ShareName As String = "WMS_CT_VC"
    Private ShareKey As String = ""
    Private _gid As String
    Private _COMPANY_CODE As String '公司編號

    Private _MATERIAL_DOCUMENT_YEAR As String '年分

    Private _MATERIAL_DOCUMENT As String '運單號

    Private _MATERIAL_DOCUMENT_ITEM As String '運單項次

    Private _MATERIAL_NUMBER As String '料號

    Private _PLANT As String 'PLANT

    Private _STORAGE_LOCATION As String 'StorageLocation

    Private _BATCH As String '批次

    Private _QUANTITY As String '數量

    Private _CREATED_DATE As String 'ERP建立日期

    Private _CREATE_TIME As String 'WMS建立時間

    Private _VC_COMMON1 As String '備用字段

    Private _VC_COMMON2 As String '備用字段

    Private _VC_COMMON3 As String '備用字段

    Private _VC_COMMON4 As String '備用字段

    Private _VC_COMMON5 As String '備用字段

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
    Public Property MATERIAL_DOCUMENT_YEAR() As String
        Get
            Return _MATERIAL_DOCUMENT_YEAR
        End Get
        Set(ByVal value As String)
            _MATERIAL_DOCUMENT_YEAR = value
        End Set
    End Property
    Public Property MATERIAL_DOCUMENT() As String
        Get
            Return _MATERIAL_DOCUMENT
        End Get
        Set(ByVal value As String)
            _MATERIAL_DOCUMENT = value
        End Set
    End Property
    Public Property MATERIAL_DOCUMENT_ITEM() As String
        Get
            Return _MATERIAL_DOCUMENT_ITEM
        End Get
        Set(ByVal value As String)
            _MATERIAL_DOCUMENT_ITEM = value
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
    Public Property STORAGE_LOCATION() As String
        Get
            Return _STORAGE_LOCATION
        End Get
        Set(ByVal value As String)
            _STORAGE_LOCATION = value
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
    Public Property QUANTITY() As String
        Get
            Return _QUANTITY
        End Get
        Set(ByVal value As String)
            _QUANTITY = value
        End Set
    End Property
    Public Property CREATED_DATE() As String
        Get
            Return _CREATED_DATE
        End Get
        Set(ByVal value As String)
            _CREATED_DATE = value
        End Set
    End Property
    Public Property CREATE_TIME() As String
        Get
            Return _CREATE_TIME
        End Get
        Set(ByVal value As String)
            _CREATE_TIME = value
        End Set
    End Property
    Public Property VC_COMMON1() As String
        Get
            Return _VC_COMMON1
        End Get
        Set(ByVal value As String)
            _VC_COMMON1 = value
        End Set
    End Property
    Public Property VC_COMMON2() As String
        Get
            Return _VC_COMMON2
        End Get
        Set(ByVal value As String)
            _VC_COMMON2 = value
        End Set
    End Property
    Public Property VC_COMMON3() As String
        Get
            Return _VC_COMMON3
        End Get
        Set(ByVal value As String)
            _VC_COMMON3 = value
        End Set
    End Property
    Public Property VC_COMMON4() As String
        Get
            Return _VC_COMMON4
        End Get
        Set(ByVal value As String)
            _VC_COMMON4 = value
        End Set
    End Property
    Public Property VC_COMMON5() As String
        Get
            Return _VC_COMMON5
        End Get
        Set(ByVal value As String)
            _VC_COMMON5 = value
        End Set
    End Property

    Public Sub New(ByVal COMPANY_CODE As String, ByVal MATERIAL_DOCUMENT_YEAR As String, ByVal MATERIAL_DOCUMENT As String, ByVal MATERIAL_DOCUMENT_ITEM As String, ByVal MATERIAL_NUMBER As String, ByVal PLANT As String, ByVal STORAGE_LOCATION As String, ByVal BATCH As String, ByVal QUANTITY As String, ByVal CREATED_DATE As String, ByVal CREATE_TIME As String, ByVal VC_COMMON1 As String, ByVal VC_COMMON2 As String, ByVal VC_COMMON3 As String, ByVal VC_COMMON4 As String, ByVal VC_COMMON5 As String)
        MyBase.New()
        Try
            Dim key As String = Get_Combination_Key(COMPANY_CODE, MATERIAL_DOCUMENT_YEAR, MATERIAL_DOCUMENT, MATERIAL_DOCUMENT_ITEM)
            _gid = key
            _COMPANY_CODE = COMPANY_CODE
            _MATERIAL_DOCUMENT_YEAR = MATERIAL_DOCUMENT_YEAR
            _MATERIAL_DOCUMENT = MATERIAL_DOCUMENT
            _MATERIAL_DOCUMENT_ITEM = MATERIAL_DOCUMENT_ITEM
            _MATERIAL_NUMBER = MATERIAL_NUMBER
            _PLANT = PLANT
            _STORAGE_LOCATION = STORAGE_LOCATION
            _BATCH = BATCH
            _QUANTITY = QUANTITY
            _CREATED_DATE = CREATED_DATE
            _CREATE_TIME = CREATE_TIME
            _VC_COMMON1 = VC_COMMON1
            _VC_COMMON2 = VC_COMMON2
            _VC_COMMON3 = VC_COMMON3
            _VC_COMMON4 = VC_COMMON4
            _VC_COMMON5 = VC_COMMON5
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
    Public Shared Function Get_Combination_Key(ByVal COMPANY_CODE As String, ByVal MATERIAL_DOCUMENT_YEAR As String, ByVal MATERIAL_DOCUMENT As String, ByVal MATERIAL_DOCUMENT_ITEM As String) As String
        Try
            Dim key As String = COMPANY_CODE & LinkKey & MATERIAL_DOCUMENT_YEAR & LinkKey & MATERIAL_DOCUMENT & LinkKey & MATERIAL_DOCUMENT_ITEM
            Return key
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return ""
        End Try
    End Function
    Public Function Clone() As clsWMS_CT_VC
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
            Dim strSQL As String = WMS_CT_VCManagement.GetInsertSQL(Me)
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
            Dim strSQL As String = WMS_CT_VCManagement.GetUpdateSQL(Me)
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
            Dim strSQL As String = WMS_CT_VCManagement.GetDeleteSQL(Me)
            lstSQL.Add(strSQL)
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Public Function Update_To_Memory(ByRef objWMS_CT_VC As clsWMS_CT_VC) As Boolean
        Try
            Dim key As String = objWMS_CT_VC._gid
            If key <> _gid Then
                SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                Return False
            End If
            _COMPANY_CODE = COMPANY_CODE
            _MATERIAL_DOCUMENT_YEAR = MATERIAL_DOCUMENT_YEAR
            _MATERIAL_DOCUMENT = MATERIAL_DOCUMENT
            _MATERIAL_DOCUMENT_ITEM = MATERIAL_DOCUMENT_ITEM
            _MATERIAL_NUMBER = MATERIAL_NUMBER
            _PLANT = PLANT
            _STORAGE_LOCATION = STORAGE_LOCATION
            _BATCH = BATCH
            _QUANTITY = QUANTITY
            _CREATED_DATE = CREATED_DATE
            '_CREATE_TIME = CREATE_TIME
            _VC_COMMON1 = VC_COMMON1
            _VC_COMMON2 = VC_COMMON2
            _VC_COMMON3 = VC_COMMON3
            _VC_COMMON4 = VC_COMMON4
            _VC_COMMON5 = VC_COMMON5
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
End Class
