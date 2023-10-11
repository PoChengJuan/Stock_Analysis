Public Class clsWMS_CT_PRODUCTION_REPORT
  Private ShareName As String = "WMS_CT_PRODUCTION_REPORT"
  Private ShareKey As String = ""
  Private _gid As String
  Private _FACTORY_NO As String '廠別

  Private _AREA_NO As String '生產線區塊

  Private _PO_ID As String '製令單號

  Private _SKU_NO As String '貨品編號

  Private _REPORT_STATUS As Double '上报狀態0:Queued1:Filed2:Success

  Private _QTY As Double '上报的数量 (人工输入)

  Private _QTY_NG As Double '上报NG的數量 (人工输入)

  Private _REPORT_QTY As Double '实际上报数量

  Private _REPORT_QTY_NG As Double '实际上報NG的數量

  Private _TB003 As String '移轉日期

  Private _TB004 As String '移出類別

  Private _TB005 As String '移出部門

  Private _TB008 As String '移入部門

  Private _TB007 As String '移入類別

  Private _TB010 As String '廠別代號

  Private _TC003 As String '加工順序

  Private _TC004 As String '製令單別

  Private _TC005 As String '製令單號

  Private _TC006 As String '移出工序

  Private _TC007 As String '移出製程

  Private _TC008 As String '移入工序

  Private _TC009 As String '移入製程

  Private _TC010 As String '單位

  Private _TC014 As Double '驗收數量

  Private _TC016 As Double '報廢數量

  Private _TC020 As String '使用人時

  Private _TC021 As String '使用機時

  Private _TC200 As String '人數

  Private _TC201 As String '當班主管

  Private _CREATE_TIME As String '建立時間

  Private _UPDATE_TIME As String '開始時間

  Private _FINISH_TIME As String '完成时间

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property FACTORY_NO() As String
    Get
      Return _FACTORY_NO
    End Get
    Set(ByVal value As String)
      _FACTORY_NO = value
    End Set
  End Property
  Public Property AREA_NO() As String
    Get
      Return _AREA_NO
    End Get
    Set(ByVal value As String)
      _AREA_NO = value
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
  Public Property SKU_NO() As String
    Get
      Return _SKU_NO
    End Get
    Set(ByVal value As String)
      _SKU_NO = value
    End Set
  End Property
  Public Property REPORT_STATUS() As Double
    Get
      Return _REPORT_STATUS
    End Get
    Set(ByVal value As Double)
      _REPORT_STATUS = value
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
  Public Property QTY_NG() As Double
    Get
      Return _QTY_NG
    End Get
    Set(ByVal value As Double)
      _QTY_NG = value
    End Set
  End Property
  Public Property REPORT_QTY() As Double
    Get
      Return _REPORT_QTY
    End Get
    Set(ByVal value As Double)
      _REPORT_QTY = value
    End Set
  End Property
  Public Property REPORT_QTY_NG() As Double
    Get
      Return _REPORT_QTY_NG
    End Get
    Set(ByVal value As Double)
      _REPORT_QTY_NG = value
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
  Public Property TB004() As String
    Get
      Return _TB004
    End Get
    Set(ByVal value As String)
      _TB004 = value
    End Set
  End Property
  Public Property TB005() As String
    Get
      Return _TB005
    End Get
    Set(ByVal value As String)
      _TB005 = value
    End Set
  End Property
  Public Property TB008() As String
    Get
      Return _TB008
    End Get
    Set(ByVal value As String)
      _TB008 = value
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
  Public Property TB010() As String
    Get
      Return _TB010
    End Get
    Set(ByVal value As String)
      _TB010 = value
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
  Public Property TC004() As String
    Get
      Return _TC004
    End Get
    Set(ByVal value As String)
      _TC004 = value
    End Set
  End Property
  Public Property TC005() As String
    Get
      Return _TC005
    End Get
    Set(ByVal value As String)
      _TC005 = value
    End Set
  End Property
  Public Property TC006() As String
    Get
      Return _TC006
    End Get
    Set(ByVal value As String)
      _TC006 = value
    End Set
  End Property
  Public Property TC007() As String
    Get
      Return _TC007
    End Get
    Set(ByVal value As String)
      _TC007 = value
    End Set
  End Property
  Public Property TC008() As String
    Get
      Return _TC008
    End Get
    Set(ByVal value As String)
      _TC008 = value
    End Set
  End Property
  Public Property TC009() As String
    Get
      Return _TC009
    End Get
    Set(ByVal value As String)
      _TC009 = value
    End Set
  End Property
  Public Property TC010() As String
    Get
      Return _TC010
    End Get
    Set(ByVal value As String)
      _TC010 = value
    End Set
  End Property
  Public Property TC014() As Double
    Get
      Return _TC014
    End Get
    Set(ByVal value As Double)
      _TC014 = value
    End Set
  End Property
  Public Property TC016() As Double
    Get
      Return _TC016
    End Get
    Set(ByVal value As Double)
      _TC016 = value
    End Set
  End Property
  Public Property TC020() As String
    Get
      Return _TC020
    End Get
    Set(ByVal value As String)
      _TC020 = value
    End Set
  End Property
  Public Property TC021() As String
    Get
      Return _TC021
    End Get
    Set(ByVal value As String)
      _TC021 = value
    End Set
  End Property
  Public Property TC200() As String
    Get
      Return _TC200
    End Get
    Set(ByVal value As String)
      _TC200 = value
    End Set
  End Property
  Public Property TC201() As String
    Get
      Return _TC201
    End Get
    Set(ByVal value As String)
      _TC201 = value
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
  Public Property UPDATE_TIME() As String
    Get
      Return _UPDATE_TIME
    End Get
    Set(ByVal value As String)
      _UPDATE_TIME = value
    End Set
  End Property
  Public Property FINISH_TIME() As String
    Get
      Return _FINISH_TIME
    End Get
    Set(ByVal value As String)
      _FINISH_TIME = value
    End Set
  End Property

  Public Sub New(ByVal FACTORY_NO As String, ByVal AREA_NO As String, ByVal PO_ID As String, ByVal SKU_NO As String, ByVal REPORT_STATUS As Double, ByVal QTY As Double, ByVal QTY_NG As Double, ByVal REPORT_QTY As Double, ByVal REPORT_QTY_NG As Double, ByVal TB003 As String, ByVal TB004 As String, ByVal TB005 As String, ByVal TB008 As String, ByVal TB007 As String, ByVal TB010 As String, ByVal TC003 As String, ByVal TC004 As String, ByVal TC005 As String, ByVal TC006 As String, ByVal TC007 As String, ByVal TC008 As String, ByVal TC009 As String, ByVal TC010 As String, ByVal TC014 As Double, ByVal TC016 As Double, ByVal TC020 As String, ByVal TC021 As String, ByVal TC200 As String, ByVal TC201 As String, ByVal CREATE_TIME As String, ByVal UPDATE_TIME As String, ByVal FINISH_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(FACTORY_NO, AREA_NO, PO_ID, SKU_NO)
      _gid = key
      _FACTORY_NO = FACTORY_NO
      _AREA_NO = AREA_NO
      _PO_ID = PO_ID
      _SKU_NO = SKU_NO
      _REPORT_STATUS = REPORT_STATUS
      _QTY = QTY
      _QTY_NG = QTY_NG
      _REPORT_QTY = REPORT_QTY
      _REPORT_QTY_NG = REPORT_QTY_NG
      _TB003 = TB003
      _TB004 = TB004
      _TB005 = TB005
      _TB008 = TB008
      _TB007 = TB007
      _TB010 = TB010
      _TC003 = TC003
      _TC004 = TC004
      _TC005 = TC005
      _TC006 = TC006
      _TC007 = TC007
      _TC008 = TC008
      _TC009 = TC009
      _TC010 = TC010
      _TC014 = TC014
      _TC016 = TC016
      _TC020 = TC020
      _TC021 = TC021
      _TC200 = TC200
      _TC201 = TC201
      _CREATE_TIME = CREATE_TIME
      _UPDATE_TIME = UPDATE_TIME
      _FINISH_TIME = FINISH_TIME
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
  Public Shared Function Get_Combination_Key(ByVal FACTORY_NO As String, ByVal AREA_NO As String, ByVal PO_ID As String, ByVal SKU_NO As String) As String
    Try
      Dim key As String = FACTORY_NO & LinkKey & AREA_NO & LinkKey & PO_ID & LinkKey & SKU_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsWMS_CT_PRODUCTION_REPORT
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
      Dim strSQL As String = WMS_CT_PRODUCTION_REPORTManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_CT_PRODUCTION_REPORTManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_CT_PRODUCTION_REPORTManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_CT_PRODUCTION_REPORT As clsWMS_CT_PRODUCTION_REPORT) As Boolean
    Try
      Dim key As String = objWMS_CT_PRODUCTION_REPORT._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _FACTORY_NO = FACTORY_NO
      _AREA_NO = AREA_NO
      _PO_ID = PO_ID
      _SKU_NO = SKU_NO
      _REPORT_STATUS = REPORT_STATUS
      _QTY = QTY
      _QTY_NG = QTY_NG
      _REPORT_QTY = REPORT_QTY
      _REPORT_QTY_NG = REPORT_QTY_NG
      _TB003 = TB003
      _TB004 = TB004
      _TB005 = TB005
      _TB008 = TB008
      _TB007 = TB007
      _TB010 = TB010
      _TC003 = TC003
      _TC004 = TC004
      _TC005 = TC005
      _TC006 = TC006
      _TC007 = TC007
      _TC008 = TC008
      _TC009 = TC009
      _TC010 = TC010
      _TC014 = TC014
      _TC016 = TC016
      _TC020 = TC020
      _TC021 = TC021
      _TC200 = TC200
      _TC201 = TC201
      _CREATE_TIME = CREATE_TIME
      _UPDATE_TIME = UPDATE_TIME
      _FINISH_TIME = FINISH_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
