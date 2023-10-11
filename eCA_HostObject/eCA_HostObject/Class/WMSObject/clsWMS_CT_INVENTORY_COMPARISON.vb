﻿Public Class clsWMS_CT_INVENTORY_COMPARISON
  Private ShareName As String = "WMS_CT_INVENTORY_COMPARISON"
  Private ShareKey As String = ""
  Private _gid As String
  Private _SKU_NO As String '貨品ID

  Private _LOT_NO As String '批號

  Private _ITEM_COMMON1 As String '條件1

  Private _ITEM_COMMON2 As String '條件2

  Private _ITEM_COMMON3 As String '條件3

  Private _ITEM_COMMON4 As String '條件4

  Private _ITEM_COMMON5 As String '條件5

  Private _ITEM_COMMON6 As String '條件6

  Private _ITEM_COMMON7 As String '條件7

  Private _ITEM_COMMON8 As String '條件8

  Private _ITEM_COMMON9 As String '條件9

  Private _ITEM_COMMON10 As String '條件10

  Private _SORT_ITEM_COMMON1 As String '優先選擇條件1

  Private _SORT_ITEM_COMMON2 As String '優先選擇條件2

  Private _SORT_ITEM_COMMON3 As String '優先選擇條件3

  Private _SORT_ITEM_COMMON4 As String '優先選擇條件4

  Private _SORT_ITEM_COMMON5 As String '優先選擇條件5

  Private _OWNER_NO As String '貨主

  Private _SUB_OWNER_NO As String '子貨主

  Private _STORAGE_TYPE As Double '是否為暫存品Store 一般品=1Temporary 暫存品=2

  Private _BND As Double '保稅0:不保稅1:保稅

  Private _QC_STATUS As Double 'QC判定狀態NA=0OK=1NG=2LOCK=3

  Private _WMS_STOCK_QTY As Double 'WMS庫存數量

  Private _WMS_UNFINISH_QTY As Double '尚未出庫完成數

  Private _WMS_COMPARSON_QTY As Double 'WMS庫存數量 扣 尚未出庫完成數

  Private _ERP_STOCK_QTY As Double 'ERP 系統數量

  Private _ERP_UNFINISH_QTY As Double 'ERP 尚未出庫完成數

  Private _ERP_COMPARSON_QTY As Double 'ERP 庫存數量 扣 尚未出庫完成數

  Private _QUANTITY_VARIANCE As Double '庫存差異(WMS COMPARSON - ERP COMPARSON)

  Private _ERP_SYSTEM As String '上位系統名稱(若無區分則為ERP)

  Private _CREATE_TIME As String '建立時間

  Private _ACC_COMMON1 As String '備用欄位

  Private _ACC_COMMON2 As String '備用欄位

  Private _ACC_COMMON3 As String '備用欄位

  Private _ACC_COMMON4 As String '備用欄位

  Private _ACC_COMMON5 As String '備用欄位

  Private _ACC_COMMON6 As String '備用欄位

  Private _ACC_COMMON7 As String '備用欄位

  Private _ACC_COMMON8 As String '備用欄位

  Private _ACC_COMMON9 As String '備用欄位

  Private _ACC_COMMON10 As String '備用欄位

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
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
  Public Property LOT_NO() As String
    Get
      Return _LOT_NO
    End Get
    Set(ByVal value As String)
      _LOT_NO = value
    End Set
  End Property
  Public Property ITEM_COMMON1() As String
    Get
      Return _ITEM_COMMON1
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON1 = value
    End Set
  End Property
  Public Property ITEM_COMMON2() As String
    Get
      Return _ITEM_COMMON2
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON2 = value
    End Set
  End Property
  Public Property ITEM_COMMON3() As String
    Get
      Return _ITEM_COMMON3
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON3 = value
    End Set
  End Property
  Public Property ITEM_COMMON4() As String
    Get
      Return _ITEM_COMMON4
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON4 = value
    End Set
  End Property
  Public Property ITEM_COMMON5() As String
    Get
      Return _ITEM_COMMON5
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON5 = value
    End Set
  End Property
  Public Property ITEM_COMMON6() As String
    Get
      Return _ITEM_COMMON6
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON6 = value
    End Set
  End Property
  Public Property ITEM_COMMON7() As String
    Get
      Return _ITEM_COMMON7
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON7 = value
    End Set
  End Property
  Public Property ITEM_COMMON8() As String
    Get
      Return _ITEM_COMMON8
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON8 = value
    End Set
  End Property
  Public Property ITEM_COMMON9() As String
    Get
      Return _ITEM_COMMON9
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON9 = value
    End Set
  End Property
  Public Property ITEM_COMMON10() As String
    Get
      Return _ITEM_COMMON10
    End Get
    Set(ByVal value As String)
      _ITEM_COMMON10 = value
    End Set
  End Property
  Public Property SORT_ITEM_COMMON1() As String
    Get
      Return _SORT_ITEM_COMMON1
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON1 = value
    End Set
  End Property
  Public Property SORT_ITEM_COMMON2() As String
    Get
      Return _SORT_ITEM_COMMON2
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON2 = value
    End Set
  End Property
  Public Property SORT_ITEM_COMMON3() As String
    Get
      Return _SORT_ITEM_COMMON3
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON3 = value
    End Set
  End Property
  Public Property SORT_ITEM_COMMON4() As String
    Get
      Return _SORT_ITEM_COMMON4
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON4 = value
    End Set
  End Property
  Public Property SORT_ITEM_COMMON5() As String
    Get
      Return _SORT_ITEM_COMMON5
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON5 = value
    End Set
  End Property
  Public Property OWNER_NO() As String
    Get
      Return _OWNER_NO
    End Get
    Set(ByVal value As String)
      _OWNER_NO = value
    End Set
  End Property
  Public Property SUB_OWNER_NO() As String
    Get
      Return _SUB_OWNER_NO
    End Get
    Set(ByVal value As String)
      _SUB_OWNER_NO = value
    End Set
  End Property
  Public Property STORAGE_TYPE() As Double
    Get
      Return _STORAGE_TYPE
    End Get
    Set(ByVal value As Double)
      _STORAGE_TYPE = value
    End Set
  End Property
  Public Property BND() As Double
    Get
      Return _BND
    End Get
    Set(ByVal value As Double)
      _BND = value
    End Set
  End Property
  Public Property QC_STATUS() As Double
    Get
      Return _QC_STATUS
    End Get
    Set(ByVal value As Double)
      _QC_STATUS = value
    End Set
  End Property
  Public Property WMS_STOCK_QTY() As Double
    Get
      Return _WMS_STOCK_QTY
    End Get
    Set(ByVal value As Double)
      _WMS_STOCK_QTY = value
    End Set
  End Property
  Public Property WMS_UNFINISH_QTY() As Double
    Get
      Return _WMS_UNFINISH_QTY
    End Get
    Set(ByVal value As Double)
      _WMS_UNFINISH_QTY = value
    End Set
  End Property
  Public Property WMS_COMPARSON_QTY() As Double
    Get
      Return _WMS_COMPARSON_QTY
    End Get
    Set(ByVal value As Double)
      _WMS_COMPARSON_QTY = value
    End Set
  End Property
  Public Property ERP_STOCK_QTY() As Double
    Get
      Return _ERP_STOCK_QTY
    End Get
    Set(ByVal value As Double)
      _ERP_STOCK_QTY = value
    End Set
  End Property
  Public Property ERP_UNFINISH_QTY() As Double
    Get
      Return _ERP_UNFINISH_QTY
    End Get
    Set(ByVal value As Double)
      _ERP_UNFINISH_QTY = value
    End Set
  End Property
  Public Property ERP_COMPARSON_QTY() As Double
    Get
      Return _ERP_COMPARSON_QTY
    End Get
    Set(ByVal value As Double)
      _ERP_COMPARSON_QTY = value
    End Set
  End Property
  Public Property QUANTITY_VARIANCE() As Double
    Get
      Return _QUANTITY_VARIANCE
    End Get
    Set(ByVal value As Double)
      _QUANTITY_VARIANCE = value
    End Set
  End Property
  Public Property ERP_SYSTEM() As String
    Get
      Return _ERP_SYSTEM
    End Get
    Set(ByVal value As String)
      _ERP_SYSTEM = value
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
  Public Property ACC_COMMON1() As String
    Get
      Return _ACC_COMMON1
    End Get
    Set(ByVal value As String)
      _ACC_COMMON1 = value
    End Set
  End Property
  Public Property ACC_COMMON2() As String
    Get
      Return _ACC_COMMON2
    End Get
    Set(ByVal value As String)
      _ACC_COMMON2 = value
    End Set
  End Property
  Public Property ACC_COMMON3() As String
    Get
      Return _ACC_COMMON3
    End Get
    Set(ByVal value As String)
      _ACC_COMMON3 = value
    End Set
  End Property
  Public Property ACC_COMMON4() As String
    Get
      Return _ACC_COMMON4
    End Get
    Set(ByVal value As String)
      _ACC_COMMON4 = value
    End Set
  End Property
  Public Property ACC_COMMON5() As String
    Get
      Return _ACC_COMMON5
    End Get
    Set(ByVal value As String)
      _ACC_COMMON5 = value
    End Set
  End Property
  Public Property ACC_COMMON6() As String
    Get
      Return _ACC_COMMON6
    End Get
    Set(ByVal value As String)
      _ACC_COMMON6 = value
    End Set
  End Property
  Public Property ACC_COMMON7() As String
    Get
      Return _ACC_COMMON7
    End Get
    Set(ByVal value As String)
      _ACC_COMMON7 = value
    End Set
  End Property
  Public Property ACC_COMMON8() As String
    Get
      Return _ACC_COMMON8
    End Get
    Set(ByVal value As String)
      _ACC_COMMON8 = value
    End Set
  End Property
  Public Property ACC_COMMON9() As String
    Get
      Return _ACC_COMMON9
    End Get
    Set(ByVal value As String)
      _ACC_COMMON9 = value
    End Set
  End Property
  Public Property ACC_COMMON10() As String
    Get
      Return _ACC_COMMON10
    End Get
    Set(ByVal value As String)
      _ACC_COMMON10 = value
    End Set
  End Property

  Public Sub New(ByVal SKU_NO As String, ByVal LOT_NO As String, ByVal ITEM_COMMON1 As String, ByVal ITEM_COMMON2 As String, ByVal ITEM_COMMON3 As String,
                 ByVal ITEM_COMMON4 As String, ByVal ITEM_COMMON5 As String, ByVal ITEM_COMMON6 As String, ByVal ITEM_COMMON7 As String,
                 ByVal ITEM_COMMON8 As String, ByVal ITEM_COMMON9 As String, ByVal ITEM_COMMON10 As String, ByVal SORT_ITEM_COMMON1 As String,
                 ByVal SORT_ITEM_COMMON2 As String, ByVal SORT_ITEM_COMMON3 As String, ByVal SORT_ITEM_COMMON4 As String, ByVal SORT_ITEM_COMMON5 As String,
                 ByVal OWNER_NO As String, ByVal SUB_OWNER_NO As String, ByVal STORAGE_TYPE As Double, ByVal BND As Double, ByVal QC_STATUS As Double,
                 ByVal WMS_STOCK_QTY As Double, ByVal WMS_UNFINISH_QTY As Double, ByVal WMS_COMPARSON_QTY As Double, ByVal ERP_STOCK_QTY As Double,
                 ByVal ERP_UNFINISH_QTY As Double, ByVal ERP_COMPARSON_QTY As Double, ByVal QUANTITY_VARIANCE As Double, ByVal ERP_SYSTEM As String,
                 ByVal CREATE_TIME As String, ByVal ACC_COMMON1 As String, ByVal ACC_COMMON2 As String, ByVal ACC_COMMON3 As String, ByVal ACC_COMMON4 As String,
                 ByVal ACC_COMMON5 As String, ByVal ACC_COMMON6 As String, ByVal ACC_COMMON7 As String, ByVal ACC_COMMON8 As String, ByVal ACC_COMMON9 As String,
                 ByVal ACC_COMMON10 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(SKU_NO, LOT_NO, OWNER_NO, SUB_OWNER_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3)
      _gid = key
      _SKU_NO = SKU_NO
      _LOT_NO = LOT_NO
      _ITEM_COMMON1 = ITEM_COMMON1
      _ITEM_COMMON2 = ITEM_COMMON2
      _ITEM_COMMON3 = ITEM_COMMON3
      _ITEM_COMMON4 = ITEM_COMMON4
      _ITEM_COMMON5 = ITEM_COMMON5
      _ITEM_COMMON6 = ITEM_COMMON6
      _ITEM_COMMON7 = ITEM_COMMON7
      _ITEM_COMMON8 = ITEM_COMMON8
      _ITEM_COMMON9 = ITEM_COMMON9
      _ITEM_COMMON10 = ITEM_COMMON10
      _SORT_ITEM_COMMON1 = SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = SORT_ITEM_COMMON5
      _OWNER_NO = OWNER_NO
      _SUB_OWNER_NO = SUB_OWNER_NO
      _STORAGE_TYPE = STORAGE_TYPE
      _BND = BND
      _QC_STATUS = QC_STATUS
      _WMS_STOCK_QTY = WMS_STOCK_QTY
      _WMS_UNFINISH_QTY = WMS_UNFINISH_QTY
      _WMS_COMPARSON_QTY = WMS_COMPARSON_QTY
      _ERP_STOCK_QTY = ERP_STOCK_QTY
      _ERP_UNFINISH_QTY = ERP_UNFINISH_QTY
      _ERP_COMPARSON_QTY = ERP_COMPARSON_QTY
      _QUANTITY_VARIANCE = QUANTITY_VARIANCE
      _ERP_SYSTEM = ERP_SYSTEM
      _CREATE_TIME = CREATE_TIME
      _ACC_COMMON1 = ACC_COMMON1
      _ACC_COMMON2 = ACC_COMMON2
      _ACC_COMMON3 = ACC_COMMON3
      _ACC_COMMON4 = ACC_COMMON4
      _ACC_COMMON5 = ACC_COMMON5
      _ACC_COMMON6 = ACC_COMMON6
      _ACC_COMMON7 = ACC_COMMON7
      _ACC_COMMON8 = ACC_COMMON8
      _ACC_COMMON9 = ACC_COMMON9
      _ACC_COMMON10 = ACC_COMMON10
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
  Public Function Clone() As clsWMS_CT_INVENTORY_COMPARISON
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function Get_Combination_Key(ByVal SKU_NO As String, ByVal LOT_NO As String, ByVal ITEM_COMMON1 As String, ByVal ITEM_COMMON2 As String, ByVal ITEM_COMMON3 As String,
                                             ByVal OWNER_NO As String, ByVal SUB_OWNER_NO As String) As String
    Try
      Dim key As String = SKU_NO & LinkKey & LOT_NO & LinkKey & OWNER_NO & LinkKey & SUB_OWNER_NO & LinkKey & ITEM_COMMON1 & LinkKey & ITEM_COMMON2 & LinkKey & ITEM_COMMON3
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String),
                                         ByRef lstQueueSql As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_CT_INVENTORY_COMPARISONManagement.GetInsertSQL(Me)
      lstQueueSql.Add(strSQL)

      Dim HIST_TIME = GetNewTime_DBFormat()
      Dim objTmp As New clsWMS_CH_INVENTORY_COMPARISON(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, WMS_UNFINISH_QTY, WMS_COMPARSON_QTY, ERP_STOCK_QTY, ERP_UNFINISH_QTY, ERP_COMPARSON_QTY, QUANTITY_VARIANCE, ERP_SYSTEM, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, ACC_COMMON6, ACC_COMMON7, ACC_COMMON8, ACC_COMMON9, ACC_COMMON10, HIST_TIME)
      strSQL = WMS_CH_INVENTORY_COMPARISONManagement.GetInsertSQL(objTmp)
      lstQueueSql.Add(strSQL)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_CT_INVENTORY_COMPARISONManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_CT_INVENTORY_COMPARISONManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_CT_INVENTORY_COMPARISON As clsWMS_CT_INVENTORY_COMPARISON) As Boolean
    Try
      Dim key As String = objWMS_CT_INVENTORY_COMPARISON._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _SKU_NO = SKU_NO
      _LOT_NO = LOT_NO
      _ITEM_COMMON1 = ITEM_COMMON1
      _ITEM_COMMON2 = ITEM_COMMON2
      _ITEM_COMMON3 = ITEM_COMMON3
      _ITEM_COMMON4 = ITEM_COMMON4
      _ITEM_COMMON5 = ITEM_COMMON5
      _ITEM_COMMON6 = ITEM_COMMON6
      _ITEM_COMMON7 = ITEM_COMMON7
      _ITEM_COMMON8 = ITEM_COMMON8
      _ITEM_COMMON9 = ITEM_COMMON9
      _ITEM_COMMON10 = ITEM_COMMON10
      _SORT_ITEM_COMMON1 = SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = SORT_ITEM_COMMON5
      _OWNER_NO = OWNER_NO
      _SUB_OWNER_NO = SUB_OWNER_NO
      _STORAGE_TYPE = STORAGE_TYPE
      _BND = BND
      _QC_STATUS = QC_STATUS
      _WMS_STOCK_QTY = WMS_STOCK_QTY
      _WMS_UNFINISH_QTY = WMS_UNFINISH_QTY
      _WMS_COMPARSON_QTY = WMS_COMPARSON_QTY
      _ERP_STOCK_QTY = ERP_STOCK_QTY
      _ERP_UNFINISH_QTY = ERP_UNFINISH_QTY
      _ERP_COMPARSON_QTY = ERP_COMPARSON_QTY
      _QUANTITY_VARIANCE = QUANTITY_VARIANCE
      _ERP_SYSTEM = ERP_SYSTEM
      _CREATE_TIME = CREATE_TIME
      _ACC_COMMON1 = ACC_COMMON1
      _ACC_COMMON2 = ACC_COMMON2
      _ACC_COMMON3 = ACC_COMMON3
      _ACC_COMMON4 = ACC_COMMON4
      _ACC_COMMON5 = ACC_COMMON5
      _ACC_COMMON6 = ACC_COMMON6
      _ACC_COMMON7 = ACC_COMMON7
      _ACC_COMMON8 = ACC_COMMON8
      _ACC_COMMON9 = ACC_COMMON9
      _ACC_COMMON10 = ACC_COMMON10
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
