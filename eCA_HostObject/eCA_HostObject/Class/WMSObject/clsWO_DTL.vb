Imports System.Collections.Concurrent


Public Class clsWO_DTL
  Private ShareName As String = "WO_DTL"
  Private ShareKey As String = ""

  Private _gid As String
  Private _WO_ID As String '工單單編號
  Private _WO_Serial_NO As String '工單明細編號
  Private _WORKING_TYPE As enuWorkingType '加工類型預設為0
  Private _WORKING_SERIAL_NO As String '加工單項次
  Private _WORKING_SERIAL_SEQ As String '加工單項次順序
  Private _QC_Method As enuQCMethod 'QC的方式
  Private _SKU_No As String '貨品編號
  Private _SKU_Catalog As enuSKU_CATALOG '貨品類型
  Private _Lot_NO As String '批號
  Private _QTY As Decimal '需求量
  Private _QTY_Transferred As Decimal '已入庫/出庫數量
  Private _QTY_Process As Decimal '已挷定數量
  Private _QTY_Replenishment As Decimal '已經產生的補貨數量
  Private _QTY_Abort As Decimal '中止揀貨數量
  Private _Carrier_ID As String '載具編號
  Private _Package_ID As String '箱ID/包裝ID
  Private _Item_Common1 As String '條件1
  Private _Item_Common2 As String '條件2
  Private _Item_Common3 As String '條件3
  Private _Item_Common4 As String '條件4
  Private _Item_Common5 As String '條件5
  Private _Item_Common6 As String '條件6
  Private _Item_Common7 As String '條件7
  Private _Item_Common8 As String '條件8
  Private _Item_Common9 As String '條件9
  Private _Item_Common10 As String '條件10
  Private _SORT_ITEM_COMMON1 As String
  Private _SORT_ITEM_COMMON2 As String
  Private _SORT_ITEM_COMMON3 As String
  Private _SORT_ITEM_COMMON4 As String
  Private _SORT_ITEM_COMMON5 As String
  Private _Comments As String '備註
  Private _SL_NO As String  'ERP暫存地點
  Private _EXPIRED_DATE As String
  Private _Storage_Type As enuStorageType
  Private _BND As Boolean
  Private _QC_Status As String
  Private _FROM_OWNER_No As String '原貨主編號
  Private _FROM_SUB_OWNER_No As String '原子貨主編號
  Private _TO_OWNER_No As String '新子貨主編號
  Private _TO_SUB_OWNER_No As String '新貨主編號
  Private _FACTORY_No As String '廠區(廠別)(預先指定出入庫區域)
  Private _DEST_AREA_No As String '倉庫編號(預先指定出入庫區域)
  Private _DEST_LOCATION_No As String '儲位編號(預先指定出入庫區域)
  Private _SOURCE_AREA_No As String '來源倉庫編號(指定從哪個區域進入的貨)
  Private _SOURCE_LOCATION_NO As String '來源儲位編號(指定從哪個區域進入的貨)
  Private _Start_Time As String
  Private _Start_Transfer_Time As String
  Private _Finish_Transfer_Time As String
  Private _Finish_Time As String
  Private _WO_DTL_DC_Status As enuWO_DTL_DC_Status
  Private _DTL_Common1 As String
  Private _DTL_Common2 As String
  Private _DTL_Common3 As String
  Private _DTL_Common4 As String
  Private _DTL_Common5 As String

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property WO_ID() As String
    Get
      Return _WO_ID
    End Get
    Set(ByVal value As String)
      _WO_ID = value
    End Set
  End Property
  Public Property WO_Serial_No() As String
    Get
      Return _WO_Serial_NO
    End Get
    Set(ByVal value As String)
      _WO_Serial_NO = value
    End Set
  End Property
  Public Property WORKING_TYPE() As enuWorkingType
    Get
      Return _WORKING_TYPE
    End Get
    Set(ByVal value As enuWorkingType)
      _WORKING_TYPE = value
    End Set
  End Property
  Public Property WORKING_SERIAL_NO() As String
    Get
      Return _WORKING_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _WORKING_SERIAL_NO = value
    End Set
  End Property
  Public Property WORKING_SERIAL_SEQ() As String
    Get
      Return _WORKING_SERIAL_SEQ
    End Get
    Set(ByVal value As String)
      _WORKING_SERIAL_SEQ = value
    End Set
  End Property
  Public Property QC_Method() As enuQCMethod
    Get
      Return _QC_Method
    End Get
    Set(ByVal value As enuQCMethod)
      _QC_Method = value
    End Set
  End Property
  Public Property SKU_No() As String
    Get
      Return _SKU_No
    End Get
    Set(ByVal value As String)
      _SKU_No = value
    End Set
  End Property
  Public Property SKU_Catalog() As enuSKU_CATALOG
    Get
      Return _SKU_Catalog
    End Get
    Set(ByVal value As enuSKU_CATALOG)
      _SKU_Catalog = value
    End Set
  End Property
  Public Property Lot_No() As String
    Get
      Return _Lot_NO
    End Get
    Set(ByVal value As String)
      _Lot_NO = value
    End Set
  End Property
  Public Property QTY() As Decimal
    Get
      Return _QTY
    End Get
    Set(ByVal value As Decimal)
      _QTY = value
    End Set
  End Property
  Public Property QTY_Transferred() As Decimal
    Get
      Return _QTY_Transferred
    End Get
    Set(ByVal value As Decimal)
      _QTY_Transferred = value
    End Set
  End Property
  Public Property QTY_Process() As Decimal
    Get
      Return _QTY_Process
    End Get
    Set(ByVal value As Decimal)
      _QTY_Process = value
    End Set
  End Property
  Public Property QTY_Replenishment() As Decimal
    Get
      Return _QTY_Replenishment
    End Get
    Set(ByVal value As Decimal)
      _QTY_Replenishment = value
    End Set
  End Property
  Public Property QTY_Abort() As Decimal
    Get
      Return _QTY_Abort
    End Get
    Set(ByVal value As Decimal)
      _QTY_Abort = value
    End Set
  End Property
  Public Property Carrier_ID() As String
    Get
      Return _Carrier_ID
    End Get
    Set(ByVal value As String)
      _Carrier_ID = value
    End Set
  End Property
  Public Property Package_ID() As String
    Get
      Return _Package_ID
    End Get
    Set(ByVal value As String)
      _Package_ID = value
    End Set
  End Property
  Public Property Item_Common1() As String
    Get
      Return _Item_Common1
    End Get
    Set(ByVal value As String)
      _Item_Common1 = value
    End Set
  End Property
  Public Property Item_Common2() As String
    Get
      Return _Item_Common2
    End Get
    Set(ByVal value As String)
      _Item_Common2 = value
    End Set
  End Property
  Public Property Item_Common3() As String
    Get
      Return _Item_Common3
    End Get
    Set(ByVal value As String)
      _Item_Common3 = value
    End Set
  End Property
  Public Property Item_Common4() As String
    Get
      Return _Item_Common4
    End Get
    Set(ByVal value As String)
      _Item_Common4 = value
    End Set
  End Property
  Public Property Item_Common5() As String
    Get
      Return _Item_Common5
    End Get
    Set(ByVal value As String)
      _Item_Common5 = value
    End Set
  End Property
  Public Property Item_Common6() As String
    Get
      Return _Item_Common6
    End Get
    Set(ByVal value As String)
      _Item_Common6 = value
    End Set
  End Property
  Public Property Item_Common7() As String
    Get
      Return _Item_Common7
    End Get
    Set(ByVal value As String)
      _Item_Common7 = value
    End Set
  End Property
  Public Property Item_Common8() As String
    Get
      Return _Item_Common8
    End Get
    Set(ByVal value As String)
      _Item_Common8 = value
    End Set
  End Property
  Public Property Item_Common9() As String
    Get
      Return _Item_Common9
    End Get
    Set(ByVal value As String)
      _Item_Common9 = value
    End Set
  End Property
  Public Property Item_Common10() As String
    Get
      Return _Item_Common10
    End Get
    Set(ByVal value As String)
      _Item_Common10 = value
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
  Public Property Comments() As String
    Get
      Return _Comments
    End Get
    Set(ByVal value As String)
      _Comments = value
    End Set
  End Property
  Public Property EXPIRED_DATE() As String
    Get
      Return _EXPIRED_DATE
    End Get
    Set(ByVal value As String)
      _EXPIRED_DATE = value
    End Set
  End Property
  Public Property SL_NO() As String
    Get
      Return _SL_NO
    End Get
    Set(ByVal value As String)
      _SL_NO = value
    End Set
  End Property
  Public Property Storage_Type() As enuStorageType
    Get
      Return _Storage_Type
    End Get
    Set(ByVal value As enuStorageType)
      _Storage_Type = value
    End Set
  End Property
  Public Property BND() As Boolean
    Get
      Return _BND
    End Get
    Set(ByVal value As Boolean)
      _BND = value
    End Set
  End Property
  Public Property QC_Status() As String
    Get
      Return _QC_Status
    End Get
    Set(ByVal value As String)
      _QC_Status = value
    End Set
  End Property
  Public Property FROM_OWNER_No() As String
    Get
      Return _FROM_OWNER_No
    End Get
    Set(ByVal value As String)
      _FROM_OWNER_No = value
    End Set
  End Property
  Public Property FROM_SUB_OWNER_No() As String
    Get
      Return _FROM_SUB_OWNER_No
    End Get
    Set(ByVal value As String)
      _FROM_SUB_OWNER_No = value
    End Set
  End Property
  Public Property TO_OWNER_No() As String
    Get
      Return _TO_OWNER_No
    End Get
    Set(ByVal value As String)
      _TO_OWNER_No = value
    End Set
  End Property
  Public Property TO_SUB_OWNER_No() As String
    Get
      Return _TO_SUB_OWNER_No
    End Get
    Set(ByVal value As String)
      _TO_SUB_OWNER_No = value
    End Set
  End Property
  Public Property FACTORY_No() As String
    Get
      Return _FACTORY_No
    End Get
    Set(ByVal value As String)
      _FACTORY_No = value
    End Set
  End Property
  Public Property DEST_AREA_No() As String
    Get
      Return _DEST_AREA_No
    End Get
    Set(ByVal value As String)
      _DEST_AREA_No = value
    End Set
  End Property
  Public Property DEST_LOCATION_No() As String
    Get
      Return _DEST_LOCATION_No
    End Get
    Set(ByVal value As String)
      _DEST_LOCATION_No = value
    End Set
  End Property
  Public Property SOURCE_AREA_No() As String
    Get
      Return _SOURCE_AREA_No
    End Get
    Set(ByVal value As String)
      _SOURCE_AREA_No = value
    End Set
  End Property
  Public Property SOURCE_LOCATION_NO() As String
    Get
      Return _SOURCE_LOCATION_NO
    End Get
    Set(ByVal value As String)
      _SOURCE_LOCATION_NO = value
    End Set
  End Property
  Public Property Start_Time() As String
    Get
      Return _Start_Time
    End Get
    Set(ByVal value As String)
      _Start_Time = value
    End Set
  End Property
  Public Property Start_Transfer_Time() As String
    Get
      Return _Start_Transfer_Time
    End Get
    Set(ByVal value As String)
      _Start_Transfer_Time = value
    End Set
  End Property
  Public Property Finish_Transfer_Time() As String
    Get
      Return _Finish_Transfer_Time
    End Get
    Set(ByVal value As String)
      _Finish_Transfer_Time = value
    End Set
  End Property
  Public Property Finish_Time() As String
    Get
      Return _Finish_Time
    End Get
    Set(ByVal value As String)
      _Finish_Time = value
    End Set
  End Property
  Public Property WO_DTL_DC_Status() As enuWO_DTL_DC_Status
    Get
      Return _WO_DTL_DC_Status
    End Get
    Set(ByVal value As enuWO_DTL_DC_Status)
      _WO_DTL_DC_Status = value
    End Set
  End Property
  Public Property DTL_Common1() As String
    Get
      Return _DTL_Common1
    End Get
    Set(ByVal value As String)
      _DTL_Common1 = value
    End Set
  End Property
  Public Property DTL_Common2() As String
    Get
      Return _DTL_Common2
    End Get
    Set(ByVal value As String)
      _DTL_Common2 = value
    End Set
  End Property
  Public Property DTL_Common3() As String
    Get
      Return _DTL_Common3
    End Get
    Set(ByVal value As String)
      _DTL_Common3 = value
    End Set
  End Property
  Public Property DTL_Common4() As String
    Get
      Return _DTL_Common4
    End Get
    Set(ByVal value As String)
      _DTL_Common4 = value
    End Set
  End Property
  Public Property DTL_Common5() As String
    Get
      Return _DTL_Common5
    End Get
    Set(ByVal value As String)
      _DTL_Common5 = value
    End Set
  End Property

  '物件建立時執行的事件(從資料庫抓資料建立物件)
  Public Sub New(ByVal WO_ID As String, ByVal WO_Serial_NO As String,
                 ByVal WORKING_TYPE As Long, ByVal WORKING_SERIAL_NO As String, ByVal WORKING_SERIAL_SEQ As String, ByVal QC_Method As enuQCMethod,
                 ByVal SKU_NO As String, ByVal SKU_Catalog As enuSKU_CATALOG, ByVal LOT_NO As String,
                 ByVal Qty As Decimal, ByVal Qty_Transferred As Decimal, ByVal Qty_Process As Decimal, ByVal QTY_Replenishment As Decimal, ByVal QTY_Abort As Decimal,
                 ByVal Carrier_ID As String, ByVal Package_ID As String,
                 ByVal Item_Common1 As String, ByVal Item_Common2 As String, ByVal Item_Common3 As String, ByVal Item_Common4 As String, ByVal Item_Common5 As String,
                 ByVal Item_Common6 As String, ByVal Item_Common7 As String, ByVal Item_Common8 As String, ByVal Item_Common9 As String, ByVal Item_Common10 As String,
                 ByVal SORT_ITEM_COMMON1 As String, ByVal SORT_ITEM_COMMON2 As String, ByVal SORT_ITEM_COMMON3 As String, ByVal SORT_ITEM_COMMON4 As String, ByVal SORT_ITEM_COMMON5 As String,
                 ByVal Comments As String, ByVal EXPIRED_DATE As String, ByVal SL_NO As String, ByVal Storage_Type As enuStorageType, ByVal BND As Boolean, ByVal QC_Status As String,
                 ByVal FROM_OWNER_No As String, ByVal FROM_SUB_OWNER_No As String, ByVal TO_OWNER_No As String, ByVal TO_SUB_OWNER_No As String,
                 ByVal FACTORY_No As String, ByVal DEST_AREA_No As String, ByVal DEST_LOCATION_No As String, ByVal SOURCE_AREA_NO As String, ByVal SOURCE_LOCATION_NO As String,
                 ByVal Start_Time As String, ByVal Start_Transfer_Time As String, ByVal Finish_Transfer_Time As String, ByVal Finish_Time As String, ByVal WO_DTL_DC_Status As enuWO_DTL_DC_Status,
                 ByVal DTL_COMMON1 As String, ByVal DTL_COMMON2 As String, ByVal DTL_COMMON3 As String, ByVal DTL_COMMON4 As String, ByVal DTL_COMMON5 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(WO_ID, WO_Serial_NO)
      _gid = key
      _WO_ID = WO_ID
      _WO_Serial_NO = WO_Serial_NO
      _WORKING_TYPE = WORKING_TYPE
      _WORKING_SERIAL_NO = WORKING_SERIAL_NO
      _WORKING_SERIAL_SEQ = WORKING_SERIAL_SEQ
      _QC_Method = QC_Method
      _SKU_No = SKU_NO
      _SKU_Catalog = SKU_Catalog
      _Lot_NO = LOT_NO
      _QTY = Qty
      _QTY_Transferred = Qty_Transferred
      _QTY_Process = Qty_Process
      _QTY_Replenishment = QTY_Replenishment
      _QTY_Abort = QTY_Abort
      _Carrier_ID = Carrier_ID
      _Package_ID = Package_ID
      _Item_Common1 = Item_Common1
      _Item_Common2 = Item_Common2
      _Item_Common3 = Item_Common3
      _Item_Common4 = Item_Common4
      _Item_Common5 = Item_Common5
      _Item_Common6 = Item_Common6
      _Item_Common7 = Item_Common7
      _Item_Common8 = Item_Common8
      _Item_Common9 = Item_Common9
      _Item_Common10 = Item_Common10
      _SORT_ITEM_COMMON1 = SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = SORT_ITEM_COMMON5
      _Comments = Comments
      _EXPIRED_DATE = EXPIRED_DATE
      _SL_NO = SL_NO
      _Storage_Type = Storage_Type
      _BND = BND
      _QC_Status = QC_Status
      _FROM_OWNER_No = FROM_OWNER_No
      _FROM_SUB_OWNER_No = FROM_SUB_OWNER_No
      _TO_OWNER_No = TO_OWNER_No
      _TO_SUB_OWNER_No = TO_SUB_OWNER_No
      _FACTORY_No = FACTORY_No
      _DEST_AREA_No = DEST_AREA_No
      _DEST_LOCATION_No = DEST_LOCATION_No
      _SOURCE_AREA_No = SOURCE_AREA_NO
      _SOURCE_LOCATION_NO = SOURCE_LOCATION_NO
      _Start_Time = Start_Time
      _Start_Transfer_Time = Start_Transfer_Time
      _Finish_Transfer_Time = Finish_Transfer_Time
      _Finish_Time = Finish_Time
      _WO_DTL_DC_Status = WO_DTL_DC_Status
      _DTL_Common1 = DTL_COMMON1
      _DTL_Common2 = DTL_COMMON2
      _DTL_Common3 = DTL_COMMON3
      _DTL_Common4 = DTL_COMMON4
      _DTL_Common5 = DTL_COMMON5
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '物件建立時執行的事件(建立工單使用)
  Public Sub New(ByVal WO_ID As String, ByVal WO_Serial_NO As String,
                 ByVal SKU_NO As String, ByVal LOT_NO As String,
                 ByVal Qty As Decimal,
                 ByVal Carrier_ID As String, ByVal Package_ID As String,
                 ByVal Item_Common1 As String, ByVal Item_Common2 As String, ByVal Item_Common3 As String, ByVal Item_Common4 As String, ByVal Item_Common5 As String,
                 ByVal Item_Common6 As String, ByVal Item_Common7 As String, ByVal Item_Common8 As String, ByVal Item_Common9 As String, ByVal Item_Common10 As String,
                 ByVal SORT_ITEM_COMMON1 As String, ByVal SORT_ITEM_COMMON2 As String, ByVal SORT_ITEM_COMMON3 As String, ByVal SORT_ITEM_COMMON4 As String, ByVal SORT_ITEM_COMMON5 As String,
                 ByVal Comments As String, ByVal EXPIRED_DATE As String, ByVal SL_NO As String, ByVal Storage_Type As enuStorageType, ByVal BND As Boolean, ByVal QC_Status As String,
                 ByVal FROM_OWNER_No As String, ByVal FROM_SUB_OWNER_No As String, ByVal TO_OWNER_No As String, ByVal TO_SUB_OWNER_No As String,
                 ByVal FACTORY_No As String, ByVal DEST_AREA_No As String, ByVal DEST_LOCATION_No As String, ByVal SOURCE_AREA_NO As String, ByVal SOURCE_LOCATION_NO As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(WO_ID, WO_Serial_NO)
      _gid = key
      _gid = key
      _WO_ID = WO_ID
      _WO_Serial_NO = WO_Serial_NO
      _WORKING_TYPE = enuWorkingType.Null
      _WORKING_SERIAL_NO = ""
      _WORKING_SERIAL_SEQ = ""
      _QC_Method = enuQCMethod.Null
      _SKU_No = SKU_NO
      _SKU_Catalog = enuSKU_CATALOG.NULL
      _Lot_NO = LOT_NO
      _QTY = Qty
      _QTY_Transferred = 0
      _QTY_Process = 0
      _QTY_Replenishment = 0
      _QTY_Abort = 0
      _Carrier_ID = Carrier_ID
      _Package_ID = Package_ID
      _Item_Common1 = Item_Common1
      _Item_Common2 = Item_Common2
      _Item_Common3 = Item_Common3
      _Item_Common4 = Item_Common4
      _Item_Common5 = Item_Common5
      _Item_Common6 = Item_Common6
      _Item_Common7 = Item_Common7
      _Item_Common8 = Item_Common8
      _Item_Common9 = Item_Common9
      _Item_Common10 = Item_Common10
      _SORT_ITEM_COMMON1 = SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = SORT_ITEM_COMMON5
      _Comments = Comments
      _EXPIRED_DATE = EXPIRED_DATE
      _SL_NO = SL_NO
      _Storage_Type = Storage_Type
      _BND = BND
      _QC_Status = QC_Status
      _FROM_OWNER_No = FROM_OWNER_No
      _FROM_SUB_OWNER_No = FROM_SUB_OWNER_No
      _TO_OWNER_No = TO_OWNER_No
      _TO_SUB_OWNER_No = TO_SUB_OWNER_No
      _FACTORY_No = FACTORY_No
      _DEST_AREA_No = DEST_AREA_No
      _DEST_LOCATION_No = DEST_LOCATION_No
      _SOURCE_AREA_No = SOURCE_AREA_NO
      _SOURCE_LOCATION_NO = SOURCE_LOCATION_NO
      _Start_Time = ""
      _Start_Transfer_Time = ""
      _Finish_Transfer_Time = ""
      _Finish_Time = ""
      _WO_DTL_DC_Status = enuWO_DTL_DC_Status.Queued
      _DTL_Common1 = ""
      _DTL_Common2 = ""
      _DTL_Common3 = ""
      _DTL_Common4 = ""
      _DTL_Common5 = ""
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal WO_ID As String, ByVal WO_Serial_No As String) As String
    Try
      Dim key As String = WO_ID & LinkKey & WO_Serial_No
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsWO_DTL
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
      Dim strSQL As String = WMS_T_WO_DTLManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL(改版寫法)
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_WO_DTLManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_WO_DTLManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '=================Public Function======================
  Public Function Update_To_Memory(ByRef obj As clsWO_DTL) As Boolean
    Try
      Dim key As String = obj._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      'WO_ID = obj.WO_ID
      'WO_Serial_No = obj.WO_Serial_No
      WORKING_TYPE = obj.WORKING_TYPE
      WORKING_SERIAL_NO = obj.WORKING_SERIAL_NO
      WORKING_SERIAL_SEQ = obj.WORKING_SERIAL_SEQ
      QC_Method = obj.QC_Method
      SKU_No = obj.SKU_No
      SKU_Catalog = obj.SKU_Catalog
      Lot_No = obj.Lot_No
      QTY = obj.QTY
      QTY_Transferred = obj.QTY_Transferred
      QTY_Process = obj.QTY_Process
      QTY_Replenishment = obj.QTY_Replenishment
      QTY_Abort = obj.QTY_Abort
      Comments = obj.Comments
      Carrier_ID = obj.Carrier_ID
      Package_ID = obj.Package_ID
      Item_Common1 = obj.Item_Common1
      Item_Common2 = obj.Item_Common2
      Item_Common3 = obj.Item_Common3
      Item_Common4 = obj.Item_Common4
      Item_Common5 = obj.Item_Common5
      Item_Common6 = obj.Item_Common6
      Item_Common7 = obj.Item_Common7
      Item_Common8 = obj.Item_Common8
      Item_Common9 = obj.Item_Common9
      Item_Common10 = obj.Item_Common10
      SL_NO = obj.SL_NO
      Storage_Type = obj.Storage_Type
      BND = obj.BND
      QC_Status = obj.QC_Status
      FROM_OWNER_No = obj.FROM_OWNER_No
      FROM_SUB_OWNER_No = obj.FROM_SUB_OWNER_No
      TO_OWNER_No = obj.TO_OWNER_No
      TO_SUB_OWNER_No = obj.TO_SUB_OWNER_No
      FACTORY_No = obj.FACTORY_No
      DEST_AREA_No = obj.DEST_AREA_No
      DEST_LOCATION_No = obj.DEST_LOCATION_No
      SOURCE_AREA_No = obj.SOURCE_AREA_No
      SOURCE_LOCATION_NO = obj.SOURCE_LOCATION_NO
      SORT_ITEM_COMMON1 = obj.SORT_ITEM_COMMON1
      SORT_ITEM_COMMON2 = obj.SORT_ITEM_COMMON2
      SORT_ITEM_COMMON3 = obj.SORT_ITEM_COMMON3
      SORT_ITEM_COMMON4 = obj.SORT_ITEM_COMMON4
      SORT_ITEM_COMMON5 = obj.SORT_ITEM_COMMON5
      EXPIRED_DATE = obj.EXPIRED_DATE
      Start_Time = obj.Start_Time
      Start_Transfer_Time = obj.Start_Transfer_Time
      Finish_Transfer_Time = obj.Finish_Transfer_Time
      Finish_Time = obj.Finish_Time
      WO_DTL_DC_Status = obj.WO_DTL_DC_Status
      DTL_Common1 = obj.DTL_Common1
      DTL_Common2 = obj.DTL_Common2
      DTL_Common3 = obj.DTL_Common3
      DTL_Common4 = obj.DTL_Common4
      DTL_Common5 = obj.DTL_Common5
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Class
