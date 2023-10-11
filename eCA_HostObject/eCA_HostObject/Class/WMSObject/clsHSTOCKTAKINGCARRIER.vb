Public Class clsHSTOCKTAKINGCARRIER
Private ShareName As String = "HSTOCKTAKINGCARRIER"
Private ShareKey As String = ""
Private _gid As String 
Private _KEY_NO As String'流水號
 
Private _STOCKTAKING_ID As String'盤點單號
 
Private _STOCKTAKING_SERIAL_NO As String'盤點明細編號(流水號)
 
Private _CARRIER_ID As String'棧板編號
 
Private _QTY As Double'貨品數量
 
Private _PACKAGE_ID As String'箱ID/包裝ID
 
Private _SKU_NO As String'貨品ID
 
Private _LOT_NO As String'批號
 
Private _ITEM_COMMON1 As String'條件1
 
Private _ITEM_COMMON2 As String'條件2
 
Private _ITEM_COMMON3 As String'條件3
 
Private _ITEM_COMMON4 As String'條件4
 
Private _ITEM_COMMON5 As String'條件5
 
Private _ITEM_COMMON6 As String'條件6
 
Private _ITEM_COMMON7 As String'條件7
 
Private _ITEM_COMMON8 As String'條件8
 
Private _ITEM_COMMON9 As String'條件9
 
Private _ITEM_COMMON10 As String'條件10
 
Private _SORT_ITEM_COMMON1 As String'優先選擇條件1
 
Private _SORT_ITEM_COMMON2 As String'優先選擇條件2
 
Private _SORT_ITEM_COMMON3 As String'優先選擇條件3
 
Private _SORT_ITEM_COMMON4 As String'優先選擇條件4
 
Private _SORT_ITEM_COMMON5 As String'優先選擇條件5
 
Private _OWNER_NO As String'貨主
 
Private _SUB_OWNER_NO As String'子貨主
 
Private _SL_NO As String'ERP的儲存地點
 
Private _STORAGE_TYPE As Long'是否為暫存品Store 一般品=1Temporary 暫存品=2
 
Private _BND As Boolean'保稅0:不保稅1:保稅
 
Private _QC_STATUS As Long'QC判定狀態NA=0OK=1NG=2LOCK=3
 
Private _MANUFACETURE_DATE As String'製造日
 
Private _EXPIRED_DATE As String'到期日
 
Private _REPORT_QTY As Double'實際盤點數量
 
Private _REPORT_PACKAGE_ID As String'箱ID/包裝ID
 
Private _REPORT_SKU_NO As String'貨品ID
 
Private _REPORT_LOT_NO As String'批號
 
Private _REPORT_ITEM_COMMON1 As String'條件1
 
Private _REPORT_ITEM_COMMON2 As String'條件2
 
Private _REPORT_ITEM_COMMON3 As String'條件3
 
Private _REPORT_ITEM_COMMON4 As String'條件4
 
Private _REPORT_ITEM_COMMON5 As String'條件5
 
Private _REPORT_ITEM_COMMON6 As String'條件6
 
Private _REPORT_ITEM_COMMON7 As String'條件7
 
Private _REPORT_ITEM_COMMON8 As String'條件8
 
Private _REPORT_ITEM_COMMON9 As String'條件9
 
Private _REPORT_ITEM_COMMON10 As String'條件10
 
Private _REPORT_SORT_ITEM_COMMON1 As String'優先選擇條件1
 
Private _REPORT_SORT_ITEM_COMMON2 As String'優先選擇條件2
 
Private _REPORT_SORT_ITEM_COMMON3 As String'優先選擇條件3
 
Private _REPORT_SORT_ITEM_COMMON4 As String'優先選擇條件4
 
Private _REPORT_SORT_ITEM_COMMON5 As String'優先選擇條件5
 
Private _REPORT_OWNER_NO As String'貨主
 
Private _REPORT_SUB_OWNER_NO As String'子貨主
 
Private _REPORT_SL_NO As String'ERP的儲存地點
 
Private _REPORT_STORAGE_TYPE As Long'是否為暫存品Store 一般品=1Temporary 暫存品=2
 
Private _REPORT_BND As Boolean'保稅0:不保稅1:保稅
 
Private _REPORT_QC_STATUS As Long'QC判定狀態NA=0OK=1NG=2LOCK=3
 
Private _REPORT_MANUFACETURE_DATE As String'製造日
 
Private _REPORT_EXPIRED_DATE As String'到期日
 
Private _STOCKTAKING_COUNT As Long'已盤點次數
 
Private _STOCKTAKING_STATUS As Long'棧板盤點狀態Queued=0(未盤)Checked=1(已盤)
 
Private _PROFITP_TYPE As Long'盤盈虧類型Unknown = 0     '未知Normal = 1      '正常Surplus = 2     '盤盈Deficit = 3     '盤虧DataError = 4 	'資料異常
 
Private _REPORT_USER As String'盤點人員
 
Private _DEST_LOCATION_NO As String'位置
 
Private _ACTUAL_AREA_NO As String'進行出庫的位置(棧板原位置)
 
Private _ACTUAL_LOCATION_NO As String'進行出庫的位置(棧板原位置)
 
Private _ACTUAL_SUBLOCATION_X As String'進行出庫的位置(棧板原位置)
 
Private _ACTUAL_SUBLOCATION_Y As String'進行出庫的位置(棧板原位置)
 
Private _ACTUAL_SUBLOCATION_Z As String'進行出庫的位置(棧板原位置)
 
Private _MANUAL_KEY As String'手動建立的Key_No，只有手動建立的會填，不是手動建立的填空的
 
Private _HIST_TIME As String '紀錄時間

  Private _objWMS As clsHandlingObject
  Public Property gid() As String
 Get
 Return _gid
 End Get
 Set(ByVal value As String)
 _gid = value
 End Set
 End Property
 Public Property KEY_NO() As  String
 Get
 Return _KEY_NO
 End Get
 Set(ByVal value As  String)
 _KEY_NO = value
 End Set
 End Property
 Public Property STOCKTAKING_ID() As  String
 Get
 Return _STOCKTAKING_ID
 End Get
 Set(ByVal value As  String)
 _STOCKTAKING_ID = value
 End Set
 End Property
 Public Property STOCKTAKING_SERIAL_NO() As  String
 Get
 Return _STOCKTAKING_SERIAL_NO
 End Get
 Set(ByVal value As  String)
 _STOCKTAKING_SERIAL_NO = value
 End Set
 End Property
 Public Property CARRIER_ID() As  String
 Get
 Return _CARRIER_ID
 End Get
 Set(ByVal value As  String)
 _CARRIER_ID = value
 End Set
 End Property
 Public Property QTY() As  Double
 Get
 Return _QTY
 End Get
 Set(ByVal value As  Double)
 _QTY = value
 End Set
 End Property
 Public Property PACKAGE_ID() As  String
 Get
 Return _PACKAGE_ID
 End Get
 Set(ByVal value As  String)
 _PACKAGE_ID = value
 End Set
 End Property
 Public Property SKU_NO() As  String
 Get
 Return _SKU_NO
 End Get
 Set(ByVal value As  String)
 _SKU_NO = value
 End Set
 End Property
 Public Property LOT_NO() As  String
 Get
 Return _LOT_NO
 End Get
 Set(ByVal value As  String)
 _LOT_NO = value
 End Set
 End Property
 Public Property ITEM_COMMON1() As  String
 Get
 Return _ITEM_COMMON1
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON1 = value
 End Set
 End Property
 Public Property ITEM_COMMON2() As  String
 Get
 Return _ITEM_COMMON2
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON2 = value
 End Set
 End Property
 Public Property ITEM_COMMON3() As  String
 Get
 Return _ITEM_COMMON3
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON3 = value
 End Set
 End Property
 Public Property ITEM_COMMON4() As  String
 Get
 Return _ITEM_COMMON4
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON4 = value
 End Set
 End Property
 Public Property ITEM_COMMON5() As  String
 Get
 Return _ITEM_COMMON5
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON5 = value
 End Set
 End Property
 Public Property ITEM_COMMON6() As  String
 Get
 Return _ITEM_COMMON6
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON6 = value
 End Set
 End Property
 Public Property ITEM_COMMON7() As  String
 Get
 Return _ITEM_COMMON7
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON7 = value
 End Set
 End Property
 Public Property ITEM_COMMON8() As  String
 Get
 Return _ITEM_COMMON8
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON8 = value
 End Set
 End Property
 Public Property ITEM_COMMON9() As  String
 Get
 Return _ITEM_COMMON9
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON9 = value
 End Set
 End Property
 Public Property ITEM_COMMON10() As  String
 Get
 Return _ITEM_COMMON10
 End Get
 Set(ByVal value As  String)
 _ITEM_COMMON10 = value
 End Set
 End Property
 Public Property SORT_ITEM_COMMON1() As  String
 Get
 Return _SORT_ITEM_COMMON1
 End Get
 Set(ByVal value As  String)
 _SORT_ITEM_COMMON1 = value
 End Set
 End Property
 Public Property SORT_ITEM_COMMON2() As  String
 Get
 Return _SORT_ITEM_COMMON2
 End Get
 Set(ByVal value As  String)
 _SORT_ITEM_COMMON2 = value
 End Set
 End Property
 Public Property SORT_ITEM_COMMON3() As  String
 Get
 Return _SORT_ITEM_COMMON3
 End Get
 Set(ByVal value As  String)
 _SORT_ITEM_COMMON3 = value
 End Set
 End Property
 Public Property SORT_ITEM_COMMON4() As  String
 Get
 Return _SORT_ITEM_COMMON4
 End Get
 Set(ByVal value As  String)
 _SORT_ITEM_COMMON4 = value
 End Set
 End Property
 Public Property SORT_ITEM_COMMON5() As  String
 Get
 Return _SORT_ITEM_COMMON5
 End Get
 Set(ByVal value As  String)
 _SORT_ITEM_COMMON5 = value
 End Set
 End Property
 Public Property OWNER_NO() As  String
 Get
 Return _OWNER_NO
 End Get
 Set(ByVal value As  String)
 _OWNER_NO = value
 End Set
 End Property
 Public Property SUB_OWNER_NO() As  String
 Get
 Return _SUB_OWNER_NO
 End Get
 Set(ByVal value As  String)
 _SUB_OWNER_NO = value
 End Set
 End Property
 Public Property SL_NO() As  String
 Get
 Return _SL_NO
 End Get
 Set(ByVal value As  String)
 _SL_NO = value
 End Set
 End Property
 Public Property STORAGE_TYPE() As  Long
 Get
 Return _STORAGE_TYPE
 End Get
 Set(ByVal value As  Long)
 _STORAGE_TYPE = value
 End Set
 End Property
 Public Property BND() As  Boolean
 Get
 Return _BND
 End Get
 Set(ByVal value As  Boolean)
 _BND = value
 End Set
 End Property
 Public Property QC_STATUS() As  Long
 Get
 Return _QC_STATUS
 End Get
 Set(ByVal value As  Long)
 _QC_STATUS = value
 End Set
 End Property
 Public Property MANUFACETURE_DATE() As  String
 Get
 Return _MANUFACETURE_DATE
 End Get
 Set(ByVal value As  String)
 _MANUFACETURE_DATE = value
 End Set
 End Property
 Public Property EXPIRED_DATE() As  String
 Get
 Return _EXPIRED_DATE
 End Get
 Set(ByVal value As  String)
 _EXPIRED_DATE = value
 End Set
 End Property
 Public Property REPORT_QTY() As  Double
 Get
 Return _REPORT_QTY
 End Get
 Set(ByVal value As  Double)
 _REPORT_QTY = value
 End Set
 End Property
 Public Property REPORT_PACKAGE_ID() As  String
 Get
 Return _REPORT_PACKAGE_ID
 End Get
 Set(ByVal value As  String)
 _REPORT_PACKAGE_ID = value
 End Set
 End Property
 Public Property REPORT_SKU_NO() As  String
 Get
 Return _REPORT_SKU_NO
 End Get
 Set(ByVal value As  String)
 _REPORT_SKU_NO = value
 End Set
 End Property
 Public Property REPORT_LOT_NO() As  String
 Get
 Return _REPORT_LOT_NO
 End Get
 Set(ByVal value As  String)
 _REPORT_LOT_NO = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON1() As  String
 Get
 Return _REPORT_ITEM_COMMON1
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON1 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON2() As  String
 Get
 Return _REPORT_ITEM_COMMON2
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON2 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON3() As  String
 Get
 Return _REPORT_ITEM_COMMON3
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON3 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON4() As  String
 Get
 Return _REPORT_ITEM_COMMON4
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON4 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON5() As  String
 Get
 Return _REPORT_ITEM_COMMON5
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON5 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON6() As  String
 Get
 Return _REPORT_ITEM_COMMON6
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON6 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON7() As  String
 Get
 Return _REPORT_ITEM_COMMON7
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON7 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON8() As  String
 Get
 Return _REPORT_ITEM_COMMON8
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON8 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON9() As  String
 Get
 Return _REPORT_ITEM_COMMON9
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON9 = value
 End Set
 End Property
 Public Property REPORT_ITEM_COMMON10() As  String
 Get
 Return _REPORT_ITEM_COMMON10
 End Get
 Set(ByVal value As  String)
 _REPORT_ITEM_COMMON10 = value
 End Set
 End Property
 Public Property REPORT_SORT_ITEM_COMMON1() As  String
 Get
 Return _REPORT_SORT_ITEM_COMMON1
 End Get
 Set(ByVal value As  String)
 _REPORT_SORT_ITEM_COMMON1 = value
 End Set
 End Property
 Public Property REPORT_SORT_ITEM_COMMON2() As  String
 Get
 Return _REPORT_SORT_ITEM_COMMON2
 End Get
 Set(ByVal value As  String)
 _REPORT_SORT_ITEM_COMMON2 = value
 End Set
 End Property
 Public Property REPORT_SORT_ITEM_COMMON3() As  String
 Get
 Return _REPORT_SORT_ITEM_COMMON3
 End Get
 Set(ByVal value As  String)
 _REPORT_SORT_ITEM_COMMON3 = value
 End Set
 End Property
 Public Property REPORT_SORT_ITEM_COMMON4() As  String
 Get
 Return _REPORT_SORT_ITEM_COMMON4
 End Get
 Set(ByVal value As  String)
 _REPORT_SORT_ITEM_COMMON4 = value
 End Set
 End Property
 Public Property REPORT_SORT_ITEM_COMMON5() As  String
 Get
 Return _REPORT_SORT_ITEM_COMMON5
 End Get
 Set(ByVal value As  String)
 _REPORT_SORT_ITEM_COMMON5 = value
 End Set
 End Property
 Public Property REPORT_OWNER_NO() As  String
 Get
 Return _REPORT_OWNER_NO
 End Get
 Set(ByVal value As  String)
 _REPORT_OWNER_NO = value
 End Set
 End Property
 Public Property REPORT_SUB_OWNER_NO() As  String
 Get
 Return _REPORT_SUB_OWNER_NO
 End Get
 Set(ByVal value As  String)
 _REPORT_SUB_OWNER_NO = value
 End Set
 End Property
 Public Property REPORT_SL_NO() As  String
 Get
 Return _REPORT_SL_NO
 End Get
 Set(ByVal value As  String)
 _REPORT_SL_NO = value
 End Set
 End Property
 Public Property REPORT_STORAGE_TYPE() As  Long
 Get
 Return _REPORT_STORAGE_TYPE
 End Get
 Set(ByVal value As  Long)
 _REPORT_STORAGE_TYPE = value
 End Set
 End Property
 Public Property REPORT_BND() As  Boolean
 Get
 Return _REPORT_BND
 End Get
 Set(ByVal value As  Boolean)
 _REPORT_BND = value
 End Set
 End Property
 Public Property REPORT_QC_STATUS() As  Long
 Get
 Return _REPORT_QC_STATUS
 End Get
 Set(ByVal value As  Long)
 _REPORT_QC_STATUS = value
 End Set
 End Property
 Public Property REPORT_MANUFACETURE_DATE() As  String
 Get
 Return _REPORT_MANUFACETURE_DATE
 End Get
 Set(ByVal value As  String)
 _REPORT_MANUFACETURE_DATE = value
 End Set
 End Property
 Public Property REPORT_EXPIRED_DATE() As  String
 Get
 Return _REPORT_EXPIRED_DATE
 End Get
 Set(ByVal value As  String)
 _REPORT_EXPIRED_DATE = value
 End Set
 End Property
 Public Property STOCKTAKING_COUNT() As  Long
 Get
 Return _STOCKTAKING_COUNT
 End Get
 Set(ByVal value As  Long)
 _STOCKTAKING_COUNT = value
 End Set
 End Property
 Public Property STOCKTAKING_STATUS() As  Long
 Get
 Return _STOCKTAKING_STATUS
 End Get
 Set(ByVal value As  Long)
 _STOCKTAKING_STATUS = value
 End Set
 End Property
 Public Property PROFITP_TYPE() As  Long
 Get
 Return _PROFITP_TYPE
 End Get
 Set(ByVal value As  Long)
 _PROFITP_TYPE = value
 End Set
 End Property
 Public Property REPORT_USER() As  String
 Get
 Return _REPORT_USER
 End Get
 Set(ByVal value As  String)
 _REPORT_USER = value
 End Set
 End Property
 Public Property DEST_LOCATION_NO() As  String
 Get
 Return _DEST_LOCATION_NO
 End Get
 Set(ByVal value As  String)
 _DEST_LOCATION_NO = value
 End Set
 End Property
 Public Property ACTUAL_AREA_NO() As  String
 Get
 Return _ACTUAL_AREA_NO
 End Get
 Set(ByVal value As  String)
 _ACTUAL_AREA_NO = value
 End Set
 End Property
 Public Property ACTUAL_LOCATION_NO() As  String
 Get
 Return _ACTUAL_LOCATION_NO
 End Get
 Set(ByVal value As  String)
 _ACTUAL_LOCATION_NO = value
 End Set
 End Property
 Public Property ACTUAL_SUBLOCATION_X() As  String
 Get
 Return _ACTUAL_SUBLOCATION_X
 End Get
 Set(ByVal value As  String)
 _ACTUAL_SUBLOCATION_X = value
 End Set
 End Property
 Public Property ACTUAL_SUBLOCATION_Y() As  String
 Get
 Return _ACTUAL_SUBLOCATION_Y
 End Get
 Set(ByVal value As  String)
 _ACTUAL_SUBLOCATION_Y = value
 End Set
 End Property
 Public Property ACTUAL_SUBLOCATION_Z() As  String
 Get
 Return _ACTUAL_SUBLOCATION_Z
 End Get
 Set(ByVal value As  String)
 _ACTUAL_SUBLOCATION_Z = value
 End Set
 End Property
 Public Property MANUAL_KEY() As  String
 Get
 Return _MANUAL_KEY
 End Get
 Set(ByVal value As  String)
 _MANUAL_KEY = value
 End Set
 End Property
 Public Property HIST_TIME() As  String
 Get
 Return _HIST_TIME
 End Get
 Set(ByVal value As  String)
 _HIST_TIME = value
 End Set
 End Property
  Public Property objWMS() As clsHandlingObject
    Get
      Return _objWMS
    End Get
    Set(ByVal value As clsHandlingObject)
      _objWMS = value
    End Set
  End Property

  Public Sub New(ByVal KEY_NO As  String,ByVal STOCKTAKING_ID As  String,ByVal STOCKTAKING_SERIAL_NO As  String,ByVal CARRIER_ID As  String,ByVal QTY As  Double,ByVal PACKAGE_ID As  String,ByVal SKU_NO As  String,ByVal LOT_NO As  String,ByVal ITEM_COMMON1 As  String,ByVal ITEM_COMMON2 As  String,ByVal ITEM_COMMON3 As  String,ByVal ITEM_COMMON4 As  String,ByVal ITEM_COMMON5 As  String,ByVal ITEM_COMMON6 As  String,ByVal ITEM_COMMON7 As  String,ByVal ITEM_COMMON8 As  String,ByVal ITEM_COMMON9 As  String,ByVal ITEM_COMMON10 As  String,ByVal SORT_ITEM_COMMON1 As  String,ByVal SORT_ITEM_COMMON2 As  String,ByVal SORT_ITEM_COMMON3 As  String,ByVal SORT_ITEM_COMMON4 As  String,ByVal SORT_ITEM_COMMON5 As  String,ByVal OWNER_NO As  String,ByVal SUB_OWNER_NO As  String,ByVal SL_NO As  String,ByVal STORAGE_TYPE As  Long,ByVal BND As  Boolean,ByVal QC_STATUS As  Long,ByVal MANUFACETURE_DATE As  String,ByVal EXPIRED_DATE As  String,ByVal REPORT_QTY As  Double,ByVal REPORT_PACKAGE_ID As  String,ByVal REPORT_SKU_NO As  String,ByVal REPORT_LOT_NO As  String,ByVal REPORT_ITEM_COMMON1 As  String,ByVal REPORT_ITEM_COMMON2 As  String,ByVal REPORT_ITEM_COMMON3 As  String,ByVal REPORT_ITEM_COMMON4 As  String,ByVal REPORT_ITEM_COMMON5 As  String,ByVal REPORT_ITEM_COMMON6 As  String,ByVal REPORT_ITEM_COMMON7 As  String,ByVal REPORT_ITEM_COMMON8 As  String,ByVal REPORT_ITEM_COMMON9 As  String,ByVal REPORT_ITEM_COMMON10 As  String,ByVal REPORT_SORT_ITEM_COMMON1 As  String,ByVal REPORT_SORT_ITEM_COMMON2 As  String,ByVal REPORT_SORT_ITEM_COMMON3 As  String,ByVal REPORT_SORT_ITEM_COMMON4 As  String,ByVal REPORT_SORT_ITEM_COMMON5 As  String,ByVal REPORT_OWNER_NO As  String,ByVal REPORT_SUB_OWNER_NO As  String,ByVal REPORT_SL_NO As  String,ByVal REPORT_STORAGE_TYPE As  Long,ByVal REPORT_BND As  Boolean,ByVal REPORT_QC_STATUS As  Long,ByVal REPORT_MANUFACETURE_DATE As  String,ByVal REPORT_EXPIRED_DATE As  String,ByVal STOCKTAKING_COUNT As  Long,ByVal STOCKTAKING_STATUS As  Long,ByVal PROFITP_TYPE As  Long,ByVal REPORT_USER As  String,ByVal DEST_LOCATION_NO As  String,ByVal ACTUAL_AREA_NO As  String,ByVal ACTUAL_LOCATION_NO As  String,ByVal ACTUAL_SUBLOCATION_X As  String,ByVal ACTUAL_SUBLOCATION_Y As  String,ByVal ACTUAL_SUBLOCATION_Z As  String,ByVal MANUAL_KEY As  String,ByVal HIST_TIME As  String)
 MyBase.New()
 Try
 Dim key As String = Get_Combination_Key(KEY_NO)
 _gid = key
_KEY_NO = KEY_NO
_STOCKTAKING_ID = STOCKTAKING_ID
_STOCKTAKING_SERIAL_NO = STOCKTAKING_SERIAL_NO
_CARRIER_ID = CARRIER_ID
_QTY = QTY
_PACKAGE_ID = PACKAGE_ID
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
_SL_NO = SL_NO
_STORAGE_TYPE = STORAGE_TYPE
_BND = BND
_QC_STATUS = QC_STATUS
_MANUFACETURE_DATE = MANUFACETURE_DATE
_EXPIRED_DATE = EXPIRED_DATE
_REPORT_QTY = REPORT_QTY
_REPORT_PACKAGE_ID = REPORT_PACKAGE_ID
_REPORT_SKU_NO = REPORT_SKU_NO
_REPORT_LOT_NO = REPORT_LOT_NO
_REPORT_ITEM_COMMON1 = REPORT_ITEM_COMMON1
_REPORT_ITEM_COMMON2 = REPORT_ITEM_COMMON2
_REPORT_ITEM_COMMON3 = REPORT_ITEM_COMMON3
_REPORT_ITEM_COMMON4 = REPORT_ITEM_COMMON4
_REPORT_ITEM_COMMON5 = REPORT_ITEM_COMMON5
_REPORT_ITEM_COMMON6 = REPORT_ITEM_COMMON6
_REPORT_ITEM_COMMON7 = REPORT_ITEM_COMMON7
_REPORT_ITEM_COMMON8 = REPORT_ITEM_COMMON8
_REPORT_ITEM_COMMON9 = REPORT_ITEM_COMMON9
_REPORT_ITEM_COMMON10 = REPORT_ITEM_COMMON10
_REPORT_SORT_ITEM_COMMON1 = REPORT_SORT_ITEM_COMMON1
_REPORT_SORT_ITEM_COMMON2 = REPORT_SORT_ITEM_COMMON2
_REPORT_SORT_ITEM_COMMON3 = REPORT_SORT_ITEM_COMMON3
_REPORT_SORT_ITEM_COMMON4 = REPORT_SORT_ITEM_COMMON4
_REPORT_SORT_ITEM_COMMON5 = REPORT_SORT_ITEM_COMMON5
_REPORT_OWNER_NO = REPORT_OWNER_NO
_REPORT_SUB_OWNER_NO = REPORT_SUB_OWNER_NO
_REPORT_SL_NO = REPORT_SL_NO
_REPORT_STORAGE_TYPE = REPORT_STORAGE_TYPE
_REPORT_BND = REPORT_BND
_REPORT_QC_STATUS = REPORT_QC_STATUS
_REPORT_MANUFACETURE_DATE = REPORT_MANUFACETURE_DATE
_REPORT_EXPIRED_DATE = REPORT_EXPIRED_DATE
_STOCKTAKING_COUNT = STOCKTAKING_COUNT
_STOCKTAKING_STATUS = STOCKTAKING_STATUS
_PROFITP_TYPE = PROFITP_TYPE
_REPORT_USER = REPORT_USER
_DEST_LOCATION_NO = DEST_LOCATION_NO
_ACTUAL_AREA_NO = ACTUAL_AREA_NO
_ACTUAL_LOCATION_NO = ACTUAL_LOCATION_NO
_ACTUAL_SUBLOCATION_X = ACTUAL_SUBLOCATION_X
_ACTUAL_SUBLOCATION_Y = ACTUAL_SUBLOCATION_Y
_ACTUAL_SUBLOCATION_Z = ACTUAL_SUBLOCATION_Z
_MANUAL_KEY = MANUAL_KEY
_HIST_TIME = HIST_TIME
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
Public Shared Function Get_Combination_Key(ByVal KEY_NO As  String) As String
 Try
 Dim key As String = KEY_NO
 Return key
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return ""
 End Try
 End Function
Public Function Clone() As clsHSTOCKTAKINGCARRIER
 Try
 Return Me.MemberwiseClone()
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return Nothing
 End Try
 End Function
  ' Public Sub Add_Relationship(ByRef objWMS As clsWMSObject)                                 
  '   Try                                                                                     
  '     '挷定WMS的關係                                                                        
  '     If objWMS IsNot Nothing Then                                                          
  '       _objWMS = objWMS                                                                    
  '       objWMS.O_Add_!!!!!這邊就是你要改的東西啦(Me)
  '     End If                                                                                
  '   Catch ex As Exception                                                                   
  '     SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)                
  '   End Try                                                                                 
  ' End Sub                                                                                   
  'Public Sub Remove_Relationship()                                          
  '  Try                                                                                    
  '    If _objWMS IsNot Nothing Then                                                                
  '      _objWMS.O_Remove_ !!!!!這也是你要改的東西        
  '    End If                                                                                       
  '  Catch ex As Exception                                                                          
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)                       
  '  End Try                                                                                        
  'End Sub                                                                                          
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
 Try
Dim strSQL As String = WMS_H_STOCKTAKING_CARRIERManagement.GetInsertSQL(Me)
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
 Dim strSQL As String = WMS_H_STOCKTAKING_CARRIERManagement.GetUpdateSQL(Me)
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
 Dim strSQL As String = WMS_H_STOCKTAKING_CARRIERManagement.GetDeleteSQL(Me)
 lstSQL.Add(strSQL)
 Return True
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return False
 End Try
 End Function
Public Function Update_To_Memory(ByRef objWMS_H_STOCKTAKING_CARRIER As  clsHSTOCKTAKINGCARRIER) As Boolean
Try
Dim key As String = objWMS_H_STOCKTAKING_CARRIER._gid
 If key <> _gid Then
 SendMessageToLog("Key can not Update, old_Key="& _gid & ",new_key=" &  key,eCALogTool.ILogTool.enuTrcLevel.lvWARN)
 Return False
 End If
_KEY_NO = objWMS_H_STOCKTAKING_CARRIER.KEY_NO
_STOCKTAKING_ID = objWMS_H_STOCKTAKING_CARRIER.STOCKTAKING_ID
_STOCKTAKING_SERIAL_NO = objWMS_H_STOCKTAKING_CARRIER.STOCKTAKING_SERIAL_NO
_CARRIER_ID = objWMS_H_STOCKTAKING_CARRIER.CARRIER_ID
_QTY = objWMS_H_STOCKTAKING_CARRIER.QTY
_PACKAGE_ID = objWMS_H_STOCKTAKING_CARRIER.PACKAGE_ID
_SKU_NO = objWMS_H_STOCKTAKING_CARRIER.SKU_NO
_LOT_NO = objWMS_H_STOCKTAKING_CARRIER.LOT_NO
_ITEM_COMMON1 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON1
_ITEM_COMMON2 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON2
_ITEM_COMMON3 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON3
_ITEM_COMMON4 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON4
_ITEM_COMMON5 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON5
_ITEM_COMMON6 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON6
_ITEM_COMMON7 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON7
_ITEM_COMMON8 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON8
_ITEM_COMMON9 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON9
_ITEM_COMMON10 = objWMS_H_STOCKTAKING_CARRIER.ITEM_COMMON10
_SORT_ITEM_COMMON1 = objWMS_H_STOCKTAKING_CARRIER.SORT_ITEM_COMMON1
_SORT_ITEM_COMMON2 = objWMS_H_STOCKTAKING_CARRIER.SORT_ITEM_COMMON2
_SORT_ITEM_COMMON3 = objWMS_H_STOCKTAKING_CARRIER.SORT_ITEM_COMMON3
_SORT_ITEM_COMMON4 = objWMS_H_STOCKTAKING_CARRIER.SORT_ITEM_COMMON4
_SORT_ITEM_COMMON5 = objWMS_H_STOCKTAKING_CARRIER.SORT_ITEM_COMMON5
_OWNER_NO = objWMS_H_STOCKTAKING_CARRIER.OWNER_NO
_SUB_OWNER_NO = objWMS_H_STOCKTAKING_CARRIER.SUB_OWNER_NO
_SL_NO = objWMS_H_STOCKTAKING_CARRIER.SL_NO
_STORAGE_TYPE = objWMS_H_STOCKTAKING_CARRIER.STORAGE_TYPE
_BND = objWMS_H_STOCKTAKING_CARRIER.BND
_QC_STATUS = objWMS_H_STOCKTAKING_CARRIER.QC_STATUS
_MANUFACETURE_DATE = objWMS_H_STOCKTAKING_CARRIER.MANUFACETURE_DATE
_EXPIRED_DATE = objWMS_H_STOCKTAKING_CARRIER.EXPIRED_DATE
_REPORT_QTY = objWMS_H_STOCKTAKING_CARRIER.REPORT_QTY
_REPORT_PACKAGE_ID = objWMS_H_STOCKTAKING_CARRIER.REPORT_PACKAGE_ID
_REPORT_SKU_NO = objWMS_H_STOCKTAKING_CARRIER.REPORT_SKU_NO
_REPORT_LOT_NO = objWMS_H_STOCKTAKING_CARRIER.REPORT_LOT_NO
_REPORT_ITEM_COMMON1 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON1
_REPORT_ITEM_COMMON2 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON2
_REPORT_ITEM_COMMON3 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON3
_REPORT_ITEM_COMMON4 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON4
_REPORT_ITEM_COMMON5 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON5
_REPORT_ITEM_COMMON6 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON6
_REPORT_ITEM_COMMON7 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON7
_REPORT_ITEM_COMMON8 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON8
_REPORT_ITEM_COMMON9 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON9
_REPORT_ITEM_COMMON10 = objWMS_H_STOCKTAKING_CARRIER.REPORT_ITEM_COMMON10
_REPORT_SORT_ITEM_COMMON1 = objWMS_H_STOCKTAKING_CARRIER.REPORT_SORT_ITEM_COMMON1
_REPORT_SORT_ITEM_COMMON2 = objWMS_H_STOCKTAKING_CARRIER.REPORT_SORT_ITEM_COMMON2
_REPORT_SORT_ITEM_COMMON3 = objWMS_H_STOCKTAKING_CARRIER.REPORT_SORT_ITEM_COMMON3
_REPORT_SORT_ITEM_COMMON4 = objWMS_H_STOCKTAKING_CARRIER.REPORT_SORT_ITEM_COMMON4
_REPORT_SORT_ITEM_COMMON5 = objWMS_H_STOCKTAKING_CARRIER.REPORT_SORT_ITEM_COMMON5
_REPORT_OWNER_NO = objWMS_H_STOCKTAKING_CARRIER.REPORT_OWNER_NO
_REPORT_SUB_OWNER_NO = objWMS_H_STOCKTAKING_CARRIER.REPORT_SUB_OWNER_NO
_REPORT_SL_NO = objWMS_H_STOCKTAKING_CARRIER.REPORT_SL_NO
_REPORT_STORAGE_TYPE = objWMS_H_STOCKTAKING_CARRIER.REPORT_STORAGE_TYPE
_REPORT_BND = objWMS_H_STOCKTAKING_CARRIER.REPORT_BND
_REPORT_QC_STATUS = objWMS_H_STOCKTAKING_CARRIER.REPORT_QC_STATUS
_REPORT_MANUFACETURE_DATE = objWMS_H_STOCKTAKING_CARRIER.REPORT_MANUFACETURE_DATE
_REPORT_EXPIRED_DATE = objWMS_H_STOCKTAKING_CARRIER.REPORT_EXPIRED_DATE
_STOCKTAKING_COUNT = objWMS_H_STOCKTAKING_CARRIER.STOCKTAKING_COUNT
_STOCKTAKING_STATUS = objWMS_H_STOCKTAKING_CARRIER.STOCKTAKING_STATUS
_PROFITP_TYPE = objWMS_H_STOCKTAKING_CARRIER.PROFITP_TYPE
_REPORT_USER = objWMS_H_STOCKTAKING_CARRIER.REPORT_USER
_DEST_LOCATION_NO = objWMS_H_STOCKTAKING_CARRIER.DEST_LOCATION_NO
_ACTUAL_AREA_NO = objWMS_H_STOCKTAKING_CARRIER.ACTUAL_AREA_NO
_ACTUAL_LOCATION_NO = objWMS_H_STOCKTAKING_CARRIER.ACTUAL_LOCATION_NO
_ACTUAL_SUBLOCATION_X = objWMS_H_STOCKTAKING_CARRIER.ACTUAL_SUBLOCATION_X
_ACTUAL_SUBLOCATION_Y = objWMS_H_STOCKTAKING_CARRIER.ACTUAL_SUBLOCATION_Y
_ACTUAL_SUBLOCATION_Z = objWMS_H_STOCKTAKING_CARRIER.ACTUAL_SUBLOCATION_Z
_MANUAL_KEY = objWMS_H_STOCKTAKING_CARRIER.MANUAL_KEY
_HIST_TIME = objWMS_H_STOCKTAKING_CARRIER.HIST_TIME
 Return True
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return False
 End Try
 End Function
End Class
