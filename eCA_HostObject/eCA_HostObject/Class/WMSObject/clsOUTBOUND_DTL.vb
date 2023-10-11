Imports System.Collections.Concurrent
Public Class clsOUTBOUND_DTL
  Private ShareName As String = "OUTBOUND_DTL"
  Private ShareKey As String = ""
  Private _gid As String
  Private _KEY_NO As String '系統產生的Key值
  Private _WO_ID As String '工單編號
  Private _WO_SERIAL_NO As String '工單明細編號
  Private _CARRIER_ID As String '棧板編號
  Private _SKU_NO As String '貨品編號
  Private _QTY_WO As Decimal '需求量
  Private _QTY_OUTBOUND As Decimal '出庫數量
  Private _ITEM_KEY_NO As String '貨品的流水號
  Private _COMMENTS As String '備註
  Private _PACKAGE_ID As String '包裝編號
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
  Private _SUPPLIER_NO As String '供應商編號
  Private _CUSTOMER_NO As String '客戶編號
  Private _OWNER_NO As String '貨主編號
  Private _SUB_OWNER_NO As String '子貨主編號
  Private _STORAGE_TYPE As enuStorageType '儲放的類型Store 庫存品=1Temporary 暫存品=2
  Private _BND As Boolean '保稅0:不保稅1:保稅
  Private _QC_STATUS As enuQCStatus 'QC判定狀態OK=1NG=2NA=3
  Private _EXPIRED_DATE As String '到期日
  Private _RECEIPT_WO_ID As String '收料單號
  Private _RECEIPT_WO_SERIAL_NO As String '收料項次
  Private _PO_ID As String '訂單編號
  Private _PO_SERIAL_NO As String '訂單明細編號
  Private _OUTBOUND_STATUS As Long 'Queued = 0                  '未執行  WaitProcess = 1             '等待出庫  Prcoess = 2                 '出庫中  PrcoessFailed = 3           '出庫搬送失敗  Completed = 4               '完成  Picking = 5                 '等待分揀  BackQueued = 6              '分揀完成(未執行回庫)  WaitBack = 8                '等待回庫  Back = 9                    '回庫中  BackFailed = 10             '回庫失敗  BackCompleted = 11          '回庫完成  Kitting = 12                '等待二次分揀  FixedCarrierUnPick = 21  '不出庫的棧板等待分揀  CreatePickUp_DTL = 22  '產生了PickUP_DTL  WeightChecked = 98          '已進行秤重作業  WaitSimulationClose = 99    '假過帳，等待Host回覆
  Private _DEST_LOCATION_NO As String '位置
  Private _ACTUAL_AREA_NO As String '進行出庫的位置(棧板原位置)
  Private _ACTUAL_LOCATION_NO As String '進行出庫的位置(棧板原位置)
  Private _ACTUAL_SUBLOCATION_X As String '進行出庫的位置(棧板原位置)
  Private _ACTUAL_SUBLOCATION_Y As String '進行出庫的位置(棧板原位置)
  Private _ACTUAL_SUBLOCATION_Z As String '進行出庫的位置(棧板原位置)
  Private _USER_ID As String '建立人員
  Private _CLIENT_ID As String '出庫作業站
  Private _UNPACK_ABLE As Long '分揀設定不需要分揀=0正常分揀=1只揀少的=2
  Private _COMMAND_ID As String '命令編號
  Private _CREATE_TIME As String '建立時間
  Private _CREATE_CMD_TIME As String '產生命令時間
  Private _COMPLETED_TIME As String '完成時間
  Private _UNPACK_TIME As String '拆棧時間
  Private _UNPACK_CARRIER_ID As String '分揀前的棧板編號
  Private _CHECKED_TIME As String '出庫確認時間
  Private _CHECKED As Boolean '是否完成出庫確認0:未確認1:已確認
  Private _STEP_NO As Long '揀貨步驟號
  Public objWMS As clsHandlingObject
  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property KEY_NO() As String
    Get
      Return _KEY_NO
    End Get
    Set(ByVal value As String)
      _KEY_NO = value
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
  Public Property WO_SERIAL_NO() As String
    Get
      Return _WO_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _WO_SERIAL_NO = value
    End Set
  End Property
  Public Property CARRIER_ID() As String
    Get
      Return _CARRIER_ID
    End Get
    Set(ByVal value As String)
      _CARRIER_ID = value
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
  Public Property QTY_WO() As Decimal
    Get
      Return _QTY_WO
    End Get
    Set(ByVal value As Decimal)
      _QTY_WO = value
    End Set
  End Property
  Public Property QTY_OUTBOUND() As Decimal
    Get
      Return _QTY_OUTBOUND
    End Get
    Set(ByVal value As Decimal)
      _QTY_OUTBOUND = value
    End Set
  End Property
  Public Property ITEM_KEY_NO() As String
    Get
      Return _ITEM_KEY_NO
    End Get
    Set(ByVal value As String)
      _ITEM_KEY_NO = value
    End Set
  End Property
  Public Property COMMENTS() As String
    Get
      Return _COMMENTS
    End Get
    Set(ByVal value As String)
      _COMMENTS = value
    End Set
  End Property
  Public Property PACKAGE_ID() As String
    Get
      Return _PACKAGE_ID
    End Get
    Set(ByVal value As String)
      _PACKAGE_ID = value
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
  Public Property SUPPLIER_NO() As String
    Get
      Return _SUPPLIER_NO
    End Get
    Set(ByVal value As String)
      _SUPPLIER_NO = value
    End Set
  End Property
  Public Property CUSTOMER_NO() As String
    Get
      Return _CUSTOMER_NO
    End Get
    Set(ByVal value As String)
      _CUSTOMER_NO = value
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
  Public Property STORAGE_TYPE() As enuStorageType
    Get
      Return _STORAGE_TYPE
    End Get
    Set(ByVal value As enuStorageType)
      _STORAGE_TYPE = value
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
  Public Property QC_STATUS() As enuQCStatus
    Get
      Return _QC_STATUS
    End Get
    Set(ByVal value As enuQCStatus)
      _QC_STATUS = value
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
  Public Property RECEIPT_WO_ID() As String
    Get
      Return _RECEIPT_WO_ID
    End Get
    Set(ByVal value As String)
      _RECEIPT_WO_ID = value
    End Set
  End Property
  Public Property RECEIPT_WO_SERIAL_NO() As String
    Get
      Return _RECEIPT_WO_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _RECEIPT_WO_SERIAL_NO = value
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
  Public Property PO_SERIAL_NO() As String
    Get
      Return _PO_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _PO_SERIAL_NO = value
    End Set
  End Property
  Public Property OUTBOUND_STATUS() As Long
    Get
      Return _OUTBOUND_STATUS
    End Get
    Set(ByVal value As Long)
      _OUTBOUND_STATUS = value
    End Set
  End Property
  Public Property DEST_LOCATION_NO() As String
    Get
      Return _DEST_LOCATION_NO
    End Get
    Set(ByVal value As String)
      _DEST_LOCATION_NO = value
    End Set
  End Property
  Public Property ACTUAL_AREA_NO() As String
    Get
      Return _ACTUAL_AREA_NO
    End Get
    Set(ByVal value As String)
      _ACTUAL_AREA_NO = value
    End Set
  End Property
  Public Property ACTUAL_LOCATION_NO() As String
    Get
      Return _ACTUAL_LOCATION_NO
    End Get
    Set(ByVal value As String)
      _ACTUAL_LOCATION_NO = value
    End Set
  End Property
  Public Property ACTUAL_SUBLOCATION_X() As String
    Get
      Return _ACTUAL_SUBLOCATION_X
    End Get
    Set(ByVal value As String)
      _ACTUAL_SUBLOCATION_X = value
    End Set
  End Property
  Public Property ACTUAL_SUBLOCATION_Y() As String
    Get
      Return _ACTUAL_SUBLOCATION_Y
    End Get
    Set(ByVal value As String)
      _ACTUAL_SUBLOCATION_Y = value
    End Set
  End Property
  Public Property ACTUAL_SUBLOCATION_Z() As String
    Get
      Return _ACTUAL_SUBLOCATION_Z
    End Get
    Set(ByVal value As String)
      _ACTUAL_SUBLOCATION_Z = value
    End Set
  End Property
  Public Property USER_ID() As String
    Get
      Return _USER_ID
    End Get
    Set(ByVal value As String)
      _USER_ID = value
    End Set
  End Property
  Public Property CLIENT_ID() As String
    Get
      Return _CLIENT_ID
    End Get
    Set(ByVal value As String)
      _CLIENT_ID = value
    End Set
  End Property
  Public Property UNPACK_ABLE() As Long
    Get
      Return _UNPACK_ABLE
    End Get
    Set(ByVal value As Long)
      _UNPACK_ABLE = value
    End Set
  End Property
  Public Property COMMAND_ID() As String
    Get
      Return _COMMAND_ID
    End Get
    Set(ByVal value As String)
      _COMMAND_ID = value
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
  Public Property CREATE_CMD_TIME() As String
    Get
      Return _CREATE_CMD_TIME
    End Get
    Set(ByVal value As String)
      _CREATE_CMD_TIME = value
    End Set
  End Property
  Public Property COMPLETED_TIME() As String
    Get
      Return _COMPLETED_TIME
    End Get
    Set(ByVal value As String)
      _COMPLETED_TIME = value
    End Set
  End Property
  Public Property UNPACK_TIME() As String
    Get
      Return _UNPACK_TIME
    End Get
    Set(ByVal value As String)
      _UNPACK_TIME = value
    End Set
  End Property
  Public Property UNPACK_CARRIER_ID() As String
    Get
      Return _UNPACK_CARRIER_ID
    End Get
    Set(ByVal value As String)
      _UNPACK_CARRIER_ID = value
    End Set
  End Property
  Public Property CHECKED_TIME() As String
    Get
      Return _CHECKED_TIME
    End Get
    Set(ByVal value As String)
      _CHECKED_TIME = value
    End Set
  End Property
  Public Property CHECKED() As Boolean
    Get
      Return _CHECKED
    End Get
    Set(ByVal value As Boolean)
      _CHECKED = value
    End Set
  End Property
  Public Property STEP_NO() As Long
    Get
      Return _STEP_NO
    End Get
    Set(ByVal value As Long)
      _STEP_NO = value
    End Set
  End Property

  Public Sub New(ByVal KEY_NO As String, ByVal WO_ID As String, ByVal WO_SERIAL_NO As String, ByVal CARRIER_ID As String, ByVal SKU_NO As String, ByVal QTY_WO As Decimal, ByVal QTY_OUTBOUND As Decimal, ByVal ITEM_KEY_NO As String, ByVal COMMENTS As String, ByVal PACKAGE_ID As String, ByVal LOT_NO As String, ByVal ITEM_COMMON1 As String, ByVal ITEM_COMMON2 As String, ByVal ITEM_COMMON3 As String, ByVal ITEM_COMMON4 As String, ByVal ITEM_COMMON5 As String, ByVal ITEM_COMMON6 As String, ByVal ITEM_COMMON7 As String, ByVal ITEM_COMMON8 As String, ByVal ITEM_COMMON9 As String, ByVal ITEM_COMMON10 As String, ByVal SORT_ITEM_COMMON1 As String, ByVal SORT_ITEM_COMMON2 As String, ByVal SORT_ITEM_COMMON3 As String, ByVal SORT_ITEM_COMMON4 As String, ByVal SORT_ITEM_COMMON5 As String, ByVal SUPPLIER_NO As String, ByVal CUSTOMER_NO As String, ByVal OWNER_NO As String, ByVal SUB_OWNER_NO As String, ByVal STORAGE_TYPE As enuStorageType, ByVal BND As Boolean, ByVal QC_STATUS As enuQCStatus, ByVal EXPIRED_DATE As String, ByVal RECEIPT_WO_ID As String, ByVal RECEIPT_WO_SERIAL_NO As String, ByVal PO_ID As String, ByVal PO_SERIAL_NO As String, ByVal OUTBOUND_STATUS As Long, ByVal DEST_LOCATION_NO As String, ByVal ACTUAL_AREA_NO As String, ByVal ACTUAL_LOCATION_NO As String, ByVal ACTUAL_SUBLOCATION_X As String, ByVal ACTUAL_SUBLOCATION_Y As String, ByVal ACTUAL_SUBLOCATION_Z As String, ByVal USER_ID As String, ByVal CLIENT_ID As String, ByVal UNPACK_ABLE As Long, ByVal COMMAND_ID As String, ByVal CREATE_TIME As String, ByVal CREATE_CMD_TIME As String, ByVal COMPLETED_TIME As String, ByVal UNPACK_TIME As String, ByVal UNPACK_CARRIER_ID As String, ByVal CHECKED_TIME As String, ByVal CHECKED As Boolean, ByVal STEP_NO As Long)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(KEY_NO)
      _gid = key
      _KEY_NO = KEY_NO
      _WO_ID = WO_ID
      _WO_SERIAL_NO = WO_SERIAL_NO
      _CARRIER_ID = CARRIER_ID
      _SKU_NO = SKU_NO
      _QTY_WO = QTY_WO
      _QTY_OUTBOUND = QTY_OUTBOUND
      _ITEM_KEY_NO = ITEM_KEY_NO
      _COMMENTS = COMMENTS
      _PACKAGE_ID = PACKAGE_ID
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
      _SUPPLIER_NO = SUPPLIER_NO
      _CUSTOMER_NO = CUSTOMER_NO
      _OWNER_NO = OWNER_NO
      _SUB_OWNER_NO = SUB_OWNER_NO
      _STORAGE_TYPE = STORAGE_TYPE
      _BND = BND
      _QC_STATUS = QC_STATUS
      _EXPIRED_DATE = EXPIRED_DATE
      _RECEIPT_WO_ID = RECEIPT_WO_ID
      _RECEIPT_WO_SERIAL_NO = RECEIPT_WO_SERIAL_NO
      _PO_ID = PO_ID
      _PO_SERIAL_NO = PO_SERIAL_NO
      _OUTBOUND_STATUS = OUTBOUND_STATUS
      _DEST_LOCATION_NO = DEST_LOCATION_NO
      _ACTUAL_AREA_NO = ACTUAL_AREA_NO
      _ACTUAL_LOCATION_NO = ACTUAL_LOCATION_NO
      _ACTUAL_SUBLOCATION_X = ACTUAL_SUBLOCATION_X
      _ACTUAL_SUBLOCATION_Y = ACTUAL_SUBLOCATION_Y
      _ACTUAL_SUBLOCATION_Z = ACTUAL_SUBLOCATION_Z
      _USER_ID = USER_ID
      _CLIENT_ID = CLIENT_ID
      _UNPACK_ABLE = UNPACK_ABLE
      _COMMAND_ID = COMMAND_ID
      _CREATE_TIME = CREATE_TIME
      _CREATE_CMD_TIME = CREATE_CMD_TIME
      _COMPLETED_TIME = COMPLETED_TIME
      _UNPACK_TIME = UNPACK_TIME
      _UNPACK_CARRIER_ID = UNPACK_CARRIER_ID
      _CHECKED_TIME = CHECKED_TIME
      _CHECKED = CHECKED
      _STEP_NO = STEP_NO
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
  Public Shared Function Get_Combination_Key(ByVal KEY_NO As String) As String
    Try
      Dim key As String = KEY_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsOUTBOUND_DTL
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '新增記憶體內容
  'Public Sub Add_Relationship(ByRef objWMS As clsWMSObject)
  'Try
  ''綁定OUTBOUND_DTL和WMS的關係
  'If objWMS IsNot Nothing Then
  'Me.objWMS = objWMS
  ''此處如有更改，須自行修改
  'objWMS.O_Add_OUTBOUND_DTL(Me)
  'End If
  ' Catch ex As Exception
  ' SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  ' End Try
  'End Sub
  ''移除記憶體內容
  'Public Sub Remove_Relationship()
  'Try
  'If Me.objWMS IsNot Nothing Then
  'Me.objWMS.O_Remove_OUTBOUND_DTL(Me)
  'End If
  ' Catch ex As Exception
  ' SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  ' End Try
  'End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_OUTBOUND_DTLManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString1(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_OUTBOUND_DTLManagement.GetUpdateSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  'Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
  ' Try
  'Dim objOUTBOUND_DTL As clsOUTBOUND_DTL =Nothing
  'Dim dicChangeColumnValue As New Dictionary(Of String, String)
  'If O_Get_UpdateColumnValue(objOUTBOUND_DTL, Me,dicChangeColumnValue)=True Then
  'Dim strSQL As String = WMS_T_OUTBOUND_DTLManagement.GetUpdateSQLForChangeValue(Me,dicChangeColumnValue)
  'If strSQL <> "" Then
  'lstSQL.Add(strSQL)
  'End If
  'Else
  'SendMessageToLog("O_Get_UpdateColumnValue Faled", eCALogTool.ILogTool.enuTrcLevel.lvError)
  ''失敗先用原來的方式
  'Dim strSQL As String = WMS_T_OUTBOUND_DTLManagement.GetUpdateSQL(Me)
  'lstSQL.Add(strSQL)
  'End If
  'Return True
  ' Catch ex As Exception
  ' SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  ' Return False
  ' End Try
  ' End Function
  '取得要Delete的SQL
  Public Function O_Add_Delete_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_OUTBOUND_DTLManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_T_OUTBOUND_DTL As clsOUTBOUND_DTL) As Boolean
    Try
      Dim key As String = objWMS_T_OUTBOUND_DTL._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _KEY_NO = objWMS_T_OUTBOUND_DTL.KEY_NO
      _WO_ID = objWMS_T_OUTBOUND_DTL.WO_ID
      _WO_SERIAL_NO = objWMS_T_OUTBOUND_DTL.WO_SERIAL_NO
      _CARRIER_ID = objWMS_T_OUTBOUND_DTL.CARRIER_ID
      _SKU_NO = objWMS_T_OUTBOUND_DTL.SKU_NO
      _QTY_WO = objWMS_T_OUTBOUND_DTL.QTY_WO
      _QTY_OUTBOUND = objWMS_T_OUTBOUND_DTL.QTY_OUTBOUND
      _ITEM_KEY_NO = objWMS_T_OUTBOUND_DTL.ITEM_KEY_NO
      _COMMENTS = objWMS_T_OUTBOUND_DTL.COMMENTS
      _PACKAGE_ID = objWMS_T_OUTBOUND_DTL.PACKAGE_ID
      _LOT_NO = objWMS_T_OUTBOUND_DTL.LOT_NO
      _ITEM_COMMON1 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON1
      _ITEM_COMMON2 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON2
      _ITEM_COMMON3 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON3
      _ITEM_COMMON4 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON4
      _ITEM_COMMON5 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON5
      _ITEM_COMMON6 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON6
      _ITEM_COMMON7 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON7
      _ITEM_COMMON8 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON8
      _ITEM_COMMON9 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON9
      _ITEM_COMMON10 = objWMS_T_OUTBOUND_DTL.ITEM_COMMON10
      _SORT_ITEM_COMMON1 = objWMS_T_OUTBOUND_DTL.SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = objWMS_T_OUTBOUND_DTL.SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = objWMS_T_OUTBOUND_DTL.SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = objWMS_T_OUTBOUND_DTL.SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = objWMS_T_OUTBOUND_DTL.SORT_ITEM_COMMON5
      _SUPPLIER_NO = objWMS_T_OUTBOUND_DTL.SUPPLIER_NO
      _CUSTOMER_NO = objWMS_T_OUTBOUND_DTL.CUSTOMER_NO
      _OWNER_NO = objWMS_T_OUTBOUND_DTL.OWNER_NO
      _SUB_OWNER_NO = objWMS_T_OUTBOUND_DTL.SUB_OWNER_NO
      _STORAGE_TYPE = objWMS_T_OUTBOUND_DTL.STORAGE_TYPE
      _BND = objWMS_T_OUTBOUND_DTL.BND
      _QC_STATUS = objWMS_T_OUTBOUND_DTL.QC_STATUS
      _EXPIRED_DATE = objWMS_T_OUTBOUND_DTL.EXPIRED_DATE
      _RECEIPT_WO_ID = objWMS_T_OUTBOUND_DTL.RECEIPT_WO_ID
      _RECEIPT_WO_SERIAL_NO = objWMS_T_OUTBOUND_DTL.RECEIPT_WO_SERIAL_NO
      _PO_ID = objWMS_T_OUTBOUND_DTL.PO_ID
      _PO_SERIAL_NO = objWMS_T_OUTBOUND_DTL.PO_SERIAL_NO
      _OUTBOUND_STATUS = objWMS_T_OUTBOUND_DTL.OUTBOUND_STATUS
      _DEST_LOCATION_NO = objWMS_T_OUTBOUND_DTL.DEST_LOCATION_NO
      _ACTUAL_AREA_NO = objWMS_T_OUTBOUND_DTL.ACTUAL_AREA_NO
      _ACTUAL_LOCATION_NO = objWMS_T_OUTBOUND_DTL.ACTUAL_LOCATION_NO
      _ACTUAL_SUBLOCATION_X = objWMS_T_OUTBOUND_DTL.ACTUAL_SUBLOCATION_X
      _ACTUAL_SUBLOCATION_Y = objWMS_T_OUTBOUND_DTL.ACTUAL_SUBLOCATION_Y
      _ACTUAL_SUBLOCATION_Z = objWMS_T_OUTBOUND_DTL.ACTUAL_SUBLOCATION_Z
      _USER_ID = objWMS_T_OUTBOUND_DTL.USER_ID
      _CLIENT_ID = objWMS_T_OUTBOUND_DTL.CLIENT_ID
      _UNPACK_ABLE = objWMS_T_OUTBOUND_DTL.UNPACK_ABLE
      _COMMAND_ID = objWMS_T_OUTBOUND_DTL.COMMAND_ID
      _CREATE_TIME = objWMS_T_OUTBOUND_DTL.CREATE_TIME
      _CREATE_CMD_TIME = objWMS_T_OUTBOUND_DTL.CREATE_CMD_TIME
      _COMPLETED_TIME = objWMS_T_OUTBOUND_DTL.COMPLETED_TIME
      _UNPACK_TIME = objWMS_T_OUTBOUND_DTL.UNPACK_TIME
      _UNPACK_CARRIER_ID = objWMS_T_OUTBOUND_DTL.UNPACK_CARRIER_ID
      _CHECKED_TIME = objWMS_T_OUTBOUND_DTL.CHECKED_TIME
      _CHECKED = objWMS_T_OUTBOUND_DTL.CHECKED
      _STEP_NO = objWMS_T_OUTBOUND_DTL.STEP_NO
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
