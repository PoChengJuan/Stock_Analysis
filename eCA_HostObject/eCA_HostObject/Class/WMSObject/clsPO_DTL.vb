﻿Public Class clsPO_DTL
  Private ShareName As String = " & item.tableName & "
  Private ShareKey As String = " & "
  Private _gid As String
  Private _PO_ID As String '訂單編號

  Private _PO_LINE_NO As String '訂單明細編號(上傳時使用)

  Private _PO_SERIAL_NO As String '訂單明細編號(WMS使用)

  Private _WORKING_TYPE As String

  Private _WORKING_SERIAL_NO As String

  Private _WORKING_SERIAL_SEQ As String

  Private _SKU_NO As String '貨品編號

  Private _LOT_NO As String '批號

  Private _QTY As Double '需求量

  Private _QTY_PROCESS As Double '已併單數量

  Private _QTY_FINISH As Double '已完成數量

  Private _COMMENTS As String '備註

  Private _PACKAGE_ID As String '箱ID/包裝ID

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

  Private _STORAGE_TYPE As Double '是否為暫存品0:Store 一般品1:Temporary 暫存品

  Private _BND As Double '保稅0:不保稅1:保稅

  Private _QC_STATUS As Double 'QC判定狀態1:OK2:NG3:NA

  Private _FROM_OWNER_ID As String '原貨主編號

  Private _FROM_SUB_OWNER_ID As String '原子貨主編號

  Private _TO_OWNER_ID As String '新子貨主編號

  Private _TO_SUB_OWNER_ID As String '新貨主編號

  Private _FACTORY_ID As String '廠區(廠別)(預先指定出入庫區域)

  Private _DEST_AREA_ID As String '倉庫編號(預先指定出入庫區域)

  Private _DEST_LOCATION_ID As String '儲位編號(預先指定出入庫區域)

  Private _H_POD_STEP_NO As Double '

  Private _H_POD_MOVE_TYPE As String '

  Private _H_POD_FINISH_TIME As String '

  Private _H_POD_BILLING_DATE As String '

  Private _H_POD_CREATE_TIME As String '

  Private _H_POD1 As String '

  Private _H_POD2 As String '

  Private _H_POD3 As String '

  Private _H_POD4 As String '

  Private _H_POD5 As String '

  Private _H_POD6 As String '

  Private _H_POD7 As String '

  Private _H_POD8 As String '

  Private _H_POD9 As String '

  Private _H_POD10 As String '

  Private _H_POD11 As String '

  Private _H_POD12 As String '

  Private _H_POD13 As String '

  Private _H_POD14 As String '

  Private _H_POD15 As String '

  Private _H_POD16 As String '

  Private _H_POD17 As String '

  Private _H_POD18 As String '

  Private _H_POD19 As String '

  Private _H_POD20 As String '

  Private _H_POD21 As String '

  Private _H_POD22 As String '

  Private _H_POD23 As String '

  Private _H_POD24 As String '

  Private _H_POD25 As String '

  Private _SORT_ITEM_COMMON1 As String

  Private _SORT_ITEM_COMMON2 As String

  Private _SORT_ITEM_COMMON3 As String

  Private _SORT_ITEM_COMMON4 As String

  Private _SORT_ITEM_COMMON5 As String



  Public Property SORT_ITEM_COMMON5() As String
    Get
      Return _SORT_ITEM_COMMON5
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON5 = value
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


  Public Property SORT_ITEM_COMMON3() As String
    Get
      Return _SORT_ITEM_COMMON3
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON3 = value
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

  Public Property SORT_ITEM_COMMON1() As String
    Get
      Return _SORT_ITEM_COMMON1
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON1 = value
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
  Public Property PO_SERIAL_NO() As String
    Get
      Return _PO_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _PO_SERIAL_NO = value
    End Set
  End Property
  Public Property WORKING_TYPE() As String
    Get
      Return _WORKING_TYPE
    End Get
    Set(ByVal value As String)
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
  Public Property QTY() As Double
    Get
      Return _QTY
    End Get
    Set(ByVal value As Double)
      _QTY = value
    End Set
  End Property
  Public Property QTY_PROCESS() As Double
    Get
      Return _QTY_PROCESS
    End Get
    Set(ByVal value As Double)
      _QTY_PROCESS = value
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
  Public Property FROM_OWNER_ID() As String
    Get
      Return _FROM_OWNER_ID
    End Get
    Set(ByVal value As String)
      _FROM_OWNER_ID = value
    End Set
  End Property
  Public Property FROM_SUB_OWNER_ID() As String
    Get
      Return _FROM_SUB_OWNER_ID
    End Get
    Set(ByVal value As String)
      _FROM_SUB_OWNER_ID = value
    End Set
  End Property
  Public Property TO_OWNER_ID() As String
    Get
      Return _TO_OWNER_ID
    End Get
    Set(ByVal value As String)
      _TO_OWNER_ID = value
    End Set
  End Property
  Public Property TO_SUB_OWNER_ID() As String
    Get
      Return _TO_SUB_OWNER_ID
    End Get
    Set(ByVal value As String)
      _TO_SUB_OWNER_ID = value
    End Set
  End Property
  Public Property FACTORY_ID() As String
    Get
      Return _FACTORY_ID
    End Get
    Set(ByVal value As String)
      _FACTORY_ID = value
    End Set
  End Property
  Public Property DEST_AREA_ID() As String
    Get
      Return _DEST_AREA_ID
    End Get
    Set(ByVal value As String)
      _DEST_AREA_ID = value
    End Set
  End Property
  Public Property DEST_LOCATION_ID() As String
    Get
      Return _DEST_LOCATION_ID
    End Get
    Set(ByVal value As String)
      _DEST_LOCATION_ID = value
    End Set
  End Property
  Public Property H_POD_STEP_NO() As Double
    Get
      Return _H_POD_STEP_NO
    End Get
    Set(ByVal value As Double)
      _H_POD_STEP_NO = value
    End Set
  End Property
  Public Property H_POD_MOVE_TYPE() As String
    Get
      Return _H_POD_MOVE_TYPE
    End Get
    Set(ByVal value As String)
      _H_POD_MOVE_TYPE = value
    End Set
  End Property
  Public Property H_POD_FINISH_TIME() As String
    Get
      Return _H_POD_FINISH_TIME
    End Get
    Set(ByVal value As String)
      _H_POD_FINISH_TIME = value
    End Set
  End Property
  Public Property H_POD_BILLING_DATE() As String
    Get
      Return _H_POD_BILLING_DATE
    End Get
    Set(ByVal value As String)
      _H_POD_BILLING_DATE = value
    End Set
  End Property
  Public Property H_POD_CREATE_TIME() As String
    Get
      Return _H_POD_CREATE_TIME
    End Get
    Set(ByVal value As String)
      _H_POD_CREATE_TIME = value
    End Set
  End Property
  Public Property H_POD1() As String
    Get
      Return _H_POD1
    End Get
    Set(ByVal value As String)
      _H_POD1 = value
    End Set
  End Property
  Public Property H_POD2() As String
    Get
      Return _H_POD2
    End Get
    Set(ByVal value As String)
      _H_POD2 = value
    End Set
  End Property
  Public Property H_POD3() As String
    Get
      Return _H_POD3
    End Get
    Set(ByVal value As String)
      _H_POD3 = value
    End Set
  End Property
  Public Property H_POD4() As String
    Get
      Return _H_POD4
    End Get
    Set(ByVal value As String)
      _H_POD4 = value
    End Set
  End Property
  Public Property H_POD5() As String
    Get
      Return _H_POD5
    End Get
    Set(ByVal value As String)
      _H_POD5 = value
    End Set
  End Property
  Public Property H_POD6() As String
    Get
      Return _H_POD6
    End Get
    Set(ByVal value As String)
      _H_POD6 = value
    End Set
  End Property
  Public Property H_POD7() As String
    Get
      Return _H_POD7
    End Get
    Set(ByVal value As String)
      _H_POD7 = value
    End Set
  End Property
  Public Property H_POD8() As String
    Get
      Return _H_POD8
    End Get
    Set(ByVal value As String)
      _H_POD8 = value
    End Set
  End Property
  Public Property H_POD9() As String
    Get
      Return _H_POD9
    End Get
    Set(ByVal value As String)
      _H_POD9 = value
    End Set
  End Property
  Public Property H_POD10() As String
    Get
      Return _H_POD10
    End Get
    Set(ByVal value As String)
      _H_POD10 = value
    End Set
  End Property
  Public Property H_POD11() As String
    Get
      Return _H_POD11
    End Get
    Set(ByVal value As String)
      _H_POD11 = value
    End Set
  End Property
  Public Property H_POD12() As String
    Get
      Return _H_POD12
    End Get
    Set(ByVal value As String)
      _H_POD12 = value
    End Set
  End Property
  Public Property H_POD13() As String
    Get
      Return _H_POD13
    End Get
    Set(ByVal value As String)
      _H_POD13 = value
    End Set
  End Property
  Public Property H_POD14() As String
    Get
      Return _H_POD14
    End Get
    Set(ByVal value As String)
      _H_POD14 = value
    End Set
  End Property
  Public Property H_POD15() As String
    Get
      Return _H_POD15
    End Get
    Set(ByVal value As String)
      _H_POD15 = value
    End Set
  End Property
  Public Property H_POD16() As String
    Get
      Return _H_POD16
    End Get
    Set(ByVal value As String)
      _H_POD16 = value
    End Set
  End Property
  Public Property H_POD17() As String
    Get
      Return _H_POD17
    End Get
    Set(ByVal value As String)
      _H_POD17 = value
    End Set
  End Property
  Public Property H_POD18() As String
    Get
      Return _H_POD18
    End Get
    Set(ByVal value As String)
      _H_POD18 = value
    End Set
  End Property
  Public Property H_POD19() As String
    Get
      Return _H_POD19
    End Get
    Set(ByVal value As String)
      _H_POD19 = value
    End Set
  End Property
  Public Property H_POD20() As String
    Get
      Return _H_POD20
    End Get
    Set(ByVal value As String)
      _H_POD20 = value
    End Set
  End Property
  Public Property H_POD21() As String
    Get
      Return _H_POD21
    End Get
    Set(ByVal value As String)
      _H_POD21 = value
    End Set
  End Property
  Public Property H_POD22() As String
    Get
      Return _H_POD22
    End Get
    Set(ByVal value As String)
      _H_POD22 = value
    End Set
  End Property
  Public Property H_POD23() As String
    Get
      Return _H_POD23
    End Get
    Set(ByVal value As String)
      _H_POD23 = value
    End Set
  End Property
  Public Property H_POD24() As String
    Get
      Return _H_POD24
    End Get
    Set(ByVal value As String)
      _H_POD24 = value
    End Set
  End Property
  Public Property H_POD25() As String
    Get
      Return _H_POD25
    End Get
    Set(ByVal value As String)
      _H_POD25 = value
    End Set
  End Property
  Private _PODTL_STATUS As enuPODTLStatus
  Public Property PODTL_STATUS() As enuPODTLStatus
    Get
      Return _PODTL_STATUS
    End Get
    Set(ByVal value As enuPODTLStatus)
      _PODTL_STATUS = value
    End Set
  End Property
  Private _CLOSE_ABLE As Boolean
  Public Property CLOSE_ABLE() As Boolean
    Get
      Return _CLOSE_ABLE
    End Get
    Set(ByVal value As Boolean)
      _CLOSE_ABLE = value
    End Set
  End Property


  Public Sub New(ByVal PO_ID As String, ByVal PO_LINE_NO As String, ByVal PO_SERIAL_NO As String, ByVal WORKING_TYPE As String, ByVal WORKING_SERIAL_NO As String, WORKING_SERIAL_SEQ As String, ByVal SKU_NO As String, ByVal LOT_NO As String, ByVal QTY As Double,
                   ByVal QTY_PROCESS As Double, ByVal QTY_FINISH As Double, ByVal COMMENTS As String, ByVal PACKAGE_ID As String, ByVal ITEM_COMMON1 As String,
                   ByVal ITEM_COMMON2 As String, ByVal ITEM_COMMON3 As String, ByVal ITEM_COMMON4 As String, ByVal ITEM_COMMON5 As String, ByVal ITEM_COMMON6 As String,
                   ByVal ITEM_COMMON7 As String, ByVal ITEM_COMMON8 As String, ByVal ITEM_COMMON9 As String, ByVal ITEM_COMMON10 As String, ByVal SORT_ITEM_COMMON1 As String,
                   ByVal SORT_ITEM_COMMON2 As String, ByVal SORT_ITEM_COMMON3 As String, ByVal SORT_ITEM_COMMON4 As String, ByVal SORT_ITEM_COMMON5 As String, ByVal STORAGE_TYPE As Double,
                   ByVal BND As Double, ByVal QC_STATUS As Double, ByVal FROM_OWNER_ID As String, ByVal FROM_SUB_OWNER_ID As String, ByVal TO_OWNER_ID As String,
                   ByVal TO_SUB_OWNER_ID As String, ByVal FACTORY_ID As String, ByVal DEST_AREA_ID As String, ByVal DEST_LOCATION_ID As String, ByVal H_POD_STEP_NO As Double,
                   ByVal H_POD_MOVE_TYPE As String, ByVal H_POD_FINISH_TIME As String, ByVal H_POD_BILLING_DATE As String, ByVal H_POD_CREATE_TIME As String, ByVal H_POD1 As String,
                   ByVal H_POD2 As String, ByVal H_POD3 As String, ByVal H_POD4 As String, ByVal H_POD5 As String, ByVal H_POD6 As String, ByVal H_POD7 As String, ByVal H_POD8 As String,
                   ByVal H_POD9 As String, ByVal H_POD10 As String, ByVal H_POD11 As String, ByVal H_POD12 As String, ByVal H_POD13 As String, ByVal H_POD14 As String, ByVal H_POD15 As String,
                   ByVal H_POD16 As String, ByVal H_POD17 As String, ByVal H_POD18 As String, ByVal H_POD19 As String, ByVal H_POD20 As String, ByVal H_POD21 As String, ByVal H_POD22 As String,
                   ByVal H_POD23 As String, ByVal H_POD24 As String, ByVal H_POD25 As String, ByVal PODTL_STATUS As enuPODTLStatus, ByVal CLOSE_ABLE As Boolean)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PO_ID, PO_SERIAL_NO)
      _gid = key
      _PO_ID = PO_ID
      _PO_LINE_NO = PO_LINE_NO
      _PO_SERIAL_NO = PO_SERIAL_NO
      _WORKING_TYPE = WORKING_TYPE
      _WORKING_SERIAL_NO = _WORKING_SERIAL_NO
      _WORKING_SERIAL_SEQ = _WORKING_SERIAL_SEQ
      _SKU_NO = SKU_NO
      _LOT_NO = LOT_NO
      _QTY = QTY
      _QTY_PROCESS = QTY_PROCESS
      _QTY_FINISH = QTY_FINISH
      _PODTL_STATUS = PODTL_STATUS
      _COMMENTS = COMMENTS
      _PACKAGE_ID = PACKAGE_ID
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
      _STORAGE_TYPE = STORAGE_TYPE
      _BND = BND
      _QC_STATUS = QC_STATUS
      _FROM_OWNER_ID = FROM_OWNER_ID
      _FROM_SUB_OWNER_ID = FROM_SUB_OWNER_ID
      _TO_OWNER_ID = TO_OWNER_ID
      _TO_SUB_OWNER_ID = TO_SUB_OWNER_ID
      _FACTORY_ID = FACTORY_ID
      _DEST_AREA_ID = DEST_AREA_ID
      _DEST_LOCATION_ID = DEST_LOCATION_ID
      _H_POD_STEP_NO = H_POD_STEP_NO
      _H_POD_MOVE_TYPE = H_POD_MOVE_TYPE
      _H_POD_FINISH_TIME = H_POD_FINISH_TIME
      _H_POD_BILLING_DATE = H_POD_BILLING_DATE
      _H_POD_CREATE_TIME = H_POD_CREATE_TIME
      _H_POD1 = H_POD1
      _H_POD2 = H_POD2
      _H_POD3 = H_POD3
      _H_POD4 = H_POD4
      _H_POD5 = H_POD5
      _H_POD6 = H_POD6
      _H_POD7 = H_POD7
      _H_POD8 = H_POD8
      _H_POD9 = H_POD9
      _H_POD10 = H_POD10
      _H_POD11 = H_POD11
      _H_POD12 = H_POD12
      _H_POD13 = H_POD13
      _H_POD14 = H_POD14
      _H_POD15 = H_POD15
      _H_POD16 = H_POD16
      _H_POD17 = H_POD17
      _H_POD18 = H_POD18
      _H_POD19 = H_POD19
      _H_POD20 = H_POD20
      _H_POD21 = H_POD21
      _H_POD22 = H_POD22
      _H_POD23 = H_POD23
      _H_POD24 = H_POD24
      _H_POD25 = H_POD25
      _CLOSE_ABLE = CLOSE_ABLE
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
  Public Shared Function Get_Combination_Key(ByVal PO_ID As String, ByVal PO_SERIAL_NO As String) As String
    Try
      Dim key As String = PO_ID & LinkKey & PO_SERIAL_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPO_DTL
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
      Dim strSQL As String = WMS_T_PO_DTLManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_PO_DTLManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_PO_DTLManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_T_PO_DTL As clsPO_DTL) As Boolean
    Try
      Dim key As String = objWMS_T_PO_DTL._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _PO_ID = objWMS_T_PO_DTL.PO_ID
      _PO_LINE_NO = objWMS_T_PO_DTL.PO_LINE_NO
      _PO_SERIAL_NO = objWMS_T_PO_DTL.PO_SERIAL_NO
      _WORKING_TYPE = objWMS_T_PO_DTL.WORKING_TYPE
      _WORKING_SERIAL_NO = objWMS_T_PO_DTL.WORKING_SERIAL_NO
      _WORKING_SERIAL_SEQ = objWMS_T_PO_DTL.WORKING_SERIAL_SEQ
      _SKU_NO = objWMS_T_PO_DTL.SKU_NO
      _LOT_NO = objWMS_T_PO_DTL.LOT_NO
      _QTY = objWMS_T_PO_DTL.QTY
      _QTY_PROCESS = objWMS_T_PO_DTL.QTY_PROCESS
      _QTY_FINISH = objWMS_T_PO_DTL.QTY_FINISH
      _COMMENTS = objWMS_T_PO_DTL.COMMENTS
      _PACKAGE_ID = objWMS_T_PO_DTL.PACKAGE_ID
      _ITEM_COMMON1 = objWMS_T_PO_DTL.ITEM_COMMON1
      _ITEM_COMMON2 = objWMS_T_PO_DTL.ITEM_COMMON2
      _ITEM_COMMON3 = objWMS_T_PO_DTL.ITEM_COMMON3
      _ITEM_COMMON4 = objWMS_T_PO_DTL.ITEM_COMMON4
      _ITEM_COMMON5 = objWMS_T_PO_DTL.ITEM_COMMON5
      _ITEM_COMMON6 = objWMS_T_PO_DTL.ITEM_COMMON6
      _ITEM_COMMON7 = objWMS_T_PO_DTL.ITEM_COMMON7
      _ITEM_COMMON8 = objWMS_T_PO_DTL.ITEM_COMMON8
      _ITEM_COMMON9 = objWMS_T_PO_DTL.ITEM_COMMON9
      _ITEM_COMMON10 = objWMS_T_PO_DTL.ITEM_COMMON10
      _SORT_ITEM_COMMON1 = objWMS_T_PO_DTL.SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = objWMS_T_PO_DTL.SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = objWMS_T_PO_DTL.SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = objWMS_T_PO_DTL.SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = objWMS_T_PO_DTL.SORT_ITEM_COMMON5
      _STORAGE_TYPE = objWMS_T_PO_DTL.STORAGE_TYPE
      _BND = objWMS_T_PO_DTL.BND
      _QC_STATUS = objWMS_T_PO_DTL.QC_STATUS
      _FROM_OWNER_ID = objWMS_T_PO_DTL.FROM_OWNER_ID
      _FROM_SUB_OWNER_ID = objWMS_T_PO_DTL.FROM_SUB_OWNER_ID
      _TO_OWNER_ID = objWMS_T_PO_DTL.TO_OWNER_ID
      _TO_SUB_OWNER_ID = objWMS_T_PO_DTL.TO_SUB_OWNER_ID
      _FACTORY_ID = objWMS_T_PO_DTL.FACTORY_ID
      _DEST_AREA_ID = objWMS_T_PO_DTL.DEST_AREA_ID
      _DEST_LOCATION_ID = objWMS_T_PO_DTL.DEST_LOCATION_ID
      _H_POD_STEP_NO = objWMS_T_PO_DTL.H_POD_STEP_NO
      _H_POD_MOVE_TYPE = objWMS_T_PO_DTL.H_POD_MOVE_TYPE
      _H_POD_FINISH_TIME = objWMS_T_PO_DTL.H_POD_FINISH_TIME
      _H_POD_BILLING_DATE = objWMS_T_PO_DTL.H_POD_BILLING_DATE
      _H_POD_CREATE_TIME = objWMS_T_PO_DTL.H_POD_CREATE_TIME
      _H_POD1 = objWMS_T_PO_DTL.H_POD1
      _H_POD2 = objWMS_T_PO_DTL.H_POD2
      _H_POD3 = objWMS_T_PO_DTL.H_POD3
      _H_POD4 = objWMS_T_PO_DTL.H_POD4
      _H_POD5 = objWMS_T_PO_DTL.H_POD5
      _H_POD6 = objWMS_T_PO_DTL.H_POD6
      _H_POD7 = objWMS_T_PO_DTL.H_POD7
      _H_POD8 = objWMS_T_PO_DTL.H_POD8
      _H_POD9 = objWMS_T_PO_DTL.H_POD9
      _H_POD10 = objWMS_T_PO_DTL.H_POD10
      _H_POD11 = objWMS_T_PO_DTL.H_POD11
      _H_POD12 = objWMS_T_PO_DTL.H_POD12
      _H_POD13 = objWMS_T_PO_DTL.H_POD13
      _H_POD14 = objWMS_T_PO_DTL.H_POD14
      _H_POD15 = objWMS_T_PO_DTL.H_POD15
      _H_POD16 = objWMS_T_PO_DTL.H_POD16
      _H_POD17 = objWMS_T_PO_DTL.H_POD17
      _H_POD18 = objWMS_T_PO_DTL.H_POD18
      _H_POD19 = objWMS_T_PO_DTL.H_POD19
      _H_POD20 = objWMS_T_PO_DTL.H_POD20
      _H_POD21 = objWMS_T_PO_DTL.H_POD21
      _H_POD22 = objWMS_T_PO_DTL.H_POD22
      _H_POD23 = objWMS_T_PO_DTL.H_POD23
      _H_POD24 = objWMS_T_PO_DTL.H_POD24
      _H_POD25 = objWMS_T_PO_DTL.H_POD25
      _CLOSE_ABLE = objWMS_T_PO_DTL.CLOSE_ABLE
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_QTY_BY_FORCED(ByVal QTY As Long) As Boolean
    Try
      _QTY = QTY

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  Public Function Update_Data_By_Forced(ByRef objWMS_T_PO_DTL As clsPO_DTL) As Boolean
    Try
      Dim key As String = objWMS_T_PO_DTL._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '_PO_ID = objWMS_T_PO_DTL.PO_ID
      _PO_LINE_NO = objWMS_T_PO_DTL.PO_LINE_NO
      '_PO_SERIAL_NO = objWMS_T_PO_DTL.PO_SERIAL_NO
      _WORKING_TYPE = objWMS_T_PO_DTL.WORKING_TYPE
      _WORKING_SERIAL_NO = objWMS_T_PO_DTL.WORKING_SERIAL_NO
      _WORKING_SERIAL_SEQ = objWMS_T_PO_DTL.WORKING_SERIAL_SEQ
      _SKU_NO = objWMS_T_PO_DTL.SKU_NO
      _LOT_NO = objWMS_T_PO_DTL.LOT_NO
      _QTY = objWMS_T_PO_DTL.QTY
      '_QTY_PROCESS = objWMS_T_PO_DTL.QTY_PROCESS
      '_QTY_FINISH = objWMS_T_PO_DTL.QTY_FINISH
      _COMMENTS = objWMS_T_PO_DTL.COMMENTS
      _PACKAGE_ID = objWMS_T_PO_DTL.PACKAGE_ID
      _ITEM_COMMON1 = objWMS_T_PO_DTL.ITEM_COMMON1
      _ITEM_COMMON2 = objWMS_T_PO_DTL.ITEM_COMMON2
      _ITEM_COMMON3 = objWMS_T_PO_DTL.ITEM_COMMON3
      _ITEM_COMMON4 = objWMS_T_PO_DTL.ITEM_COMMON4
      _ITEM_COMMON5 = objWMS_T_PO_DTL.ITEM_COMMON5
      _ITEM_COMMON6 = objWMS_T_PO_DTL.ITEM_COMMON6
      _ITEM_COMMON7 = objWMS_T_PO_DTL.ITEM_COMMON7
      _ITEM_COMMON8 = objWMS_T_PO_DTL.ITEM_COMMON8
      _ITEM_COMMON9 = objWMS_T_PO_DTL.ITEM_COMMON9
      _ITEM_COMMON10 = objWMS_T_PO_DTL.ITEM_COMMON10
      _SORT_ITEM_COMMON1 = objWMS_T_PO_DTL.SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = objWMS_T_PO_DTL.SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = objWMS_T_PO_DTL.SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = objWMS_T_PO_DTL.SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = objWMS_T_PO_DTL.SORT_ITEM_COMMON5
      _STORAGE_TYPE = objWMS_T_PO_DTL.STORAGE_TYPE
      _BND = objWMS_T_PO_DTL.BND
      _QC_STATUS = objWMS_T_PO_DTL.QC_STATUS
      _FROM_OWNER_ID = objWMS_T_PO_DTL.FROM_OWNER_ID
      _FROM_SUB_OWNER_ID = objWMS_T_PO_DTL.FROM_SUB_OWNER_ID
      _TO_OWNER_ID = objWMS_T_PO_DTL.TO_OWNER_ID
      _TO_SUB_OWNER_ID = objWMS_T_PO_DTL.TO_SUB_OWNER_ID
      _FACTORY_ID = objWMS_T_PO_DTL.FACTORY_ID
      _DEST_AREA_ID = objWMS_T_PO_DTL.DEST_AREA_ID
      _DEST_LOCATION_ID = objWMS_T_PO_DTL.DEST_LOCATION_ID
      _H_POD_STEP_NO = objWMS_T_PO_DTL.H_POD_STEP_NO
      _H_POD_MOVE_TYPE = objWMS_T_PO_DTL.H_POD_MOVE_TYPE
      '_H_POD_FINISH_TIME = objWMS_T_PO_DTL.H_POD_FINISH_TIME
      '_H_POD_BILLING_DATE = objWMS_T_PO_DTL.H_POD_BILLING_DATE
      '_H_POD_CREATE_TIME = objWMS_T_PO_DTL.H_POD_CREATE_TIME
      _H_POD1 = objWMS_T_PO_DTL.H_POD1
      _H_POD2 = objWMS_T_PO_DTL.H_POD2
      _H_POD3 = objWMS_T_PO_DTL.H_POD3
      _H_POD4 = objWMS_T_PO_DTL.H_POD4
      _H_POD5 = objWMS_T_PO_DTL.H_POD5
      _H_POD6 = objWMS_T_PO_DTL.H_POD6
      _H_POD7 = objWMS_T_PO_DTL.H_POD7
      _H_POD8 = objWMS_T_PO_DTL.H_POD8
      _H_POD9 = objWMS_T_PO_DTL.H_POD9
      _H_POD10 = objWMS_T_PO_DTL.H_POD10
      _H_POD11 = objWMS_T_PO_DTL.H_POD11
      _H_POD12 = objWMS_T_PO_DTL.H_POD12
      _H_POD13 = objWMS_T_PO_DTL.H_POD13
      _H_POD14 = objWMS_T_PO_DTL.H_POD14
      _H_POD15 = objWMS_T_PO_DTL.H_POD15
      _H_POD16 = objWMS_T_PO_DTL.H_POD16
      _H_POD17 = objWMS_T_PO_DTL.H_POD17
      _H_POD18 = objWMS_T_PO_DTL.H_POD18
      _H_POD19 = objWMS_T_PO_DTL.H_POD19
      _H_POD20 = objWMS_T_PO_DTL.H_POD20
      _H_POD21 = objWMS_T_PO_DTL.H_POD21
      _H_POD22 = objWMS_T_PO_DTL.H_POD22
      _H_POD23 = objWMS_T_PO_DTL.H_POD23
      _H_POD24 = objWMS_T_PO_DTL.H_POD24
      _H_POD25 = objWMS_T_PO_DTL.H_POD25
      _CLOSE_ABLE = objWMS_T_PO_DTL.CLOSE_ABLE
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Class
