Public Class clsTSTOCKTAKINGDTL
  Private ShareName As String = "TSTOCKTAKINGDTL"
  Private ShareKey As String = ""
  Private _gid As String
  Private _STOCKTAKING_ID As String '盤點單號

  Private _STOCKTAKING_SERIAL_NO As String '盤點明細編號(流水號)

  Private _AREA_NO As String '指定的盤點倉別

  Private _BLOCK_NO As String '指定的盤點區塊

  Private _SKU_NO As String '指定的盤點料號

  Private _OWNER_NO As String '指定的盤點Owner

  Private _SUB_OWNER_NO As String '指定的盤點Sub_Owner

  Private _SL_NO As String '指定的盤點ERP儲存區域

  Private _STORAGE_TYPE As String '指定的盤點，一般品=1/暫存品=2

  Private _BND As String '指定的盤點，非保稅=0/保稅=1

  Private _CARRIER_ID As String '指定的盤點Carrier_ID

  Private _PERCENTAGE As Long '盤點百分比(1~100)

  Private _CARRIER_QTY As Long '需盤點棧板數

  Private _CARRIER_QTY_CHECKED As Long '已盤點棧板數

  Private _LOT_NO As String '指定的批號

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

  Private _RECEIPT_DATE As String '收料日

  Private _MANUFACETURE_DATE As String '製造日

  Private _EXPIRED_DATE As String '到期日

  Private _ERP_QTY As Double 'ERP傳入貨品數量

  Private _objWMS As clsHandlingObject
  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property STOCKTAKING_ID() As String
    Get
      Return _STOCKTAKING_ID
    End Get
    Set(ByVal value As String)
      _STOCKTAKING_ID = value
    End Set
  End Property
  Public Property STOCKTAKING_SERIAL_NO() As String
    Get
      Return _STOCKTAKING_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _STOCKTAKING_SERIAL_NO = value
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
  Public Property BLOCK_NO() As String
    Get
      Return _BLOCK_NO
    End Get
    Set(ByVal value As String)
      _BLOCK_NO = value
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
  Public Property SL_NO() As String
    Get
      Return _SL_NO
    End Get
    Set(ByVal value As String)
      _SL_NO = value
    End Set
  End Property
  Public Property STORAGE_TYPE() As String
    Get
      Return _STORAGE_TYPE
    End Get
    Set(ByVal value As String)
      _STORAGE_TYPE = value
    End Set
  End Property
  Public Property BND() As String
    Get
      Return _BND
    End Get
    Set(ByVal value As String)
      _BND = value
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
  Public Property PERCENTAGE() As Long
    Get
      Return _PERCENTAGE
    End Get
    Set(ByVal value As Long)
      _PERCENTAGE = value
    End Set
  End Property
  Public Property CARRIER_QTY() As Long
    Get
      Return _CARRIER_QTY
    End Get
    Set(ByVal value As Long)
      _CARRIER_QTY = value
    End Set
  End Property
  Public Property CARRIER_QTY_CHECKED() As Long
    Get
      Return _CARRIER_QTY_CHECKED
    End Get
    Set(ByVal value As Long)
      _CARRIER_QTY_CHECKED = value
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
  Public Property RECEIPT_DATE() As String
    Get
      Return _RECEIPT_DATE
    End Get
    Set(ByVal value As String)
      _RECEIPT_DATE = value
    End Set
  End Property
  Public Property MANUFACETURE_DATE() As String
    Get
      Return _MANUFACETURE_DATE
    End Get
    Set(ByVal value As String)
      _MANUFACETURE_DATE = value
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
  Public Property ERP_QTY() As Double
    Get
      Return _ERP_QTY
    End Get
    Set(ByVal value As Double)
      _ERP_QTY = value
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

  Public Sub New(ByVal STOCKTAKING_ID As String, ByVal STOCKTAKING_SERIAL_NO As String, ByVal AREA_NO As String, ByVal BLOCK_NO As String, ByVal SKU_NO As String, ByVal OWNER_NO As String, ByVal SUB_OWNER_NO As String, ByVal SL_NO As String, ByVal STORAGE_TYPE As String, ByVal BND As String, ByVal CARRIER_ID As String, ByVal PERCENTAGE As Long, ByVal CARRIER_QTY As Long, ByVal CARRIER_QTY_CHECKED As Long, ByVal LOT_NO As String, ByVal ITEM_COMMON1 As String, ByVal ITEM_COMMON2 As String, ByVal ITEM_COMMON3 As String, ByVal ITEM_COMMON4 As String, ByVal ITEM_COMMON5 As String, ByVal ITEM_COMMON6 As String, ByVal ITEM_COMMON7 As String, ByVal ITEM_COMMON8 As String, ByVal ITEM_COMMON9 As String, ByVal ITEM_COMMON10 As String, ByVal SORT_ITEM_COMMON1 As String, ByVal SORT_ITEM_COMMON2 As String, ByVal SORT_ITEM_COMMON3 As String, ByVal SORT_ITEM_COMMON4 As String, ByVal SORT_ITEM_COMMON5 As String, ByVal SUPPLIER_NO As String, ByVal CUSTOMER_NO As String, ByVal RECEIPT_DATE As String, ByVal MANUFACETURE_DATE As String, ByVal EXPIRED_DATE As String, ByVal ERP_QTY As Double)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(STOCKTAKING_ID, STOCKTAKING_SERIAL_NO)
      _gid = key
      _STOCKTAKING_ID = STOCKTAKING_ID
      _STOCKTAKING_SERIAL_NO = STOCKTAKING_SERIAL_NO
      _AREA_NO = AREA_NO
      _BLOCK_NO = BLOCK_NO
      _SKU_NO = SKU_NO
      _OWNER_NO = OWNER_NO
      _SUB_OWNER_NO = SUB_OWNER_NO
      _SL_NO = SL_NO
      _STORAGE_TYPE = STORAGE_TYPE
      _BND = BND
      _CARRIER_ID = CARRIER_ID
      _PERCENTAGE = PERCENTAGE
      _CARRIER_QTY = CARRIER_QTY
      _CARRIER_QTY_CHECKED = CARRIER_QTY_CHECKED
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
      _RECEIPT_DATE = RECEIPT_DATE
      _MANUFACETURE_DATE = MANUFACETURE_DATE
      _EXPIRED_DATE = EXPIRED_DATE
      _ERP_QTY = ERP_QTY
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
  Public Shared Function Get_Combination_Key(ByVal STOCKTAKING_ID As String, ByVal STOCKTAKING_SERIAL_NO As String) As String
    Try
      Dim key As String = STOCKTAKING_ID & LinkKey & STOCKTAKING_SERIAL_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsTSTOCKTAKINGDTL
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
      Dim strSQL As String = WMS_T_STOCKTAKING_DTLManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_STOCKTAKING_DTLManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_STOCKTAKING_DTLManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_T_STOCKTAKING_DTL As clsTSTOCKTAKINGDTL) As Boolean
    Try
      Dim key As String = objWMS_T_STOCKTAKING_DTL._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _STOCKTAKING_ID = objWMS_T_STOCKTAKING_DTL.STOCKTAKING_ID
      _STOCKTAKING_SERIAL_NO = objWMS_T_STOCKTAKING_DTL.STOCKTAKING_SERIAL_NO
      _AREA_NO = objWMS_T_STOCKTAKING_DTL.AREA_NO
      _BLOCK_NO = objWMS_T_STOCKTAKING_DTL.BLOCK_NO
      _SKU_NO = objWMS_T_STOCKTAKING_DTL.SKU_NO
      _OWNER_NO = objWMS_T_STOCKTAKING_DTL.OWNER_NO
      _SUB_OWNER_NO = objWMS_T_STOCKTAKING_DTL.SUB_OWNER_NO
      _SL_NO = objWMS_T_STOCKTAKING_DTL.SL_NO
      _STORAGE_TYPE = objWMS_T_STOCKTAKING_DTL.STORAGE_TYPE
      _BND = objWMS_T_STOCKTAKING_DTL.BND
      _CARRIER_ID = objWMS_T_STOCKTAKING_DTL.CARRIER_ID
      _PERCENTAGE = objWMS_T_STOCKTAKING_DTL.PERCENTAGE
      _CARRIER_QTY = objWMS_T_STOCKTAKING_DTL.CARRIER_QTY
      _CARRIER_QTY_CHECKED = objWMS_T_STOCKTAKING_DTL.CARRIER_QTY_CHECKED
      _LOT_NO = objWMS_T_STOCKTAKING_DTL.LOT_NO
      _ITEM_COMMON1 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON1
      _ITEM_COMMON2 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON2
      _ITEM_COMMON3 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON3
      _ITEM_COMMON4 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON4
      _ITEM_COMMON5 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON5
      _ITEM_COMMON6 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON6
      _ITEM_COMMON7 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON7
      _ITEM_COMMON8 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON8
      _ITEM_COMMON9 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON9
      _ITEM_COMMON10 = objWMS_T_STOCKTAKING_DTL.ITEM_COMMON10
      _SORT_ITEM_COMMON1 = objWMS_T_STOCKTAKING_DTL.SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = objWMS_T_STOCKTAKING_DTL.SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = objWMS_T_STOCKTAKING_DTL.SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = objWMS_T_STOCKTAKING_DTL.SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = objWMS_T_STOCKTAKING_DTL.SORT_ITEM_COMMON5
      _SUPPLIER_NO = objWMS_T_STOCKTAKING_DTL.SUPPLIER_NO
      _CUSTOMER_NO = objWMS_T_STOCKTAKING_DTL.CUSTOMER_NO
      _RECEIPT_DATE = objWMS_T_STOCKTAKING_DTL.RECEIPT_DATE
      _MANUFACETURE_DATE = objWMS_T_STOCKTAKING_DTL.MANUFACETURE_DATE
      _EXPIRED_DATE = objWMS_T_STOCKTAKING_DTL.EXPIRED_DATE
      _ERP_QTY = objWMS_T_STOCKTAKING_DTL.ERP_QTY
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
