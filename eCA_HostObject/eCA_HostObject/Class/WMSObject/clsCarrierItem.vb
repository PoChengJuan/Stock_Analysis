Public Class clsCarrierItem
  '不諄記憶體建立關聯
  Private ShareName As String = "CarrierItem"
  Private ShareKey As String = ""

  Private _gid As String
  Private _Carrier_ID As String
  Private _Package_ID As String
  Private _SKU_NO As String
  Private _LOT_No As String
  Private _Item_Common1 As String
  Private _Item_Common2 As String
  Private _Item_Common3 As String
  Private _Item_Common4 As String
  Private _Item_Common5 As String
  Private _Item_Common6 As String
  Private _Item_Common7 As String
  Private _Item_Common8 As String
  Private _Item_Common9 As String
  Private _Item_Common10 As String
  Private _SORT_ITEM_COMMON1 As String
  Private _SORT_ITEM_COMMON2 As String
  Private _SORT_ITEM_COMMON3 As String
  Private _SORT_ITEM_COMMON4 As String
  Private _SORT_ITEM_COMMON5 As String
  Private _Qty As Double
  Private _Owner_No As String
  Private _Sub_Owner_No As String
  Private _Receipt_Date As String
  Private _Manufaceture_Date As String
  Private _Expired_Date As String
  Private _Org_Expired_Date As String
  Private _Expired_Comments As String
  Private _Expired_Warning_Flag As Boolean
  Private _Create_Time As String
  Private _In_Time As String
  Private _In_Client_ID As String
  Private _Receipt_WO_ID As String
  Private _Receipt_WO_Serial_No As String
  Private _Accepting_Status As enuAcceptingStatus
  Private _Storage_Type As enuStorageType
  Private _BND As Boolean
  Private _QC_Status As enuQCStatus
  Private _QC_Time As String
  Private _To_ERP As Boolean
  Private _To_ERP_Time As String
  Private _Effective_Date As String


  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
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
  Public Property SKU_No() As String
    Get
      Return _SKU_NO
    End Get
    Set(ByVal value As String)
      _SKU_NO = value
    End Set
  End Property
  Public Property Lot_No() As String
    Get
      Return _LOT_No
    End Get
    Set(ByVal value As String)
      _LOT_No = value
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
  Public Property QTY() As Double
    Get
      Return _Qty
    End Get
    Set(ByVal value As Double)
      _Qty = value
    End Set
  End Property
  Public Property Owner_No() As String
    Get
      Return _Owner_No
    End Get
    Set(ByVal value As String)
      _Owner_No = value
    End Set
  End Property
  Public Property Sub_Owner_No() As String
    Get
      Return _Sub_Owner_No
    End Get
    Set(ByVal value As String)
      _Sub_Owner_No = value
    End Set
  End Property
  Public Property Receipt_Date() As String
    Get
      Return _Receipt_Date
    End Get
    Set(ByVal value As String)
      _Receipt_Date = value
    End Set
  End Property
  Public Property Manufaceture_Date() As String
    Get
      Return _Manufaceture_Date
    End Get
    Set(ByVal value As String)
      _Manufaceture_Date = value
    End Set
  End Property
  Public Property Expired_Date() As String
    Get
      Return _Expired_Date
    End Get
    Set(ByVal value As String)
      _Expired_Date = value
    End Set
  End Property
  Public Property Org_Expired_Date() As String
    Get
      Return _Org_Expired_Date
    End Get
    Set(ByVal value As String)
      _Org_Expired_Date = value
    End Set
  End Property
  Public Property Expired_Comments() As String
    Get
      Return _Expired_Comments
    End Get
    Set(ByVal value As String)
      _Expired_Comments = value
    End Set
  End Property
  Public Property Expired_Warning_Flag() As Boolean
    Get
      Return _Expired_Warning_Flag
    End Get
    Set(ByVal value As Boolean)
      _Expired_Warning_Flag = value
    End Set
  End Property
  Public Property Create_Time() As String
    Get
      Return _Create_Time
    End Get
    Set(ByVal value As String)
      _Create_Time = value
    End Set
  End Property
  Public Property In_Time() As String
    Get
      Return _In_Time
    End Get
    Set(ByVal value As String)
      _In_Time = value
    End Set
  End Property
  Public Property In_Client_ID() As String
    Get
      Return _In_Client_ID
    End Get
    Set(ByVal value As String)
      _In_Client_ID = value
    End Set
  End Property
  Public Property Receipt_WO_ID() As String
    Get
      Return _Receipt_WO_ID
    End Get
    Set(ByVal value As String)
      _Receipt_WO_ID = value
    End Set
  End Property
  Public Property Receipt_WO_Serial_No() As String
    Get
      Return _Receipt_WO_Serial_No
    End Get
    Set(ByVal value As String)
      _Receipt_WO_Serial_No = value
    End Set
  End Property
  Public Property Accepting_Status() As enuAcceptingStatus
    Get
      Return _Accepting_Status
    End Get
    Set(ByVal value As enuAcceptingStatus)
      _Accepting_Status = value
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
  Public Property QC_Status() As enuQCStatus
    Get
      Return _QC_Status
    End Get
    Set(ByVal value As enuQCStatus)
      _QC_Status = value
    End Set
  End Property
  Public Property QC_Time() As String
    Get
      Return _QC_Time
    End Get
    Set(ByVal value As String)
      _QC_Time = value
    End Set
  End Property
  Public Property To_ERP() As Boolean
    Get
      Return _To_ERP
    End Get
    Set(ByVal value As Boolean)
      _To_ERP = value
    End Set
  End Property
  Public Property To_ERP_Time() As String
    Get
      Return _To_ERP_Time
    End Get
    Set(ByVal value As String)
      _To_ERP_Time = value
    End Set
  End Property
  Public Property Effective_Date() As String
    Get
      Return _Effective_Date
    End Get
    Set(ByVal value As String)
      _Effective_Date = value
    End Set
  End Property

  '物件建立時執行的事件
  Public Sub New(ByVal Carrier_ID As String, ByVal Package_ID As String, ByVal SKU_NO As String, ByVal LOT_No As String,
                 ByVal Item_Common1 As String, ByVal Item_Common2 As String, ByVal Item_Common3 As String,
                 ByVal Item_Common4 As String, ByVal Item_Common5 As String,
                 ByVal Item_Common6 As String, ByVal Item_Common7 As String, ByVal Item_Common8 As String,
                 ByVal Item_Common9 As String, ByVal Item_Common10 As String, ByVal Sort_Item_Common1 As String,
                 ByVal Sort_Item_Common2 As String, ByVal Sort_Item_Common3 As String, ByVal Sort_Item_Common4 As String, ByVal Sort_Item_Common5 As String,
                 ByVal Qty As Double, ByVal Owner_No As String, ByVal Sub_Owner_No As String, ByVal Receipt_Date As String, ByVal Manufaceture_Date As String,
                 ByVal Expired_Date As String, ByVal Org_Expired_Date As String, ByVal Expired_Comments As String,
                 ByVal Expired_Warning_Flag As Boolean, ByVal Create_Time As String,
                 ByVal In_Time As String, ByVal In_Client_ID As String, ByVal Receipt_WO_ID As String, ByVal Receipt_WO_Serial_No As String,
                 ByVal Accepting_Status As enuAcceptingStatus, ByVal Storage_Type As enuStorageType,
                 ByVal BND As Boolean, ByVal QC_Status As enuQCStatus, ByVal QC_Time As String,
                 ByVal To_ERP As Boolean, ByVal To_ERP_Time As String, ByVal Effective_Date As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(Carrier_ID, SKU_NO, Package_ID, LOT_No, Item_Common1, Item_Common2, Item_Common3,
                                              Item_Common4, Item_Common5, Item_Common6, Item_Common7, Item_Common8, Item_Common9,
                                              Item_Common10, Sort_Item_Common1, Sort_Item_Common2, Sort_Item_Common3, Sort_Item_Common4, Sort_Item_Common5,
                                              Receipt_WO_ID, Receipt_WO_Serial_No)
      _gid = key
      _Carrier_ID = Carrier_ID
      _Package_ID = Package_ID
      _SKU_NO = SKU_NO
      _LOT_No = LOT_No
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
      _SORT_ITEM_COMMON1 = Sort_Item_Common1
      _SORT_ITEM_COMMON2 = Sort_Item_Common2
      _SORT_ITEM_COMMON3 = Sort_Item_Common3
      _SORT_ITEM_COMMON4 = Sort_Item_Common4
      _SORT_ITEM_COMMON5 = Sort_Item_Common5
      _Qty = Qty
      _Owner_No = Owner_No
      _Sub_Owner_No = Sub_Owner_No
      _Receipt_Date = Receipt_Date
      _Manufaceture_Date = Manufaceture_Date
      _Expired_Date = Expired_Date
      _Org_Expired_Date = Org_Expired_Date
      _Expired_Comments = Expired_Comments
      _Expired_Warning_Flag = Expired_Warning_Flag
      _Create_Time = Create_Time
      _In_Time = In_Time
      _In_Client_ID = In_Client_ID
      _Receipt_WO_ID = Receipt_WO_ID
      _Receipt_WO_Serial_No = Receipt_WO_Serial_No
      _Accepting_Status = Accepting_Status
      _Storage_Type = Storage_Type
      _BND = BND
      _QC_Status = QC_Status
      _QC_Time = QC_Time
      _To_ERP = To_ERP
      _To_ERP_Time = To_ERP_Time
      _Effective_Date = Effective_Date
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '物件結束時觸發的事件，用來清除物件的內容
  Protected Overrides Sub Finalize()
    Class_Terminate_Renamed()
    MyBase.Finalize()
  End Sub
  Private Sub Class_Terminate_Renamed()
    '目的:結束物件
  End Sub

  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal CARRIER_ID As String, ByVal SKU_NO As String, ByVal PACKAGE_ID As String, ByVal LOT_NO As String,
                                             ByVal ITEM_COMMON1 As String, ByVal ITEM_COMMON2 As String, ByVal ITEM_COMMON3 As String,
                                             ByVal ITEM_COMMON4 As String, ByVal ITEM_COMMON5 As String, ByVal ITEM_COMMON6 As String,
                                             ByVal ITEM_COMMON7 As String, ByVal ITEM_COMMON8 As String, ByVal ITEM_COMMON9 As String,
                                             ByVal ITEM_COMMON10 As String, ByVal SORT_ITEM_COMMON1 As String, ByVal SORT_ITEM_COMMON2 As String, ByVal SORT_ITEM_COMMON3 As String,
                                             ByVal SORT_ITEM_COMMON4 As String, ByVal SORT_ITEM_COMMON5 As String, ByVal Receipt_WO_ID As String, ByVal Receipt_Serial_No As String) As String
    Try
      Dim key As String = CARRIER_ID & LinkKey & SKU_NO & LinkKey & PACKAGE_ID & LinkKey & LOT_NO & LinkKey & ITEM_COMMON1 & LinkKey & ITEM_COMMON2 & LinkKey & ITEM_COMMON3 & LinkKey & ITEM_COMMON4 & LinkKey & ITEM_COMMON5 & LinkKey & ITEM_COMMON6 & LinkKey & ITEM_COMMON7 & LinkKey & ITEM_COMMON8 & LinkKey & ITEM_COMMON9 & LinkKey & ITEM_COMMON10 & LinkKey & SORT_ITEM_COMMON1 & LinkKey & SORT_ITEM_COMMON2 & LinkKey & SORT_ITEM_COMMON3 & LinkKey & SORT_ITEM_COMMON4 & LinkKey & SORT_ITEM_COMMON5 & LinkKey & Receipt_WO_ID & LinkKey & Receipt_Serial_No
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsCarrierItem
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
      Dim strSQL As String = WMS_T_Carrier_ItemManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_Carrier_ItemManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_Carrier_ItemManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsCarrierItem) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _Carrier_ID = obj.Carrier_ID
      _Package_ID = obj.Package_ID
      _SKU_NO = obj.SKU_No
      _LOT_No = obj.Lot_No
      _Item_Common1 = obj.Item_Common1
      _Item_Common2 = obj.Item_Common2
      _Item_Common3 = obj.Item_Common3
      _Item_Common4 = obj.Item_Common4
      _Item_Common5 = obj.Item_Common5
      _Item_Common6 = obj.Item_Common6
      _Item_Common7 = obj.Item_Common7
      _Item_Common8 = obj.Item_Common8
      _Item_Common9 = obj.Item_Common9
      _Item_Common10 = obj.Item_Common10
      _SORT_ITEM_COMMON1 = obj.SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = obj.SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = obj.SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = obj.SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = obj.SORT_ITEM_COMMON5
      _Qty = obj.QTY
      _Owner_No = obj.Owner_No
      _Sub_Owner_No = obj.Sub_Owner_No
      _Receipt_Date = obj.Receipt_Date
      _Manufaceture_Date = obj.Manufaceture_Date
      _Expired_Date = obj.Expired_Date
      _Org_Expired_Date = obj.Org_Expired_Date
      _Expired_Comments = obj.Expired_Comments
      _Expired_Warning_Flag = obj.Expired_Warning_Flag
      _Create_Time = obj.Create_Time
      _In_Time = obj.In_Time
      _In_Client_ID = obj.In_Client_ID
      _Receipt_WO_ID = obj.Receipt_WO_ID
      _Receipt_WO_Serial_No = obj.Receipt_WO_Serial_No
      _Accepting_Status = obj.Accepting_Status
      _Storage_Type = obj._Storage_Type
      _BND = obj.BND
      _QC_Status = obj.QC_Status
      _QC_Time = obj.QC_Time
      _To_ERP = obj.To_ERP
      _To_ERP_Time = obj.To_ERP_Time
      _Effective_Date = obj.Effective_Date
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
