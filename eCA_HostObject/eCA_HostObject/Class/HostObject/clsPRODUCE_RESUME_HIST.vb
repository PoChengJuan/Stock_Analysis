Public Class clsPRODUCE_RESUME_HIST
	Private ShareName As String = "HIST"
	Private ShareKey As String = ""
	Private _gid As String
	Private _WO_ID As String '工單號

	Private _WO_TYPE As Double '工單類型

	Private _PO_ID As String '訂單號

	Private _PO_TYPE1 As String '訂單類型

	Private _PO_TYPE2 As String '訂單類型

	Private _PO_TYPE3 As String '訂單類型

	Private _RELATED_WO As String '關係的工單號

	Private _CARRIER_ID As String '棧板號

	Private _PACKAGE_ID As String '箱ID/包裝ID

	Private _SKU_NO As String '貨品ID

	Private _LOT_NO As String '批號

	Private _ITEM_COMMON1 As String '條件1(供應商)

	Private _ITEM_COMMON2 As String '條件2

	Private _ITEM_COMMON3 As String '條件3

	Private _ITEM_COMMON4 As String '條件4

	Private _ITEM_COMMON5 As String '條件5

	Private _ITEM_COMMON6 As String '條件6

	Private _ITEM_COMMON7 As String '條件7

	Private _ITEM_COMMON8 As String '條件8

	Private _ITEM_COMMON9 As String '條件9

	Private _ITEM_COMMON10 As String '條件10

	Private _RECEIPT_DATE As String '收料日

	Private _MANUFACETURE_DATE As String '製造日

	Private _EXPIRED_DATE As String '到期日

	Private _EFFECTIVE_DATE As String '生效日(該日期之後才可以進行出庫)

	Private _CREATE_TIME As String '建立時間

	Private _QTY As Long '貨品數量

	Private _HIST_TIME As String '記錄時間

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
	Public Property WO_TYPE() As Double
		Get
			Return _WO_TYPE
		End Get
		Set(ByVal value As Double)
			_WO_TYPE = value
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
	Public Property PO_TYPE1() As String
		Get
			Return _PO_TYPE1
		End Get
		Set(ByVal value As String)
			_PO_TYPE1 = value
		End Set
	End Property
	Public Property PO_TYPE2() As String
		Get
			Return _PO_TYPE2
		End Get
		Set(ByVal value As String)
			_PO_TYPE2 = value
		End Set
	End Property
	Public Property PO_TYPE3() As String
		Get
			Return _PO_TYPE3
		End Get
		Set(ByVal value As String)
			_PO_TYPE3 = value
		End Set
	End Property
	Public Property RELATED_WO() As String
		Get
			Return _RELATED_WO
		End Get
		Set(ByVal value As String)
			_RELATED_WO = value
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
	Public Property PACKAGE_ID() As String
		Get
			Return _PACKAGE_ID
		End Get
		Set(ByVal value As String)
			_PACKAGE_ID = value
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
	Public Property EFFECTIVE_DATE() As String
		Get
			Return _EFFECTIVE_DATE
		End Get
		Set(ByVal value As String)
			_EFFECTIVE_DATE = value
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
	Public Property QTY() As Double
		Get
			Return _QTY
		End Get
		Set(ByVal value As Double)
			_QTY = value
		End Set
	End Property
	Public Property HIST_TIME() As String
		Get
			Return _HIST_TIME
		End Get
		Set(ByVal value As String)
			_HIST_TIME = value
		End Set
	End Property

	Public Sub New(ByVal WO_ID As String, ByVal WO_TYPE As Double, ByVal PO_ID As String, ByVal PO_TYPE1 As String, ByVal PO_TYPE2 As String, ByVal PO_TYPE3 As String, ByVal RELATED_WO As String, ByVal CARRIER_ID As String, ByVal PACKAGE_ID As String, ByVal SKU_NO As String, ByVal LOT_NO As String, ByVal ITEM_COMMON1 As String, ByVal ITEM_COMMON2 As String, ByVal ITEM_COMMON3 As String, ByVal ITEM_COMMON4 As String, ByVal ITEM_COMMON5 As String, ByVal ITEM_COMMON6 As String, ByVal ITEM_COMMON7 As String, ByVal ITEM_COMMON8 As String, ByVal ITEM_COMMON9 As String, ByVal ITEM_COMMON10 As String, ByVal RECEIPT_DATE As String, ByVal MANUFACETURE_DATE As String, ByVal EXPIRED_DATE As String, ByVal EFFECTIVE_DATE As String, ByVal CREATE_TIME As String, ByVal QTY As Long, ByVal HIST_TIME As String)
		MyBase.New()
		Try
			_WO_ID = WO_ID
			_WO_TYPE = WO_TYPE
			_PO_ID = PO_ID
			_PO_TYPE1 = PO_TYPE1
			_PO_TYPE2 = PO_TYPE2
			_PO_TYPE3 = PO_TYPE3
			_RELATED_WO = RELATED_WO
			_CARRIER_ID = CARRIER_ID
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
			_RECEIPT_DATE = RECEIPT_DATE
			_MANUFACETURE_DATE = MANUFACETURE_DATE
			_EXPIRED_DATE = EXPIRED_DATE
			_EFFECTIVE_DATE = EFFECTIVE_DATE
			_CREATE_TIME = CREATE_TIME
			_QTY = QTY
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
	Public Shared Function Get_Combination_Key(ByVal WO_ID As String, ByVal WO_TYPE As Double, ByVal PO_ID As String, ByVal PO_TYPE1 As String) As String
		Try
			Dim key As String = WO_ID & LinkKey & WO_TYPE & LinkKey & PO_ID & LinkKey & PO_TYPE1

			Return key
 Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return ""
		End Try
	End Function
	Public Function Clone() As clsPRODUCE_RESUME_HIST
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
			Dim strSQL As String = WMS_CH_PRODUCE_RESUME_HISTManagement.GetInsertSQL(Me)
			lstSQL.Add(strSQL)
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function

End Class
