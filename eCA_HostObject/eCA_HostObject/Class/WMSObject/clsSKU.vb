Public Class clsSKU
  Private ShareName As String = "SKU"
  Private ShareKey As String = ""

  Private _gid As String    '等於 gSKU_NO
  Private _SKU_NO As String
  Private _SKU_ID1 As String
  Private _SKU_ID2 As String
  Private _SKU_ID3 As String
  Private _SKU_ALIS1 As String
  Private _SKU_ALIS2 As String
  Private _SKU_DESC As String
  Private _SKU_CATALOG As enuSKU_CATALOG
  Private _SKU_TYPE1 As String
  Private _SKU_TYPE2 As String
  Private _SKU_TYPE3 As String
  Private _SKU_COMMON1 As String
  Private _SKU_COMMON2 As String
  Private _SKU_COMMON3 As String
  Private _SKU_COMMON4 As String
  Private _SKU_COMMON5 As String
  Private _SKU_COMMON6 As String
  Private _SKU_COMMON7 As String
  Private _SKU_COMMON8 As String
  Private _SKU_COMMON9 As String
  Private _SKU_COMMON10 As String
  Private _SKU_L As Long
  Private _SKU_W As Long
  Private _SKU_H As Long
  Private _SKU_WEIGHT As Long
  Private _SKU_VALUE As Long
  Private _SKU_UNIT As String
  Private _INBOUND_UNIT As String
  Private _OUTBOUND_UNIT As String
  Private _HIGH_WATER As Long
  Private _LOW_WATER As Long
  Private _AVAILABLE_DAYS As Long
  Private _SAVE_DAYS As String
  Private _CREATE_TIME As String
  Private _UPDATE_TIME As String
  Private _WEIGHT_DIFFERENCE As Long
  Private _ENABLE As Boolean
  Private _EFFECTIVE_DATE As String
  Private _FAILURE_DATE As String
  Private _QC_METHOD As String
  Private _COMMENTS As String

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
  Public Property SKU_ID1() As String
    Get
      Return _SKU_ID1
    End Get
    Set(ByVal value As String)
      _SKU_ID1 = value
    End Set
  End Property
  Public Property SKU_ID2() As String
    Get
      Return _SKU_ID2
    End Get
    Set(ByVal value As String)
      _SKU_ID2 = value
    End Set
  End Property
  Public Property SKU_ID3() As String
    Get
      Return _SKU_ID3
    End Get
    Set(ByVal value As String)
      _SKU_ID3 = value
    End Set
  End Property
  Public Property SKU_ALIS1() As String
    Get
      Return _SKU_ALIS1
    End Get
    Set(ByVal value As String)
      _SKU_ALIS1 = value
    End Set
  End Property
  Public Property SKU_ALIS2() As String
    Get
      Return _SKU_ALIS2
    End Get
    Set(ByVal value As String)
      _SKU_ALIS2 = value
    End Set
  End Property
  Public Property SKU_DESC() As String
    Get
      Return _SKU_DESC
    End Get
    Set(ByVal value As String)
      _SKU_DESC = value
    End Set
  End Property
  Public Property SKU_CATALOG() As enuSKU_CATALOG
    Get
      Return _SKU_CATALOG
    End Get
    Set(ByVal value As enuSKU_CATALOG)
      _SKU_CATALOG = value
    End Set
  End Property
  Public Property SKU_TYPE1() As String
    Get
      Return _SKU_TYPE1
    End Get
    Set(ByVal value As String)
      _SKU_TYPE1 = value
    End Set
  End Property
  Public Property SKU_TYPE2() As String
    Get
      Return _SKU_TYPE2
    End Get
    Set(ByVal value As String)
      _SKU_TYPE2 = value
    End Set
  End Property
  Public Property SKU_TYPE3() As String
    Get
      Return _SKU_TYPE3
    End Get
    Set(ByVal value As String)
      _SKU_TYPE3 = value
    End Set
  End Property
  Public Property SKU_COMMON1() As String
    Get
      Return _SKU_COMMON1
    End Get
    Set(ByVal value As String)
      _SKU_COMMON1 = value
    End Set
  End Property
  Public Property SKU_COMMON2() As String
    Get
      Return _SKU_COMMON2
    End Get
    Set(ByVal value As String)
      _SKU_COMMON2 = value
    End Set
  End Property
  Public Property SKU_COMMON3() As String
    Get
      Return _SKU_COMMON3
    End Get
    Set(ByVal value As String)
      _SKU_COMMON3 = value
    End Set
  End Property
  Public Property SKU_COMMON4() As String
    Get
      Return _SKU_COMMON4
    End Get
    Set(ByVal value As String)
      _SKU_COMMON4 = value
    End Set
  End Property
  Public Property SKU_COMMON5() As String
    Get
      Return _SKU_COMMON5
    End Get
    Set(ByVal value As String)
      _SKU_COMMON5 = value
    End Set
  End Property
  Public Property SKU_COMMON6() As String
    Get
      Return _SKU_COMMON6
    End Get
    Set(ByVal value As String)
      _SKU_COMMON6 = value
    End Set
  End Property
  Public Property SKU_COMMON7() As String
    Get
      Return _SKU_COMMON7
    End Get
    Set(ByVal value As String)
      _SKU_COMMON7 = value
    End Set
  End Property
  Public Property SKU_COMMON8() As String
    Get
      Return _SKU_COMMON8
    End Get
    Set(ByVal value As String)
      _SKU_COMMON8 = value
    End Set
  End Property
  Public Property SKU_COMMON9() As String
    Get
      Return _SKU_COMMON9
    End Get
    Set(ByVal value As String)
      _SKU_COMMON9 = value
    End Set
  End Property
  Public Property SKU_COMMON10() As String
    Get
      Return _SKU_COMMON10
    End Get
    Set(ByVal value As String)
      _SKU_COMMON10 = value
    End Set
  End Property
  Public Property SKU_L() As Long
    Get
      Return _SKU_L
    End Get
    Set(ByVal value As Long)
      _SKU_L = value
    End Set
  End Property
  Public Property SKU_W() As Long
    Get
      Return _SKU_W
    End Get
    Set(ByVal value As Long)
      _SKU_W = value
    End Set
  End Property
  Public Property SKU_H() As Long
    Get
      Return _SKU_H
    End Get
    Set(ByVal value As Long)
      _SKU_H = value
    End Set
  End Property
  Public Property SKU_WEIGHT() As Long
    Get
      Return _SKU_WEIGHT
    End Get
    Set(ByVal value As Long)
      _SKU_WEIGHT = value
    End Set
  End Property
  Public Property SKU_VALUE() As Long
    Get
      Return _SKU_VALUE
    End Get
    Set(ByVal value As Long)
      _SKU_VALUE = value
    End Set
  End Property
  Public Property SKU_UNIT() As String
    Get
      Return _SKU_UNIT
    End Get
    Set(ByVal value As String)
      _SKU_UNIT = value
    End Set
  End Property
  Public Property INBOUND_UNIT() As String
    Get
      Return _INBOUND_UNIT
    End Get
    Set(ByVal value As String)
      _INBOUND_UNIT = value
    End Set
  End Property
  Public Property OUTBOUND_UNIT() As String
    Get
      Return _OUTBOUND_UNIT
    End Get
    Set(ByVal value As String)
      _OUTBOUND_UNIT = value
    End Set
  End Property
  Public Property HIGH_WATER() As Long
    Get
      Return _HIGH_WATER
    End Get
    Set(ByVal value As Long)
      _HIGH_WATER = value
    End Set
  End Property
  Public Property LOW_WATER() As Long
    Get
      Return _LOW_WATER
    End Get
    Set(ByVal value As Long)
      _LOW_WATER = value
    End Set
  End Property
  Public Property AVAILABLE_DAYS() As Long
    Get
      Return _AVAILABLE_DAYS
    End Get
    Set(ByVal value As Long)
      _AVAILABLE_DAYS = value
    End Set
  End Property
  Public Property SAVE_DAYS() As String
    Get
      Return _SAVE_DAYS
    End Get
    Set(ByVal value As String)
      _SAVE_DAYS = value
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
  Public Property WEIGHT_DIFFERENCE() As Long
    Get
      Return _WEIGHT_DIFFERENCE
    End Get
    Set(ByVal value As Long)
      _WEIGHT_DIFFERENCE = value
    End Set
  End Property
  Public Property ENABLE() As Boolean
    Get
      Return _ENABLE
    End Get
    Set(ByVal value As Boolean)
      _ENABLE = value
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
  Public Property FAILURE_DATE() As String
    Get
      Return _FAILURE_DATE
    End Get
    Set(ByVal value As String)
      _FAILURE_DATE = value
    End Set
  End Property
  Public Property QC_METHOD() As String
    Get
      Return _QC_METHOD
    End Get
    Set(ByVal value As String)
      _QC_METHOD = value
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

  '物件建立時執行的事件
  Public Sub New(ByVal SKU_NO As String, ByVal SKU_ID1 As String, ByVal SKU_ID2 As String, ByVal SKU_ID3 As String, ByVal SKU_ALIS1 As String,
                 ByVal SKU_ALIS2 As String, ByVal SKU_DESC As String, ByVal SKU_CATALOG As enuSKU_CATALOG, ByVal SKU_TYPE1 As String, ByVal SKU_TYPE2 As String,
                 ByVal SKU_TYPE3 As String, ByVal SKU_COMMON1 As String, ByVal SKU_COMMON2 As String, ByVal SKU_COMMON3 As String, ByVal SKU_COMMON4 As String,
                 ByVal SKU_COMMON5 As String, ByVal SKU_COMMON6 As String, ByVal SKU_COMMON7 As String, ByVal SKU_COMMON8 As String, ByVal SKU_COMMON9 As String,
                 ByVal SKU_COMMON10 As String, ByVal SKU_L As Long, ByVal SKU_W As Long, ByVal SKU_H As Long, ByVal SKU_WEIGHT As Long, ByVal SKU_VALUE As Long,
                 ByVal SKU_UNIT As String, ByVal INBOUND_UNIT As String, ByVal OUTBOUND_UNIT As String, ByVal HIGH_WATER As Long, ByVal LOW_WATER As Long,
                 ByVal AVAILABLE_DAYS As Long, ByVal SAVE_DAYS As String, ByVal CREATE_TIME As String, ByVal UPDATE_TIME As String, ByVal WEIGHT_DIFFERENCE As Long,
                 ByVal ENABLE As Boolean, ByVal EFFECTIVE_DATE As String, ByVal FAILURE_DATE As String, ByVal QC_METHOD As String, ByVal COMMENTS As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(SKU_NO)
      _gid = key
      _SKU_NO = SKU_NO
      _SKU_ID1 = SKU_ID1
      _SKU_ID2 = SKU_ID2
      _SKU_ID3 = SKU_ID3
      _SKU_ALIS1 = SKU_ALIS1
      _SKU_ALIS2 = SKU_ALIS2
      _SKU_DESC = SKU_DESC
      _SKU_CATALOG = SKU_CATALOG
      _SKU_TYPE1 = SKU_TYPE1
      _SKU_TYPE2 = SKU_TYPE2
      _SKU_TYPE3 = SKU_TYPE3
      _SKU_COMMON1 = SKU_COMMON1
      _SKU_COMMON2 = SKU_COMMON2
      _SKU_COMMON3 = SKU_COMMON3
      _SKU_COMMON4 = SKU_COMMON4
      _SKU_COMMON5 = SKU_COMMON5
      _SKU_COMMON6 = SKU_COMMON6
      _SKU_COMMON7 = SKU_COMMON7
      _SKU_COMMON8 = SKU_COMMON8
      _SKU_COMMON9 = SKU_COMMON9
      _SKU_COMMON10 = SKU_COMMON10
      _SKU_L = SKU_L
      _SKU_W = SKU_W
      _SKU_H = SKU_H
      _SKU_WEIGHT = SKU_WEIGHT
      _SKU_VALUE = SKU_VALUE
      _SKU_UNIT = SKU_UNIT
      _INBOUND_UNIT = INBOUND_UNIT
      _OUTBOUND_UNIT = OUTBOUND_UNIT
      _HIGH_WATER = HIGH_WATER
      _LOW_WATER = LOW_WATER
      _AVAILABLE_DAYS = AVAILABLE_DAYS
      _SAVE_DAYS = SAVE_DAYS
      _CREATE_TIME = CREATE_TIME
      _UPDATE_TIME = UPDATE_TIME
      _WEIGHT_DIFFERENCE = WEIGHT_DIFFERENCE
      _ENABLE = ENABLE
      _EFFECTIVE_DATE = EFFECTIVE_DATE
      _FAILURE_DATE = FAILURE_DATE
      _QC_METHOD = QC_METHOD
      _COMMENTS = COMMENTS
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
  Public Shared Function Get_Combination_Key(ByVal SKU_No As String) As String
    Try
      Dim key As String = SKU_No
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

  Public Function Clone() As clsSKU
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
      Dim strSQL As String = WMS_M_SKUManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_SKUManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_SKUManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
