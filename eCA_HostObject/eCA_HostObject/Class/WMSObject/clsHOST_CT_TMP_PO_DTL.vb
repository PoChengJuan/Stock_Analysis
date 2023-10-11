Public Class clsHOST_CT_TMP_PO_DTL
  Private ShareName As String = "HOST_CT_TMP_PO_DTL"
  Private ShareKey As String = ""
  Private _gid As String
  Private _PO_ID As String '訂單編號

  Private _PO_LINE_NO As String '訂單明細編號(上傳時使用)

  Private _PO_SERIAL_NO As String '訂單明細編號(WMS使用)

  Private _WO_TYPE As String '工單類型

  Private _SKU_NO As String '貨品編號

  Private _LOT_NO As String '批號

  Private _QTY As Double '需求量

  Private _OWNER_ID As String '貨主編號

  Private _SUB_OWNER_ID As String '子貨主編號

  Private _COMMON1 As String '條件1

  Private _COMMON2 As String '條件2

  Private _COMMON3 As String '條件3

  Private _COMMON4 As String '條件4

  Private _COMMON5 As String '條件5

  Private _COMMENTS As String '備註

  Private _CREATE_TIME As String '時間


  Private _objHost As clsHandlingObject


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
  Public Property WO_TYPE() As String
    Get
      Return _WO_TYPE
    End Get
    Set(ByVal value As String)
      _WO_TYPE = value
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
  Public Property OWNER_ID() As String
    Get
      Return _OWNER_ID
    End Get
    Set(ByVal value As String)
      _OWNER_ID = value
    End Set
  End Property
  Public Property SUB_OWNER_ID() As String
    Get
      Return _SUB_OWNER_ID
    End Get
    Set(ByVal value As String)
      _SUB_OWNER_ID = value
    End Set
  End Property
  Public Property COMMON1() As String
    Get
      Return _COMMON1
    End Get
    Set(ByVal value As String)
      _COMMON1 = value
    End Set
  End Property
  Public Property COMMON2() As String
    Get
      Return _COMMON2
    End Get
    Set(ByVal value As String)
      _COMMON2 = value
    End Set
  End Property
  Public Property COMMON3() As String
    Get
      Return _COMMON3
    End Get
    Set(ByVal value As String)
      _COMMON3 = value
    End Set
  End Property
  Public Property COMMON4() As String
    Get
      Return _COMMON4
    End Get
    Set(ByVal value As String)
      _COMMON4 = value
    End Set
  End Property
  Public Property COMMON5() As String
    Get
      Return _COMMON5
    End Get
    Set(ByVal value As String)
      _COMMON5 = value
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
  Public Property CREATE_TIME() As String
    Get
      Return _CREATE_TIME
    End Get
    Set(ByVal value As String)
      _CREATE_TIME = value
    End Set
  End Property

  Public Property objHost() As clsHandlingObject
    Get
      Return _objHost
    End Get
    Set(ByVal value As clsHandlingObject)
      _objHost = value
    End Set
  End Property

  Public Sub New(ByVal PO_ID As String, ByVal PO_LINE_NO As String, ByVal PO_SERIAL_NO As String, ByVal WO_TYPE As String, ByVal SKU_NO As String, ByVal LOT_NO As String, ByVal QTY As Double, ByVal OWNER_ID As String, ByVal SUB_OWNER_ID As String, ByVal COMMON1 As String, ByVal COMMON2 As String, ByVal COMMON3 As String, ByVal COMMON4 As String, ByVal COMMON5 As String, ByVal COMMENTS As String, ByVal CREATE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PO_ID, PO_SERIAL_NO)
      _gid = key
      _PO_ID = PO_ID
      _PO_LINE_NO = PO_LINE_NO
      _PO_SERIAL_NO = PO_SERIAL_NO
      _WO_TYPE = WO_TYPE
      _SKU_NO = SKU_NO
      _LOT_NO = LOT_NO
      _QTY = QTY
      _OWNER_ID = OWNER_ID
      _SUB_OWNER_ID = SUB_OWNER_ID
      _COMMON1 = COMMON1
      _COMMON2 = COMMON2
      _COMMON3 = COMMON3
      _COMMON4 = COMMON4
      _COMMON5 = COMMON5
      _COMMENTS = COMMENTS
      _CREATE_TIME = CREATE_TIME
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
  Public Sub Add_Relationship(ByRef objHost As clsHandlingObject)
    Try
      '挷定WMS的關係
      If objHost IsNot Nothing Then
        _objHost = objHost
        objHost.O_Add_CT_PO_DTL(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      '解除WO和WMS的關係
      If _objHost IsNot Nothing Then
        _objHost.O_Remove_CT_PO_DTL(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Function Clone() As clsHOST_CT_TMP_PO_DTL
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
      Dim strSQL As String = HOST_CT_TMP_PO_DTLManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = HOST_CT_TMP_PO_DTLManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = HOST_CT_TMP_PO_DTLManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objHOST_CT_TMP_PO_DTL As clsHOST_CT_TMP_PO_DTL) As Boolean
    Try
      Dim key As String = objHOST_CT_TMP_PO_DTL._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _PO_ID = PO_ID
      _PO_LINE_NO = PO_LINE_NO
      _PO_SERIAL_NO = PO_SERIAL_NO
      _WO_TYPE = WO_TYPE
      _SKU_NO = SKU_NO
      _LOT_NO = LOT_NO
      _QTY = QTY
      _OWNER_ID = OWNER_ID
      _SUB_OWNER_ID = SUB_OWNER_ID
      _COMMON1 = COMMON1
      _COMMON2 = COMMON2
      _COMMON3 = COMMON3
      _COMMON4 = COMMON4
      _COMMON5 = COMMON5
      _COMMENTS = COMMENTS
      _CREATE_TIME = CREATE_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
