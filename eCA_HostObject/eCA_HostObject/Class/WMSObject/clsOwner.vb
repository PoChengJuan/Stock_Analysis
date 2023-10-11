Public Class clsOwner
  Private ShareName As String = "WMS_M_Owner"
  Private ShareKey As String = ""
  Private _gid As String
  Private _OWNER_NO As String '貨主編號(系統使用的流水號)

  Private _OWNER_ID1 As String '貨主ID畫面上顯示用

  Private _OWNER_ID2 As String '貨主ID畫面上顯示用

  Private _OWNER_ID3 As String '貨主ID畫面上顯示用

  Private _OWNER_ALIS1 As String '貨主名稱畫面上顯示用

  Private _OWNER_ALIS2 As String '貨主名稱畫面上顯示用

  Private _OWNER_DESC As String '貨主描述

  Private _OWNER_TYPE As Double '貨主類型

  Private _OWNER_COMMON1 As String '通用欄位1

  Private _OWNER_COMMON2 As String '通用欄位2

  Private _OWNER_COMMON3 As String '通用欄位3

  Private _OWNER_COMMON4 As String '通用欄位4

  Private _OWNER_COMMON5 As String '通用欄位5

  Private _OWNER_COMMON6 As String '通用欄位6

  Private _OWNER_COMMON7 As String '通用欄位7

  Private _OWNER_COMMON8 As String '通用欄位8

  Private _OWNER_COMMON9 As String '通用欄位9

  Private _OWNER_COMMON10 As String '通用欄位10

  Private _COMMENTS As String '備註

  Private _ENABLE As Double '是否啟用0:禁用1:啟用

  Private _objHost As clsHandlingObject

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
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
  Public Property OWNER_ID1() As String
    Get
      Return _OWNER_ID1
    End Get
    Set(ByVal value As String)
      _OWNER_ID1 = value
    End Set
  End Property
  Public Property OWNER_ID2() As String
    Get
      Return _OWNER_ID2
    End Get
    Set(ByVal value As String)
      _OWNER_ID2 = value
    End Set
  End Property
  Public Property OWNER_ID3() As String
    Get
      Return _OWNER_ID3
    End Get
    Set(ByVal value As String)
      _OWNER_ID3 = value
    End Set
  End Property
  Public Property OWNER_ALIS1() As String
    Get
      Return _OWNER_ALIS1
    End Get
    Set(ByVal value As String)
      _OWNER_ALIS1 = value
    End Set
  End Property
  Public Property OWNER_ALIS2() As String
    Get
      Return _OWNER_ALIS2
    End Get
    Set(ByVal value As String)
      _OWNER_ALIS2 = value
    End Set
  End Property
  Public Property OWNER_DESC() As String
    Get
      Return _OWNER_DESC
    End Get
    Set(ByVal value As String)
      _OWNER_DESC = value
    End Set
  End Property
  Public Property OWNER_TYPE() As Double
    Get
      Return _OWNER_TYPE
    End Get
    Set(ByVal value As Double)
      _OWNER_TYPE = value
    End Set
  End Property
  Public Property OWNER_COMMON1() As String
    Get
      Return _OWNER_COMMON1
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON1 = value
    End Set
  End Property
  Public Property OWNER_COMMON2() As String
    Get
      Return _OWNER_COMMON2
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON2 = value
    End Set
  End Property
  Public Property OWNER_COMMON3() As String
    Get
      Return _OWNER_COMMON3
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON3 = value
    End Set
  End Property
  Public Property OWNER_COMMON4() As String
    Get
      Return _OWNER_COMMON4
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON4 = value
    End Set
  End Property
  Public Property OWNER_COMMON5() As String
    Get
      Return _OWNER_COMMON5
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON5 = value
    End Set
  End Property
  Public Property OWNER_COMMON6() As String
    Get
      Return _OWNER_COMMON6
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON6 = value
    End Set
  End Property
  Public Property OWNER_COMMON7() As String
    Get
      Return _OWNER_COMMON7
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON7 = value
    End Set
  End Property
  Public Property OWNER_COMMON8() As String
    Get
      Return _OWNER_COMMON8
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON8 = value
    End Set
  End Property
  Public Property OWNER_COMMON9() As String
    Get
      Return _OWNER_COMMON9
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON9 = value
    End Set
  End Property
  Public Property OWNER_COMMON10() As String
    Get
      Return _OWNER_COMMON10
    End Get
    Set(ByVal value As String)
      _OWNER_COMMON10 = value
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
  Public Property ENABLE() As Double
    Get
      Return _ENABLE
    End Get
    Set(ByVal value As Double)
      _ENABLE = value
    End Set
  End Property

  Public Sub New(ByVal OWNER_NO As String, ByVal OWNER_ID1 As String, ByVal OWNER_ID2 As String, ByVal OWNER_ID3 As String, ByVal OWNER_ALIS1 As String, ByVal OWNER_ALIS2 As String, ByVal OWNER_DESC As String, ByVal OWNER_TYPE As Double, ByVal OWNER_COMMON1 As String, ByVal OWNER_COMMON2 As String, ByVal OWNER_COMMON3 As String, ByVal OWNER_COMMON4 As String, ByVal OWNER_COMMON5 As String, ByVal OWNER_COMMON6 As String, ByVal OWNER_COMMON7 As String, ByVal OWNER_COMMON8 As String, ByVal OWNER_COMMON9 As String, ByVal OWNER_COMMON10 As String, ByVal COMMENTS As String, ByVal ENABLE As Double)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(OWNER_NO)
      _gid = key
      _OWNER_NO = OWNER_NO
      _OWNER_ID1 = OWNER_ID1
      _OWNER_ID2 = OWNER_ID2
      _OWNER_ID3 = OWNER_ID3
      _OWNER_ALIS1 = OWNER_ALIS1
      _OWNER_ALIS2 = OWNER_ALIS2
      _OWNER_DESC = OWNER_DESC
      _OWNER_TYPE = OWNER_TYPE
      _OWNER_COMMON1 = OWNER_COMMON1
      _OWNER_COMMON2 = OWNER_COMMON2
      _OWNER_COMMON3 = OWNER_COMMON3
      _OWNER_COMMON4 = OWNER_COMMON4
      _OWNER_COMMON5 = OWNER_COMMON5
      _OWNER_COMMON6 = OWNER_COMMON6
      _OWNER_COMMON7 = OWNER_COMMON7
      _OWNER_COMMON8 = OWNER_COMMON8
      _OWNER_COMMON9 = OWNER_COMMON9
      _OWNER_COMMON10 = OWNER_COMMON10
      _COMMENTS = COMMENTS
      _ENABLE = ENABLE
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
  Public Shared Function Get_Combination_Key(ByVal OWNER_NO As String) As String
    Try
      Dim key As String = OWNER_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsOwner
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
      Dim strSQL As String = WMS_M_OwnerManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_OwnerManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_OwnerManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_M_Owner As clsOwner) As Boolean
    Try
      Dim key As String = objWMS_M_Owner._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _OWNER_NO = OWNER_NO
      _OWNER_ID1 = OWNER_ID1
      _OWNER_ID2 = OWNER_ID2
      _OWNER_ID3 = OWNER_ID3
      _OWNER_ALIS1 = OWNER_ALIS1
      _OWNER_ALIS2 = OWNER_ALIS2
      _OWNER_DESC = OWNER_DESC
      _OWNER_TYPE = OWNER_TYPE
      _OWNER_COMMON1 = OWNER_COMMON1
      _OWNER_COMMON2 = OWNER_COMMON2
      _OWNER_COMMON3 = OWNER_COMMON3
      _OWNER_COMMON4 = OWNER_COMMON4
      _OWNER_COMMON5 = OWNER_COMMON5
      _OWNER_COMMON6 = OWNER_COMMON6
      _OWNER_COMMON7 = OWNER_COMMON7
      _OWNER_COMMON8 = OWNER_COMMON8
      _OWNER_COMMON9 = OWNER_COMMON9
      _OWNER_COMMON10 = OWNER_COMMON10
      _COMMENTS = COMMENTS
      _ENABLE = ENABLE
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Sub Add_Relationship(ByRef objHost As clsHandlingObject)
    Try
      '挷定WMS的關係
      If objHost IsNot Nothing Then
        _objHost = objHost
        objHost.O_Add_Owner(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      '解除WO和WMS的關係
      If _objHost IsNot Nothing Then
        _objHost.O_Remove_Owner(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
End Class
