Public Class clsClass
  Private ShareName As String = "Class"
  Private ShareKey As String = ""

  Private _gid As String
  Private _CLASS_NO As String '班別編號
  Private _CLASS_ID As String '班別 ID
  Private _CLASS_ALIS As String '班別名稱
  Private _CLASS_DESC As String '班別描述
  Private _CLASS_MANAGER As String '班別管理者(倉管人員)
  Private _PHONE As String '連絡電話
  Private _CLASS_START_TIME As String '班別起始時間
  Private _CLASS_END_TIME As String '班別結束時間

  Private _objHandling As clsHandlingObject


  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property

  Public Property CLASS_NO() As String
    Get
      Return _CLASS_NO
    End Get
    Set(ByVal value As String)
      _CLASS_NO = value
    End Set
  End Property

  Public Property CLASS_ID() As String
    Get
      Return _CLASS_ID
    End Get
    Set(ByVal value As String)
      _CLASS_ID = value
    End Set
  End Property

  Public Property CLASS_ALIS() As String
    Get
      Return _CLASS_ALIS
    End Get
    Set(ByVal value As String)
      _CLASS_ALIS = value
    End Set
  End Property

  Public Property CLASS_DESC() As String
    Get
      Return _CLASS_DESC
    End Get
    Set(ByVal value As String)
      _CLASS_DESC = value
    End Set
  End Property

  Public Property CLASS_MANAGER() As String
    Get
      Return _CLASS_MANAGER
    End Get
    Set(ByVal value As String)
      _CLASS_MANAGER = value
    End Set
  End Property

  Public Property PHONE() As String
    Get
      Return _PHONE
    End Get
    Set(ByVal value As String)
      _PHONE = value
    End Set
  End Property

  Public Property CLASS_START_TIME() As String
    Get
      Return _CLASS_START_TIME
    End Get
    Set(ByVal value As String)
      _CLASS_START_TIME = value
    End Set
  End Property

  Public Property CLASS_END_TIME() As String
    Get
      Return _CLASS_END_TIME
    End Get
    Set(ByVal value As String)
      _CLASS_END_TIME = value
    End Set
  End Property

  Public Property objHandling() As clsHandlingObject
    Get
      Return _objHandling
    End Get
    Set(ByVal value As clsHandlingObject)
      _objHandling = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal CLASS_NO As String, ByVal CLASS_ID As String, ByVal CLASS_ALIS As String, ByVal CLASS_DESC As String, ByVal CLASS_MANAGER As String,
                 ByVal PHONE As String, ByVal CLASS_START_TIME As String, ByVal CLASS_END_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(CLASS_NO)
      gid = key
      _CLASS_NO = CLASS_NO
      _CLASS_ID = CLASS_ID
      _CLASS_ALIS = CLASS_ALIS
      _CLASS_DESC = CLASS_DESC
      _CLASS_MANAGER = CLASS_MANAGER
      _PHONE = PHONE
      _CLASS_START_TIME = CLASS_START_TIME
      _CLASS_END_TIME = CLASS_END_TIME
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
    _objHandling = Nothing
  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal CLASS_NO As String) As String
    Try
      Dim key As String = CLASS_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsClass
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Sub Add_Relationship(ByRef objHandling As clsHandlingObject)
    Try
      '挷定Customer和WMS的關係
      If objHandling IsNot Nothing Then
        _objHandling = objHandling
        objHandling.O_Add_Class(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      '解除Block和WMS的關係
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_Class(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_M_ClassManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_ClassManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_ClassManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '資料加入Dictionary

  '把Device加入gcolDevice

  '資料從Dictionary刪除

  '取得Dictionary內的資料

  '非標準的Function
  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsClass) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> gid Then
        SendMessageToLog("Key can not Update, old_Key=" & gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _CLASS_NO = obj.CLASS_NO
      _CLASS_ID = obj.CLASS_ID
      _CLASS_ALIS = obj.CLASS_ALIS
      _CLASS_DESC = obj.CLASS_DESC
      _CLASS_MANAGER = obj.CLASS_MANAGER
      _PHONE = obj.PHONE
      _CLASS_START_TIME = obj.CLASS_START_TIME
      _CLASS_END_TIME = obj.CLASS_END_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Class
