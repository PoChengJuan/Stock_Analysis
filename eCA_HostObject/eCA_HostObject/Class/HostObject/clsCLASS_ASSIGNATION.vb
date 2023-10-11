Public Class clsCLASS_ASSIGNATION
  Private ShareName As String = "CLASS_ASSIGNATION"
  Private ShareKey As String = ""
  Private _gid As String
  Private _FACTORY_NO As String '廠別

  Private _AREA_NO As String '區域編號

  Private _CLASS_NO As String '班別編號

  Private _ASSIGNATION_RATE As Double '分配比率(%)

  Private _UPDATE_USER As String '更新人員

  Private _UPDATE_TIME As String '更新時間

  Private _objHandling As clsHandlingObject

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property FACTORY_NO() As String
    Get
      Return _FACTORY_NO
    End Get
    Set(ByVal value As String)
      _FACTORY_NO = value
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
  Public Property CLASS_NO() As String
    Get
      Return _CLASS_NO
    End Get
    Set(ByVal value As String)
      _CLASS_NO = value
    End Set
  End Property
  Public Property ASSIGNATION_RATE() As Double
    Get
      Return _ASSIGNATION_RATE
    End Get
    Set(ByVal value As Double)
      _ASSIGNATION_RATE = value
    End Set
  End Property
  Public Property UPDATE_USER() As String
    Get
      Return _UPDATE_USER
    End Get
    Set(ByVal value As String)
      _UPDATE_USER = value
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
  Public Property objHandling() As clsHandlingObject
    Get
      Return _objHandling
    End Get
    Set(ByVal value As clsHandlingObject)
      _objHandling = value
    End Set
  End Property

  Public Sub New(ByVal FACTORY_NO As String, ByVal AREA_NO As String, ByVal CLASS_NO As String, ByVal ASSIGNATION_RATE As Double, ByVal UPDATE_USER As String, ByVal UPDATE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(FACTORY_NO, AREA_NO, CLASS_NO)
      _gid = key
      _FACTORY_NO = FACTORY_NO
      _AREA_NO = AREA_NO
      _CLASS_NO = CLASS_NO
      _ASSIGNATION_RATE = ASSIGNATION_RATE
      _UPDATE_USER = UPDATE_USER
      _UPDATE_TIME = UPDATE_TIME
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
  Public Sub Add_Relationship(ByRef objHandling As clsHandlingObject)
    Try
      '挷定Customer和WMS的關係
      If objHandling IsNot Nothing Then
        _objHandling = objHandling
        objHandling.O_Add_ClassAssignation(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      '解除Block和WMS的關係
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_ClassAssignation(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal FACTORY_NO As String, ByVal AREA_NO As String, ByVal CLASS_NO As String) As String
    Try
      Dim key As String = FACTORY_NO & LinkKey & AREA_NO & LinkKey & CLASS_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsCLASS_ASSIGNATION
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
      Dim strSQL As String = WMS_CM_CLASS_ASSIGNATIONManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_CM_CLASS_ASSIGNATIONManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_CM_CLASS_ASSIGNATIONManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef obj As clsCLASS_ASSIGNATION) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _FACTORY_NO = FACTORY_NO
      _AREA_NO = AREA_NO
      _CLASS_NO = CLASS_NO
      _ASSIGNATION_RATE = ASSIGNATION_RATE
      _UPDATE_USER = UPDATE_USER
      _UPDATE_TIME = UPDATE_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
