Public Class clsSystemStatus
  Private ShareName As String = "SystemStatus"
  Private ShareKey As String = ""

  Private _gid As String
  Private _STATUS_NO As enuSystemStatus '狀態編號
  Private _STATUS_NAME As String '狀態名稱
  Private _STATUS_VALUE As String '狀態值
  Private _UPDATE_TIME As String '更新時間

  Private _STATUS_MODE As enuStatusMode
  Private _STATUS_TYPE1 As String
  Private _STATUS_TYPE2 As String
  Private _STATUS_TYPE3 As String
  Private _STATUS_DESC As String




  Private _objHandling As clsHandlingObject

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property

  Public Property STATUS_NO() As enuSystemStatus
    Get
      Return _STATUS_NO
    End Get
    Set(ByVal value As enuSystemStatus)
      _STATUS_NO = value
    End Set
  End Property

  Public Property STATUS_NAME() As String
    Get
      Return _STATUS_NAME
    End Get
    Set(ByVal value As String)
      _STATUS_NAME = value
    End Set
  End Property

  Public Property STATUS_VALUE() As String
    Get
      Return _STATUS_VALUE
    End Get
    Set(ByVal value As String)
      _STATUS_VALUE = value
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

  Public Property STATUS_MODE() As enuStatusMode
    Get
      Return _STATUS_MODE
    End Get
    Set(ByVal value As enuStatusMode)
      _STATUS_MODE = value
    End Set
  End Property
  Public Property STATUS_TYPE1() As String
    Get
      Return _STATUS_TYPE1
    End Get
    Set(ByVal value As String)
      _STATUS_TYPE1 = value
    End Set
  End Property
  Public Property STATUS_TYPE2() As String
    Get
      Return _STATUS_TYPE2
    End Get
    Set(ByVal value As String)
      _STATUS_TYPE2 = value
    End Set
  End Property
  Public Property STATUS_TYPE3() As String
    Get
      Return _STATUS_TYPE3
    End Get
    Set(ByVal value As String)
      _STATUS_TYPE3 = value
    End Set
  End Property
  Public Property STATUS_DESC() As String
    Get
      Return _STATUS_DESC
    End Get
    Set(ByVal value As String)
      _STATUS_DESC = value
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
  Public Sub New(ByVal STATUS_NO As enuSystemStatus,
                 ByVal STATUS_NAME As String,
                 ByVal STATUS_VALUE As String,
                 ByVal UPDATE_TIME As String,
                 ByVal STATUS_MODE As enuStatusMode,
                 ByVal STATUS_TYPE1 As String,
                 ByVal STATUS_TYPE2 As String,
                 ByVal STATUS_TYPE3 As String,
                 ByVal STATUS_DESC As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(STATUS_NO)
      gid = key
      _STATUS_NO = STATUS_NO
      _STATUS_NAME = STATUS_NAME
      _STATUS_VALUE = STATUS_VALUE
      _UPDATE_TIME = UPDATE_TIME
      _STATUS_MODE = STATUS_MODE
      _STATUS_TYPE1 = STATUS_TYPE1
      _STATUS_TYPE2 = STATUS_TYPE2
      _STATUS_TYPE3 = STATUS_TYPE3
      _STATUS_DESC = STATUS_DESC

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
  Public Shared Function Get_Combination_Key(ByVal STATUS_NO As String) As String
    Try
      Dim key As String = STATUS_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsSystemStatus
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Sub Add_Relationship(ByRef objHandling As clsHandlingObject)
    Try
      If objHandling IsNot Nothing Then
        _objHandling = objHandling
        objHandling.O_Add_SystemStatus(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_SystemStatus(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_SystemStatusManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_SystemStatusManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_SystemStatusManagement.GetDeleteSQL(Me)
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


  '-供他人使用的GET
  '-取得gid

  '非標準的Function
  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsSystemStatus) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> gid Then
        SendMessageToLog("Key can not Update, old_Key=" & gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _STATUS_NAME = obj.STATUS_NAME
      _STATUS_VALUE = obj.STATUS_VALUE
      _UPDATE_TIME = obj.UPDATE_TIME
      _STATUS_MODE = obj.STATUS_MODE
      _STATUS_TYPE1 = obj.STATUS_TYPE1
      _STATUS_TYPE2 = obj.STATUS_TYPE2
      _STATUS_TYPE3 = obj.STATUS_TYPE3
      _STATUS_DESC = obj.STATUS_DESC

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Class
