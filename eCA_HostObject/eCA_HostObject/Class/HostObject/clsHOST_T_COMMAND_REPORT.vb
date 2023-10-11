Imports System.Collections.Concurrent
Public Class clsHOST_T_COMMAND_REPORT
  Private ShareName As String = "HOST_COMMAND_REPORT"
  Private ShareKey As String = ""
  Private _gid As String

  Private _UUID As String

  Private _REPORT_SYSTEM_TYPE As enuSystemType

  Private _REPORT_SYSTEM_UUID As String

  Private _CREATE_TIME As String

  Public _objHost As clsHandlingObject

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property UUID() As String
    Get
      Return _UUID
    End Get
    Set(ByVal value As String)
      _UUID = value
    End Set
  End Property
  Public Property REPORT_SYSTEM_TYPE() As enuSystemType
    Get
      Return _REPORT_SYSTEM_TYPE
    End Get
    Set(ByVal value As enuSystemType)
      _REPORT_SYSTEM_TYPE = value
    End Set
  End Property
  Public Property REPORT_SYSTEM_UUID() As String
    Get
      Return _REPORT_SYSTEM_UUID
    End Get
    Set(ByVal value As String)
      _REPORT_SYSTEM_UUID = value
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

  Public Sub New(ByVal UUID As String, ByVal REPORT_SYSTEM_TYPE As enuSystemType, ByVal REPORT_SYSTEM_UUID As String, ByVal CREATE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(REPORT_SYSTEM_TYPE, REPORT_SYSTEM_UUID)
      _gid = key
      _UUID = UUID
      _REPORT_SYSTEM_TYPE = REPORT_SYSTEM_TYPE
      _REPORT_SYSTEM_UUID = REPORT_SYSTEM_UUID
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
  '傳入指定參數取得Key值
  Public Function Clone() As clsALARM
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function Get_Combination_Key(ByVal REPORT_SYSTEM_TYPE As enuSystemType, ByVal REPORT_SYSTEM_UUID As String) As String
    Try
      Dim key As String = REPORT_SYSTEM_TYPE & LinkKey & REPORT_SYSTEM_UUID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = HOST_T_COMMAND_REPORTManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = HOST_T_COMMAND_REPORTManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '綁定物件和HOST的關係
  Public Sub Add_Relationship(ByRef objHost As clsHandlingObject)
    Try
      '挷定HOST的關係
      If objHost IsNot Nothing Then
        _objHost = objHost
        objHost.O_Add_HOST_T_COMMAND_REPORT(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '解除物件和HOST的關係
  Public Sub Remove_Relationship()
    Try
      '解除和HOST的關係
      If _objHost.gdicCommand_Report.ContainsKey(_gid) Then
        If _objHost.gdicCommand_Report.Remove(_gid) = False Then
          SendMessageToLog("TryRemove False", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return
        End If
      Else
        SendMessageToLog("Reomve ConcurrentDictionary Failed, key not exists key=" & gid, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '用另一個物件的資料，更新物件資料
  Public Function Update_To_Memory(ByRef obj As clsHOST_T_COMMAND_REPORT) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _UUID = obj.UUID
      _REPORT_SYSTEM_TYPE = obj.REPORT_SYSTEM_TYPE
      _REPORT_SYSTEM_UUID = obj.REPORT_SYSTEM_UUID
      _CREATE_TIME = obj.CREATE_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
