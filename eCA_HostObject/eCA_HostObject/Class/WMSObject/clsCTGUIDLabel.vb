Public Class clsCTGUIDLabel
  Private ShareName As String = "CTGUIDLabel"
  Private ShareKey As String = ""
  Private _gid As String
  Private _GUID As String '

  Private _UniqueID As String '

  Private _objWMS As clsHandlingObject
  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property GUID() As String
    Get
      Return _GUID
    End Get
    Set(ByVal value As String)
      _GUID = value
    End Set
  End Property
  Public Property UniqueID() As String
    Get
      Return _UniqueID
    End Get
    Set(ByVal value As String)
      _UniqueID = value
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

  Public Sub New(ByVal GUID As String, ByVal UniqueID As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(GUID)
      _gid = key
      _GUID = GUID
      _UniqueID = UniqueID
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
  Public Shared Function Get_Combination_Key(ByVal GUID As String) As String
    Try
      Dim key As String = GUID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsCTGUIDLabel
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  'Public Sub Add_Relationship(ByRef objWMS As clsWMSObject)
  '  Try
  '    '挷定WMS的關係                                                                        
  '    If objWMS IsNot Nothing Then
  '      _objWMS = objWMS
  '      objWMS.O_Add_!!!!!這邊就是你要改的東西啦(Me)
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub
  'Public Sub Remove_Relationship()
  '  Try
  '    If _objWMS IsNot Nothing Then
  '      _objWMS.O_Remove_!!!!!這也是你要改的東西
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_CT_GUID_LabelManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_CT_GUID_LabelManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_CT_GUID_LabelManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_CT_GUID_Label As clsCTGUIDLabel) As Boolean
    Try
      Dim key As String = objWMS_CT_GUID_Label._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _GUID = objWMS_CT_GUID_Label.GUID
      _UniqueID = objWMS_CT_GUID_Label.UniqueID
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
