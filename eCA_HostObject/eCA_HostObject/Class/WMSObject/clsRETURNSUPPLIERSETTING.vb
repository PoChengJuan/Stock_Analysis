Public Class clsRETURNSUPPLIERSETTING
  Private ShareName As String = "MRETURNSUPPLIERSETTING"
  Private ShareKey As String = ""
  Private _gid As String
  Private _LOCATION_NO As String '位置編號

  Private _SUPPLIER_NO As String '供應商編號

  Private _HIGH_WATER As Long '高水位

  Private _LOW_WATER As Long '低水位

  Private _objWMS As clsHandlingObject
  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property LOCATION_NO() As String
    Get
      Return _LOCATION_NO
    End Get
    Set(ByVal value As String)
      _LOCATION_NO = value
    End Set
  End Property
  Public Property SUPPLIER_NO() As String
    Get
      Return _SUPPLIER_NO
    End Get
    Set(ByVal value As String)
      _SUPPLIER_NO = value
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
  Public Property objWMS() As clsHandlingObject
    Get
      Return _objWMS
    End Get
    Set(ByVal value As clsHandlingObject)
      _objWMS = value
    End Set
  End Property

  Public Sub New(ByVal LOCATION_NO As String, ByVal SUPPLIER_NO As String, ByVal HIGH_WATER As Long, ByVal LOW_WATER As Long)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(LOCATION_NO)
      _gid = key
      _LOCATION_NO = LOCATION_NO
      _SUPPLIER_NO = SUPPLIER_NO
      _HIGH_WATER = HIGH_WATER
      _LOW_WATER = LOW_WATER
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
  Public Shared Function Get_Combination_Key(ByVal LOCATION_NO As String) As String
    Try
      Dim key As String = LOCATION_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsRETURNSUPPLIERSETTING
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
      Dim strSQL As String = WMS_M_RETURN_SUPPLIER_SETTINGManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_RETURN_SUPPLIER_SETTINGManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_RETURN_SUPPLIER_SETTINGManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_M_RETURN_SUPPLIER_SETTING As clsRETURNSUPPLIERSETTING) As Boolean
    Try
      Dim key As String = objWMS_M_RETURN_SUPPLIER_SETTING._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _LOCATION_NO = objWMS_M_RETURN_SUPPLIER_SETTING.LOCATION_NO
      _SUPPLIER_NO = objWMS_M_RETURN_SUPPLIER_SETTING.SUPPLIER_NO
      _HIGH_WATER = objWMS_M_RETURN_SUPPLIER_SETTING.HIGH_WATER
      _LOW_WATER = objWMS_M_RETURN_SUPPLIER_SETTING.LOW_WATER
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
