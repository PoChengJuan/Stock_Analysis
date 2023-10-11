Public Class clsSL
  Private ShareName As String = "MSL"
  Private ShareKey As String = ""
  Private _gid As String
  Private _OWNER_NO As String '貨主編號

  Private _SL_NO As String 'ERP儲存地點

  Private _SL_ID As String

  Private _SL_ALIS As String

  Private _SL_DESC As String

  Private _BND As Long '保稅0:不保稅1:保稅

  Private _QC_STATUS As Long 'QC判定狀態OK	'良品=1NG	'不良品=2

  Private _REPORT_TO_HOST As Long '是否上報給上位系統0:不上報1:要上報

  Private _CREATE_TIME As String '建立時間

  Private _UPDATE_TIME As String '更新時間

  Private _objWMS As clsHandlingObject
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
  Public Property SL_NO() As String
    Get
      Return _SL_NO
    End Get
    Set(ByVal value As String)
      _SL_NO = value
    End Set
  End Property
  Public Property SL_ID() As String
    Get
      Return _SL_ID
    End Get
    Set(ByVal value As String)
      _SL_ID = value
    End Set
  End Property
  Public Property SL_ALIS() As String
    Get
      Return _SL_ALIS
    End Get
    Set(ByVal value As String)
      _SL_ALIS = value
    End Set
  End Property
  Public Property SL_DESC() As String
    Get
      Return _SL_DESC
    End Get
    Set(ByVal value As String)
      _SL_DESC = value
    End Set
  End Property
  Public Property BND() As Long
    Get
      Return _BND
    End Get
    Set(ByVal value As Long)
      _BND = value
    End Set
  End Property
  Public Property QC_STATUS() As Long
    Get
      Return _QC_STATUS
    End Get
    Set(ByVal value As Long)
      _QC_STATUS = value
    End Set
  End Property
  Public Property REPORT_TO_HOST() As Long
    Get
      Return _REPORT_TO_HOST
    End Get
    Set(ByVal value As Long)
      _REPORT_TO_HOST = value
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
  Public Property objWMS() As clsHandlingObject
    Get
      Return _objWMS
    End Get
    Set(ByVal value As clsHandlingObject)
      _objWMS = value
    End Set
  End Property

  Public Sub New(ByVal OWNER_NO As String, ByVal SL_NO As String, ByVal SL_ID As String, ByVal SL_ALIS As String, ByVal SL_DESC As String, ByVal BND As Long, ByVal QC_STATUS As Long, ByVal REPORT_TO_HOST As Long, ByVal CREATE_TIME As String, ByVal UPDATE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(OWNER_NO, SL_NO)
      _gid = key
      _OWNER_NO = OWNER_NO
      _SL_NO = SL_NO
      _SL_ID = SL_ID
      _SL_ALIS = SL_ALIS
      _SL_DESC = SL_DESC
      _BND = BND
      _QC_STATUS = QC_STATUS
      _REPORT_TO_HOST = REPORT_TO_HOST
      _CREATE_TIME = CREATE_TIME
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
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal OWNER_NO As String, ByVal SL_NO As String) As String
    Try
      Dim key As String = OWNER_NO & LinkKey & SL_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsSL
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
      Dim strSQL As String = WMS_M_SLManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_SLManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_SLManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_M_SL As clsSL) As Boolean
    Try
      Dim key As String = objWMS_M_SL._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _OWNER_NO = objWMS_M_SL.OWNER_NO
      _SL_NO = objWMS_M_SL.SL_NO
      _SL_ID = objWMS_M_SL.SL_ID
      _SL_ALIS = objWMS_M_SL.SL_ALIS
      _SL_DESC = objWMS_M_SL.SL_DESC
      _BND = objWMS_M_SL.BND
      _QC_STATUS = objWMS_M_SL.QC_STATUS
      _REPORT_TO_HOST = objWMS_M_SL.REPORT_TO_HOST
      _CREATE_TIME = objWMS_M_SL.CREATE_TIME
      _UPDATE_TIME = objWMS_M_SL.UPDATE_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
