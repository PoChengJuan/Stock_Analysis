Public Class clsTSTOCKTAKING
  Private ShareName As String = "TSTOCKTAKING"
  Private ShareKey As String = ""
  Private _gid As String
  Private _STOCKTAKING_ID As String '盤點單號

  Private _STOCKTAKING_TYPE1 As Long '盤點單類型Auto(自動)=1Manual(手動)=2

  Private _STOCKTAKING_TYPE2 As Long '貨品類型不區分=0原料=1半成品=2成品=3

  Private _STOCKTAKING_TYPE3 As String '單據來源1.ERP單據2.WMS單據

  Private _CREATE_TIME As String '建立時間

  Private _START_TIME As String '執行時間

  Private _FINISH_TIME As String '完成時間

  Private _CREATE_USER As String '建立者

  Private _STATUS As Long '盤點單狀態Queued=0Process=1Normal End=2Abort End=3Cancel End=4

  Private _LOCATION_GROUP_NO As String '位置群組編號

  Private _PRIORITY As Long '單據的優先順序(1~99)

  Private _CARRIER_QTY As Long '需盤點棧板數

  Private _CARRIER_QTY_CHECKED As Long '已盤點棧板數

  Private _MATCH_TYPE As Long '盤點結果Unknow=0(未知) Match=1(符合) Mismatch=2(不符合)

  Private _SEND_TO_HOST As Long '是否上傳1:上傳0:不上傳

  Private _CHANGE_INVENTORY As Long '是否改庫存1:更改庫存0:不更改庫存

  Private _UPLOAD_STATUS As Long '上傳狀態(建立在上傳的時候才有用)  None = 0(預設)HostReportSuccess = 32     '過帳成功  HostReportFailed = 33      '過帳失敗

  Private _UPLOAD_COMMENTS As String '上傳狀態為過帳失敗時註明失敗原因

  Private _objWMS As clsHandlingObject
  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property STOCKTAKING_ID() As String
    Get
      Return _STOCKTAKING_ID
    End Get
    Set(ByVal value As String)
      _STOCKTAKING_ID = value
    End Set
  End Property
  Public Property STOCKTAKING_TYPE1() As Long
    Get
      Return _STOCKTAKING_TYPE1
    End Get
    Set(ByVal value As Long)
      _STOCKTAKING_TYPE1 = value
    End Set
  End Property
  Public Property STOCKTAKING_TYPE2() As Long
    Get
      Return _STOCKTAKING_TYPE2
    End Get
    Set(ByVal value As Long)
      _STOCKTAKING_TYPE2 = value
    End Set
  End Property
  Public Property STOCKTAKING_TYPE3() As String
    Get
      Return _STOCKTAKING_TYPE3
    End Get
    Set(ByVal value As String)
      _STOCKTAKING_TYPE3 = value
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
  Public Property START_TIME() As String
    Get
      Return _START_TIME
    End Get
    Set(ByVal value As String)
      _START_TIME = value
    End Set
  End Property
  Public Property FINISH_TIME() As String
    Get
      Return _FINISH_TIME
    End Get
    Set(ByVal value As String)
      _FINISH_TIME = value
    End Set
  End Property
  Public Property CREATE_USER() As String
    Get
      Return _CREATE_USER
    End Get
    Set(ByVal value As String)
      _CREATE_USER = value
    End Set
  End Property
  Public Property STATUS() As Long
    Get
      Return _STATUS
    End Get
    Set(ByVal value As Long)
      _STATUS = value
    End Set
  End Property
  Public Property LOCATION_GROUP_NO() As String
    Get
      Return _LOCATION_GROUP_NO
    End Get
    Set(ByVal value As String)
      _LOCATION_GROUP_NO = value
    End Set
  End Property
  Public Property PRIORITY() As Long
    Get
      Return _PRIORITY
    End Get
    Set(ByVal value As Long)
      _PRIORITY = value
    End Set
  End Property
  Public Property CARRIER_QTY() As Long
    Get
      Return _CARRIER_QTY
    End Get
    Set(ByVal value As Long)
      _CARRIER_QTY = value
    End Set
  End Property
  Public Property CARRIER_QTY_CHECKED() As Long
    Get
      Return _CARRIER_QTY_CHECKED
    End Get
    Set(ByVal value As Long)
      _CARRIER_QTY_CHECKED = value
    End Set
  End Property
  Public Property MATCH_TYPE() As Long
    Get
      Return _MATCH_TYPE
    End Get
    Set(ByVal value As Long)
      _MATCH_TYPE = value
    End Set
  End Property
  Public Property SEND_TO_HOST() As Long
    Get
      Return _SEND_TO_HOST
    End Get
    Set(ByVal value As Long)
      _SEND_TO_HOST = value
    End Set
  End Property
  Public Property CHANGE_INVENTORY() As Long
    Get
      Return _CHANGE_INVENTORY
    End Get
    Set(ByVal value As Long)
      _CHANGE_INVENTORY = value
    End Set
  End Property
  Public Property UPLOAD_STATUS() As Long
    Get
      Return _UPLOAD_STATUS
    End Get
    Set(ByVal value As Long)
      _UPLOAD_STATUS = value
    End Set
  End Property
  Public Property UPLOAD_COMMENTS() As String
    Get
      Return _UPLOAD_COMMENTS
    End Get
    Set(ByVal value As String)
      _UPLOAD_COMMENTS = value
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

  Public Sub New(ByVal STOCKTAKING_ID As String, ByVal STOCKTAKING_TYPE1 As Long, ByVal STOCKTAKING_TYPE2 As Long, ByVal STOCKTAKING_TYPE3 As String, ByVal CREATE_TIME As String, ByVal START_TIME As String, ByVal FINISH_TIME As String, ByVal CREATE_USER As String, ByVal STATUS As Long, ByVal LOCATION_GROUP_NO As String, ByVal PRIORITY As Long, ByVal CARRIER_QTY As Long, ByVal CARRIER_QTY_CHECKED As Long, ByVal MATCH_TYPE As Long, ByVal SEND_TO_HOST As Long, ByVal CHANGE_INVENTORY As Long, ByVal UPLOAD_STATUS As Long, ByVal UPLOAD_COMMENTS As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(STOCKTAKING_ID)
      _gid = key
      _STOCKTAKING_ID = STOCKTAKING_ID
      _STOCKTAKING_TYPE1 = STOCKTAKING_TYPE1
      _STOCKTAKING_TYPE2 = STOCKTAKING_TYPE2
      _STOCKTAKING_TYPE3 = STOCKTAKING_TYPE3
      _CREATE_TIME = CREATE_TIME
      _START_TIME = START_TIME
      _FINISH_TIME = FINISH_TIME
      _CREATE_USER = CREATE_USER
      _STATUS = STATUS
      _LOCATION_GROUP_NO = LOCATION_GROUP_NO
      _PRIORITY = PRIORITY
      _CARRIER_QTY = CARRIER_QTY
      _CARRIER_QTY_CHECKED = CARRIER_QTY_CHECKED
      _MATCH_TYPE = MATCH_TYPE
      _SEND_TO_HOST = SEND_TO_HOST
      _CHANGE_INVENTORY = CHANGE_INVENTORY
      _UPLOAD_STATUS = UPLOAD_STATUS
      _UPLOAD_COMMENTS = UPLOAD_COMMENTS
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
  Public Shared Function Get_Combination_Key(ByVal STOCKTAKING_ID As String) As String
    Try
      Dim key As String = STOCKTAKING_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsTSTOCKTAKING
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  ' Public Sub Add_Relationship(ByRef objWMS As clsWMSObject)                                 
  '   Try                                                                                     
  '     '挷定WMS的關係                                                                        
  '     If objWMS IsNot Nothing Then                                                          
  '       _objWMS = objWMS                                                                    
  '       objWMS.O_Add_!!!!!這邊就是你要改的東西啦(Me)
  '     End If                                                                                
  '   Catch ex As Exception                                                                   
  '     SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)                
  '   End Try                                                                                 
  ' End Sub                                                                                   
  'Public Sub Remove_Relationship()                                          
  '  Try                                                                                    
  '    If _objWMS IsNot Nothing Then                                                                
  '      _objWMS.O_Remove_ !!!!!這也是你要改的東西        
  '    End If                                                                                       
  '  Catch ex As Exception                                                                          
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)                       
  '  End Try                                                                                        
  'End Sub                                                                                          
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_STOCKTAKINGManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_STOCKTAKINGManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_STOCKTAKINGManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_T_STOCKTAKING As clsTSTOCKTAKING) As Boolean
    Try
      Dim key As String = objWMS_T_STOCKTAKING._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _STOCKTAKING_ID = objWMS_T_STOCKTAKING.STOCKTAKING_ID
      _STOCKTAKING_TYPE1 = objWMS_T_STOCKTAKING.STOCKTAKING_TYPE1
      _STOCKTAKING_TYPE2 = objWMS_T_STOCKTAKING.STOCKTAKING_TYPE2
      _STOCKTAKING_TYPE3 = objWMS_T_STOCKTAKING.STOCKTAKING_TYPE3
      _CREATE_TIME = objWMS_T_STOCKTAKING.CREATE_TIME
      _START_TIME = objWMS_T_STOCKTAKING.START_TIME
      _FINISH_TIME = objWMS_T_STOCKTAKING.FINISH_TIME
      _CREATE_USER = objWMS_T_STOCKTAKING.CREATE_USER
      _STATUS = objWMS_T_STOCKTAKING.STATUS
      _LOCATION_GROUP_NO = objWMS_T_STOCKTAKING.LOCATION_GROUP_NO
      _PRIORITY = objWMS_T_STOCKTAKING.PRIORITY
      _CARRIER_QTY = objWMS_T_STOCKTAKING.CARRIER_QTY
      _CARRIER_QTY_CHECKED = objWMS_T_STOCKTAKING.CARRIER_QTY_CHECKED
      _MATCH_TYPE = objWMS_T_STOCKTAKING.MATCH_TYPE
      _SEND_TO_HOST = objWMS_T_STOCKTAKING.SEND_TO_HOST
      _CHANGE_INVENTORY = objWMS_T_STOCKTAKING.CHANGE_INVENTORY
      _UPLOAD_STATUS = objWMS_T_STOCKTAKING.UPLOAD_STATUS
      _UPLOAD_COMMENTS = objWMS_T_STOCKTAKING.UPLOAD_COMMENTS
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
