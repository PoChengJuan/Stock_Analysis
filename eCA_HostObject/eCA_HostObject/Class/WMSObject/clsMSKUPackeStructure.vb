Public Class clsMSKUPackeStructure
  Private ShareName As String = "MSKUPackeStructure"
  Private ShareKey As String = ""
  Private _gid As String
  Private _SKU_NO As String '貨品編號

  Private _PACKE_LV As Long '包裝層級(最小為1，一直往上加)

  Private _PACKE_UNIT As String '包裝單位

  Private _SUB_PACKE_UNIT As String '子包裝單位

  Private _PACKE_WEIGHT As Double '包裝重量

  Private _PACKE_VOLUME As Double '包裝體積

  Private _PACKE_BCR As String '包裝條碼

  Private _OUT_MAX_UNIT As Long '出庫最大單位

  Private _IN_MAX_UNIT As Long '入庫最大單位

  Private _QTY As Double '包裝數量

  Private _COMMENTS As String '備註

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
  Public Property SKU_NO() As String
    Get
      Return _SKU_NO
    End Get
    Set(ByVal value As String)
      _SKU_NO = value
    End Set
  End Property
  Public Property PACKE_LV() As Long
    Get
      Return _PACKE_LV
    End Get
    Set(ByVal value As Long)
      _PACKE_LV = value
    End Set
  End Property
  Public Property PACKE_UNIT() As String
    Get
      Return _PACKE_UNIT
    End Get
    Set(ByVal value As String)
      _PACKE_UNIT = value
    End Set
  End Property
  Public Property SUB_PACKE_UNIT() As String
    Get
      Return _SUB_PACKE_UNIT
    End Get
    Set(ByVal value As String)
      _SUB_PACKE_UNIT = value
    End Set
  End Property
  Public Property PACKE_WEIGHT() As Double
    Get
      Return _PACKE_WEIGHT
    End Get
    Set(ByVal value As Double)
      _PACKE_WEIGHT = value
    End Set
  End Property
  Public Property PACKE_VOLUME() As Double
    Get
      Return _PACKE_VOLUME
    End Get
    Set(ByVal value As Double)
      _PACKE_VOLUME = value
    End Set
  End Property
  Public Property PACKE_BCR() As String
    Get
      Return _PACKE_BCR
    End Get
    Set(ByVal value As String)
      _PACKE_BCR = value
    End Set
  End Property
  Public Property OUT_MAX_UNIT() As Long
    Get
      Return _OUT_MAX_UNIT
    End Get
    Set(ByVal value As Long)
      _OUT_MAX_UNIT = value
    End Set
  End Property
  Public Property IN_MAX_UNIT() As Long
    Get
      Return _IN_MAX_UNIT
    End Get
    Set(ByVal value As Long)
      _IN_MAX_UNIT = value
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

  Public Sub New(ByVal SKU_NO As String, ByVal PACKE_LV As Long, ByVal PACKE_UNIT As String, ByVal SUB_PACKE_UNIT As String, ByVal PACKE_WEIGHT As Double, ByVal PACKE_VOLUME As Double, ByVal PACKE_BCR As String, ByVal OUT_MAX_UNIT As Long, ByVal IN_MAX_UNIT As Long, ByVal QTY As Double, ByVal COMMENTS As String, ByVal CREATE_TIME As String, ByVal UPDATE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(SKU_NO, PACKE_LV, PACKE_UNIT, SUB_PACKE_UNIT)
      _gid = key
      _SKU_NO = SKU_NO
      _PACKE_LV = PACKE_LV
      _PACKE_UNIT = PACKE_UNIT
      _SUB_PACKE_UNIT = SUB_PACKE_UNIT
      _PACKE_WEIGHT = PACKE_WEIGHT
      _PACKE_VOLUME = PACKE_VOLUME
      _PACKE_BCR = PACKE_BCR
      _OUT_MAX_UNIT = OUT_MAX_UNIT
      _IN_MAX_UNIT = IN_MAX_UNIT
      _QTY = QTY
      _COMMENTS = COMMENTS
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
  Public Shared Function Get_Combination_Key(ByVal SKU_NO As String, ByVal PACKE_LV As Long, ByVal PACKE_UNIT As String, ByVal SUB_PACKE_UNIT As String) As String
    Try
      Dim key As String = SKU_NO & LinkKey & PACKE_LV & LinkKey & PACKE_UNIT & LinkKey & SUB_PACKE_UNIT
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMSKUPackeStructure
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
      Dim strSQL As String = WMS_M_SKU_Packe_StructureManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_SKU_Packe_StructureManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_SKU_Packe_StructureManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_M_SKU_Packe_Structure As clsMSKUPackeStructure) As Boolean
    Try
      Dim key As String = objWMS_M_SKU_Packe_Structure._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _SKU_NO = objWMS_M_SKU_Packe_Structure.SKU_NO
      _PACKE_LV = objWMS_M_SKU_Packe_Structure.PACKE_LV
      _PACKE_UNIT = objWMS_M_SKU_Packe_Structure.PACKE_UNIT
      _SUB_PACKE_UNIT = objWMS_M_SKU_Packe_Structure.SUB_PACKE_UNIT
      _PACKE_WEIGHT = objWMS_M_SKU_Packe_Structure.PACKE_WEIGHT
      _PACKE_VOLUME = objWMS_M_SKU_Packe_Structure.PACKE_VOLUME
      _PACKE_BCR = objWMS_M_SKU_Packe_Structure.PACKE_BCR
      _OUT_MAX_UNIT = objWMS_M_SKU_Packe_Structure.OUT_MAX_UNIT
      _IN_MAX_UNIT = objWMS_M_SKU_Packe_Structure.IN_MAX_UNIT
      _QTY = objWMS_M_SKU_Packe_Structure.QTY
      _COMMENTS = objWMS_M_SKU_Packe_Structure.COMMENTS
      _CREATE_TIME = objWMS_M_SKU_Packe_Structure.CREATE_TIME
      _UPDATE_TIME = objWMS_M_SKU_Packe_Structure.UPDATE_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
