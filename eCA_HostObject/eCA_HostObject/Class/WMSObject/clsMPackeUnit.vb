Public Class clsMPackeUnit
  Private ShareName As String = "MPackeUnit"
  Private ShareKey As String = ""
  Private _gid As String
  Private _PACKE_UNIT As String '包裝單位

  Private _PACKE_UNIT_NAME As String '包裝單位名稱

  Private _PACKE_UNIT_COMMON1 As String '通用欄位1

  Private _PACKE_UNIT_COMMON2 As String '通用欄位2

  Private _PACKE_UNIT_COMMON3 As String '通用欄位3

  Private _PACKE_UNIT_COMMON4 As String '通用欄位4

  Private _PACKE_UNIT_COMMON5 As String '通用欄位5

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
  Public Property PACKE_UNIT() As String
    Get
      Return _PACKE_UNIT
    End Get
    Set(ByVal value As String)
      _PACKE_UNIT = value
    End Set
  End Property
  Public Property PACKE_UNIT_NAME() As String
    Get
      Return _PACKE_UNIT_NAME
    End Get
    Set(ByVal value As String)
      _PACKE_UNIT_NAME = value
    End Set
  End Property
  Public Property PACKE_UNIT_COMMON1() As String
    Get
      Return _PACKE_UNIT_COMMON1
    End Get
    Set(ByVal value As String)
      _PACKE_UNIT_COMMON1 = value
    End Set
  End Property
  Public Property PACKE_UNIT_COMMON2() As String
    Get
      Return _PACKE_UNIT_COMMON2
    End Get
    Set(ByVal value As String)
      _PACKE_UNIT_COMMON2 = value
    End Set
  End Property
  Public Property PACKE_UNIT_COMMON3() As String
    Get
      Return _PACKE_UNIT_COMMON3
    End Get
    Set(ByVal value As String)
      _PACKE_UNIT_COMMON3 = value
    End Set
  End Property
  Public Property PACKE_UNIT_COMMON4() As String
    Get
      Return _PACKE_UNIT_COMMON4
    End Get
    Set(ByVal value As String)
      _PACKE_UNIT_COMMON4 = value
    End Set
  End Property
  Public Property PACKE_UNIT_COMMON5() As String
    Get
      Return _PACKE_UNIT_COMMON5
    End Get
    Set(ByVal value As String)
      _PACKE_UNIT_COMMON5 = value
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

  Public Sub New(ByVal PACKE_UNIT As String, ByVal PACKE_UNIT_NAME As String, ByVal PACKE_UNIT_COMMON1 As String, ByVal PACKE_UNIT_COMMON2 As String, ByVal PACKE_UNIT_COMMON3 As String, ByVal PACKE_UNIT_COMMON4 As String, ByVal PACKE_UNIT_COMMON5 As String, ByVal COMMENTS As String, ByVal CREATE_TIME As String, ByVal UPDATE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PACKE_UNIT)
      _gid = key
      _PACKE_UNIT = PACKE_UNIT
      _PACKE_UNIT_NAME = PACKE_UNIT_NAME
      _PACKE_UNIT_COMMON1 = PACKE_UNIT_COMMON1
      _PACKE_UNIT_COMMON2 = PACKE_UNIT_COMMON2
      _PACKE_UNIT_COMMON3 = PACKE_UNIT_COMMON3
      _PACKE_UNIT_COMMON4 = PACKE_UNIT_COMMON4
      _PACKE_UNIT_COMMON5 = PACKE_UNIT_COMMON5
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
  Public Shared Function Get_Combination_Key(ByVal PACKE_UNIT As String) As String
    Try
      Dim key As String = PACKE_UNIT
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsMPackeUnit
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
      Dim strSQL As String = WMS_M_Packe_UnitManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_Packe_UnitManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_Packe_UnitManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_M_Packe_Unit As clsMPackeUnit) As Boolean
    Try
      Dim key As String = objWMS_M_Packe_Unit._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _PACKE_UNIT = objWMS_M_Packe_Unit.PACKE_UNIT
      _PACKE_UNIT_NAME = objWMS_M_Packe_Unit.PACKE_UNIT_NAME
      _PACKE_UNIT_COMMON1 = objWMS_M_Packe_Unit.PACKE_UNIT_COMMON1
      _PACKE_UNIT_COMMON2 = objWMS_M_Packe_Unit.PACKE_UNIT_COMMON2
      _PACKE_UNIT_COMMON3 = objWMS_M_Packe_Unit.PACKE_UNIT_COMMON3
      _PACKE_UNIT_COMMON4 = objWMS_M_Packe_Unit.PACKE_UNIT_COMMON4
      _PACKE_UNIT_COMMON5 = objWMS_M_Packe_Unit.PACKE_UNIT_COMMON5
      _COMMENTS = objWMS_M_Packe_Unit.COMMENTS
      _CREATE_TIME = objWMS_M_Packe_Unit.CREATE_TIME
      _UPDATE_TIME = objWMS_M_Packe_Unit.UPDATE_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
