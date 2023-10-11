Public Class clsLine_Area
  Private ShareName As String = "Line_Area"
  Private ShareKey As String = ""

  Private _gid As String
  Private _Factory_No As String
  Private _Area_No As String
  Private _Area_ID As String
  Private _Area_Alis As String
  Private _Area_Desc As String
  Private _Area_Type2 As enuAreaType2
  Private _High_Water As Long
  Private _Low_Water As Long
  Private _Device_No As String
  Private _Enable As Boolean
  Private _Show_Index As Long
  Private _Show_Group As Long
  Private _Show_Color As String

  Private _Process_ID As String
  Private _Process_CODE As String
  Private _TB004 As String
  Private _TB005 As String
  Private _TB007 As String
  Private _TB008 As String
  Private _TB010 As String

  Private _Previous_Area_No As String
  Private _Area_Index As Double
  Private _Area_Type1 As Double
  Private _Report As Boolean

  Private _objHandling As clsHandlingObject
  '1.Line
  Public gdicLine As New Concurrent.ConcurrentDictionary(Of String, clsLine_Status)
  '2.LineProduction_Info
  Public gdicLineProduction_Info As New Concurrent.ConcurrentDictionary(Of String, clsLineProduction_Info)

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property Factory_No() As String
    Get
      Return _Factory_No
    End Get
    Set(ByVal value As String)
      _Factory_No = value
    End Set
  End Property
  Public Property Area_No() As String
    Get
      Return _Area_No
    End Get
    Set(ByVal value As String)
      _Area_No = value
    End Set
  End Property
  Public Property Area_ID() As String
    Get
      Return _Area_ID
    End Get
    Set(ByVal value As String)
      _Area_ID = value
    End Set
  End Property
  Public Property Area_Alis() As String
    Get
      Return _Area_Alis
    End Get
    Set(ByVal value As String)
      _Area_Alis = value
    End Set
  End Property
  Public Property Area_Desc() As String
    Get
      Return _Area_Desc
    End Get
    Set(ByVal value As String)
      _Area_Desc = value
    End Set
  End Property
  Public Property Area_Type2() As enuAreaType2
    Get
      Return _Area_Type2
    End Get
    Set(ByVal value As enuAreaType2)
      _Area_Type2 = value
    End Set
  End Property
  Public Property High_Water() As Long
    Get
      Return _High_Water
    End Get
    Set(ByVal value As Long)
      _High_Water = value
    End Set
  End Property
  Public Property Low_Water() As Long
    Get
      Return _Low_Water
    End Get
    Set(ByVal value As Long)
      _Low_Water = value
    End Set
  End Property
  Public Property Enable() As Boolean
    Get
      Return _Enable
    End Get
    Set(ByVal value As Boolean)
      _Enable = value
    End Set
  End Property
  Public Property Device_No() As String
    Get
      Return _Device_No
    End Get
    Set(ByVal value As String)
      _Device_No = value
    End Set
  End Property
  Public Property Show_Index() As Long
    Get
      Return _Show_Index
    End Get
    Set(ByVal value As Long)
      _Show_Index = value
    End Set
  End Property
  Public Property Show_Group() As Long
    Get
      Return _Show_Group
    End Get
    Set(ByVal value As Long)
      _Show_Group = value
    End Set
  End Property
  Public Property Show_Color() As String
    Get
      Return _Show_Color
    End Get
    Set(ByVal value As String)
      _Show_Color = value
    End Set
  End Property
  Public Property Process_ID() As String
    Get
      Return _Process_ID
    End Get
    Set(ByVal value As String)
      _Process_ID = value
    End Set
  End Property
  Public Property Process_CODE() As String
    Get
      Return _Process_CODE
    End Get
    Set(ByVal value As String)
      _Process_CODE = value
    End Set
  End Property
  Public Property TB004() As String
    Get
      Return _TB004
    End Get
    Set(ByVal value As String)
      _TB004 = value
    End Set
  End Property
  Public Property TB005() As String
    Get
      Return _TB005
    End Get
    Set(ByVal value As String)
      _TB005 = value
    End Set
  End Property
  Public Property TB007() As String
    Get
      Return _TB007
    End Get
    Set(ByVal value As String)
      _TB007 = value
    End Set
  End Property
  Public Property TB008() As String
    Get
      Return _TB008
    End Get
    Set(ByVal value As String)
      _TB008 = value
    End Set
  End Property
  Public Property TB010() As String
    Get
      Return _TB010
    End Get
    Set(ByVal value As String)
      _TB010 = value
    End Set
  End Property

  Public Property PREVIOUS_AREA_NO() As String
    Get
      Return _Previous_Area_No
    End Get
    Set(ByVal value As String)
      _Previous_Area_No = value
    End Set
  End Property
  Public Property AREA_INDEX() As Long
    Get
      Return _Area_Index
    End Get
    Set(ByVal value As Long)
      _Area_Index = value
    End Set
  End Property
  Public Property AREA_TYPE1() As enuAreaType1
    Get
      Return _Area_Type1
    End Get
    Set(ByVal value As enuAreaType1)
      _Area_Type1 = value
    End Set
  End Property
  Public Property Report() As Boolean
    Get
      Return _Report
    End Get
    Set(ByVal value As Boolean)
      _Report = value
    End Set
  End Property



  '物件建立時執行的事件
  Public Sub New(ByVal Factory_No As String,
                               ByVal Area_No As String,
                               ByVal Area_ID As String,
                               ByVal Area_Alis As String,
                               ByVal Area_Desc As String,
                               ByVal Area_Type2 As Long,
                               ByVal High_Water As Long,
                               ByVal Low_Water As Long,
                               ByVal Device_No As String,
                               ByVal Enable As Boolean,
                               ByVal Show_Index As Long,
                               ByVal Show_Group As Long,
                               ByVal Show_Color As String,
                               ByVal Process_ID As String,
                               ByVal Process_CODE As String,
                               ByVal TB004 As String,
                               ByVal TB005 As String,
                               ByVal TB007 As String,
                               ByVal TB008 As String,
                               ByVal TB010 As String,
                               ByVal PREVIOUS_AREA_NO As String,
                               ByVal AREA_INDEX As Long,
                               ByVal AREA_TYPE1 As Long,
                               ByVal Report As Boolean)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(Factory_No, Area_No)
      _gid = key
      _Factory_No = Factory_No
      _Area_No = Area_No
      _Area_ID = Area_ID
      _Area_Alis = Area_Alis
      _Area_Desc = Area_Desc
      _Area_Type2 = Area_Type2
      _High_Water = High_Water
      _Low_Water = Low_Water
      _Device_No = Device_No
      _Enable = Enable
      _Show_Index = Show_Index
      _Show_Group = Show_Group
      _Show_Color = Show_Color
      _Process_ID = Process_ID
      _Process_CODE = Process_CODE
      _TB004 = TB004
      _TB005 = TB005
      _TB007 = TB007
      _TB008 = TB008
      _TB010 = TB010

      _Previous_Area_No = PREVIOUS_AREA_NO
      _Area_Index = AREA_INDEX
      _Area_Type1 = AREA_TYPE1
      _Report = Report



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
  Public Shared Function Get_Combination_Key(ByVal Factory_No As String, ByVal Area_No As String) As String
    Try
      Dim key As String = Factory_No & LinkKey & Area_No
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsLine_Area
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
        objHandling.O_Add_Line_Area(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_Line_Area(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_CM_Line_AreaManagement.GetUpdateSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '把Line加入gcolLine
  Public Function O_Add_Line(ByRef obj As clsLine_Status) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicLine.ContainsKey(key) Then
        gdicLine.TryAdd(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '資料從Dictionary刪除
  '從gcolCLine刪除objCLine
  Public Function O_Remove_Line(ByRef obj As clsLine_Status) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicLine.ContainsKey(key) Then
        gdicLine.TryRemove(key, obj)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得Dictionary內的資料
  '從gcolLine取得指定的objLine
  Public Function O_Get_Line(ByVal Factory_No As String, ByVal Area_No As String,
                                                       ByVal Device_No As String, ByVal Unt_ID As String,
                                                       Optional ByRef RetObj As clsLine_Status = Nothing) As Boolean
    Try
      Dim key As String = clsLine_Status.Get_Combination_Key(Factory_No, Area_No, Device_No, Unt_ID)
      Dim obj As clsLine_Status
      If gdicLine.ContainsKey(key) Then
        obj = gdicLine.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '把LineProduction_Info加入gcolLineProduction_Info
  Public Function O_Add_LineProduction_Info(ByRef obj As clsLineProduction_Info) As Boolean
    Try
      Dim key As String = obj.gid()
      If Not gdicLineProduction_Info.ContainsKey(key) Then
        gdicLineProduction_Info.TryAdd(key, obj)
      Else
        SendMessageToLog("Add Dictionary Failed, key already exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '資料從Dictionary刪除
  '從gcolCLineProduction_Info刪除objCLineProduction_Info
  Public Function O_Remove_LineProduction_Info(ByRef obj As clsLineProduction_Info) As Boolean
    Try
      Dim key As String = obj.gid
      If gdicLineProduction_Info.ContainsKey(key) Then
        gdicLineProduction_Info.TryRemove(key, obj)
      Else
        SendMessageToLog("Reomve Dictionary Failed, key not exists key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得Dictionary內的資料
  '從gcolLineProduction_Info取得指定的objLineProduction_Info
  Public Function O_Get_LineProduction_Info(ByVal Factory_No As String, ByVal Area_No As String,
                                                       ByVal Device_No As String, ByVal Unt_ID As String,
                                                       Optional ByRef RetObj As clsLineProduction_Info = Nothing) As Boolean
    Try
      Dim key As String = clsLineProduction_Info.Get_Combination_Key(Factory_No, Area_No, Device_No, Unt_ID)
      Dim obj As clsLineProduction_Info
      If gdicLineProduction_Info.ContainsKey(key) Then
        obj = gdicLineProduction_Info.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsLine_Area) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> gid Then
        SendMessageToLog("Key can not Update, old_Key=" & gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _Factory_No = obj.Factory_No
      _Area_No = obj.Area_No
      _Area_ID = obj.Area_ID
      _Area_Alis = obj.Area_Alis
      _Area_Desc = obj.Area_Desc
      _Area_Type2 = obj.Area_Type2
      _High_Water = obj.High_Water
      _Low_Water = obj.Low_Water
      _Device_No = obj.Device_No
      _Enable = obj.Enable
      _Show_Index = obj.Show_Index
      _Show_Group = obj._Show_Group
      _Show_Color = obj.Show_Color

      _Process_ID = obj.Process_ID
      _Process_CODE = obj.Process_CODE
      _TB004 = obj.TB004
      _TB005 = obj.TB005
      _TB007 = obj.TB007
      _TB008 = obj.TB008
      _TB010 = obj.TB010
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
