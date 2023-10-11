Public Class clsCarrier
  Private ShareName As String = "Carrier"
  Private ShareKey As String = ""

  Private _gid As String
  Private _Carrier_ID As String
  Private _Carrier_Alis As String
  Private _Carrier_Desc As String
  Private _Carrier_Type As Long
  Private _Carrier_Mode As enuCarrierMode
  Private _Create_Time As String
  Private _Comments As String
  Private _SubLocation_Index_X As Long
  Private _SubLocation_Index_Y As Long
  Private _SubLocation_Index_Z As Long
  Private _WEIGHT As Long
  Private _LENGTH As Long
  Private _WIDTH As Long
  Private _HEIGHT As Long
  Private _CARRIER_COMMON01 As String
  Private _CARRIER_COMMON02 As String
  Private _CARRIER_COMMON03 As String
  Private _CARRIER_COMMON04 As String
  Private _CARRIER_COMMON05 As String
  Private _CARRIER_COMMON06 As String
  Private _CARRIER_COMMON07 As String
  Private _CARRIER_COMMON08 As String
  Private _CARRIER_COMMON09 As String
  Private _CARRIER_COMMON10 As String
  Private _CREATE_USER_ID As String
  Private _Location_No As String
  Private _Reserved As Boolean
  Private _Locked As Boolean
  Private _Locked_User As String
  Private _Locked_Reason As String
  Private _Locked_Time As String
  Private _Stage_ID As String
  Private _Return_Location_No As String
  Private _Stocktaking_Time As String
  Private _Firstin_Time As String
  Private _Last_Transfer_Time As String
  Private _Update_Time As String
  Private _SubLocation_X As Long
  Private _SubLocation_Y As Long
  Private _SubLocation_Z As Long
  Private _Update_User_ID As String
  Private _UNPACK_TIME As String
  Private _TALLY_ENABLE As Boolean '是否需要進行理貨0:不需要理貨1:需要理貨



  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property Carrier_ID() As String
    Get
      Return _Carrier_ID
    End Get
    Set(ByVal value As String)
      _Carrier_ID = value
    End Set
  End Property
  Public Property Carrier_Alis() As String
    Get
      Return _Carrier_Alis
    End Get
    Set(ByVal value As String)
      _Carrier_Alis = value
    End Set
  End Property
  Public Property Carrier_Desc() As String
    Get
      Return _Carrier_Desc
    End Get
    Set(ByVal value As String)
      _Carrier_Desc = value
    End Set
  End Property
  Public Property Carrier_Type() As Long
    Get
      Return _Carrier_Type
    End Get
    Set(ByVal value As Long)
      _Carrier_Type = value
    End Set
  End Property
  Public Property Carrier_Mode() As enuCarrierMode
    Get
      Return _Carrier_Mode
    End Get
    Set(ByVal value As enuCarrierMode)
      _Carrier_Mode = value
    End Set
  End Property
  Public Property Create_Time() As String
    Get
      Return _Create_Time
    End Get
    Set(ByVal value As String)
      _Create_Time = value
    End Set
  End Property
  Public Property Comments() As String
    Get
      Return _Comments
    End Get
    Set(ByVal value As String)
      _Comments = value
    End Set
  End Property
  Public Property SubLocation_Index_X() As Long
    Get
      Return _SubLocation_Index_X
    End Get
    Set(ByVal value As Long)
      _SubLocation_Index_X = value
    End Set
  End Property
  Public Property SubLocation_Index_Y() As Long
    Get
      Return _SubLocation_Index_Y
    End Get
    Set(ByVal value As Long)
      _SubLocation_Index_Y = value
    End Set
  End Property
  Public Property SubLocation_Index_Z() As Long
    Get
      Return _SubLocation_Index_Z
    End Get
    Set(ByVal value As Long)
      _SubLocation_Index_Z = value
    End Set
  End Property
  Public Property WEIGHT() As Long
    Get
      Return _WEIGHT
    End Get
    Set(ByVal value As Long)
      _WEIGHT = value
    End Set
  End Property
  Public Property LENGTH() As Long
    Get
      Return _LENGTH
    End Get
    Set(ByVal value As Long)
      _LENGTH = value
    End Set
  End Property
  Public Property WIDTH() As Long
    Get
      Return _WIDTH
    End Get
    Set(ByVal value As Long)
      _WIDTH = value
    End Set
  End Property
  Public Property HEIGHT() As Long
    Get
      Return _HEIGHT
    End Get
    Set(ByVal value As Long)
      _HEIGHT = value
    End Set
  End Property
  Public Property CARRIER_COMMON01() As String
    Get
      Return _CARRIER_COMMON01
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON01 = value
    End Set
  End Property
  Public Property CARRIER_COMMON02() As String
    Get
      Return _CARRIER_COMMON02
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON02 = value
    End Set
  End Property
  Public Property CARRIER_COMMON03() As String
    Get
      Return _CARRIER_COMMON03
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON03 = value
    End Set
  End Property
  Public Property CARRIER_COMMON04() As String
    Get
      Return _CARRIER_COMMON04
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON04 = value
    End Set
  End Property
  Public Property CARRIER_COMMON05() As String
    Get
      Return _CARRIER_COMMON05
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON05 = value
    End Set
  End Property
  Public Property CARRIER_COMMON06() As String
    Get
      Return _CARRIER_COMMON06
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON06 = value
    End Set
  End Property
  Public Property CARRIER_COMMON07() As String
    Get
      Return _CARRIER_COMMON07
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON07 = value
    End Set
  End Property
  Public Property CARRIER_COMMON08() As String
    Get
      Return _CARRIER_COMMON08
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON08 = value
    End Set
  End Property
  Public Property CARRIER_COMMON09() As String
    Get
      Return _CARRIER_COMMON09
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON09 = value
    End Set
  End Property
  Public Property CARRIER_COMMON10() As String
    Get
      Return _CARRIER_COMMON10
    End Get
    Set(ByVal value As String)
      _CARRIER_COMMON10 = value
    End Set
  End Property
  Public Property Location_No() As String
    Get
      Return _Location_No
    End Get
    Set(ByVal value As String)
      _Location_No = value
    End Set
  End Property
  Public Property RESERVED() As Boolean
    Get
      Return _Reserved
    End Get
    Set(ByVal value As Boolean)
      _Reserved = value
    End Set
  End Property
  Public Property LOCKED() As Boolean
    Get
      Return _Locked
    End Get
    Set(ByVal value As Boolean)
      _Locked = value
    End Set
  End Property
  Public Property LOCKED_USER() As String
    Get
      Return _Locked_User
    End Get
    Set(ByVal value As String)
      _Locked_User = value
    End Set
  End Property
  Public Property LOCKED_REASON() As String
    Get
      Return _Locked_Reason
    End Get
    Set(ByVal value As String)
      _Locked_Reason = value
    End Set
  End Property
  Public Property LOCKED_TIME() As String
    Get
      Return _Locked_Time
    End Get
    Set(ByVal value As String)
      _Locked_Time = value
    End Set
  End Property
  Public Property STAGE_ID() As String
    Get
      Return _Stage_ID
    End Get
    Set(ByVal value As String)
      _Stage_ID = value
    End Set
  End Property
  Public Property RETURN_LOCATION_NO() As String
    Get
      Return _Return_Location_No
    End Get
    Set(ByVal value As String)
      _Return_Location_No = value
    End Set
  End Property
  Public Property STOCKTAKING_TIME() As String
    Get
      Return _Stocktaking_Time
    End Get
    Set(ByVal value As String)
      _Stocktaking_Time = value
    End Set
  End Property
  Public Property FIRSTIN_TIME() As String
    Get
      Return _Firstin_Time
    End Get
    Set(ByVal value As String)
      _Firstin_Time = value
    End Set
  End Property
  Public Property LAST_TRANSFER_TIME() As String
    Get
      Return _Last_Transfer_Time
    End Get
    Set(ByVal value As String)
      _Last_Transfer_Time = value
    End Set
  End Property
  Public Property UPDATE_TIME() As String
    Get
      Return _Update_Time
    End Get
    Set(ByVal value As String)
      _Update_Time = value
    End Set
  End Property
  Public Property SUBLOCATION_X() As Long
    Get
      Return _SubLocation_X
    End Get
    Set(ByVal value As Long)
      _SubLocation_X = value
    End Set
  End Property
  Public Property SUBLOCATION_Y() As Long
    Get
      Return _SubLocation_Y
    End Get
    Set(ByVal value As Long)
      _SubLocation_Y = value
    End Set
  End Property
  Public Property SUBLOCATION_Z() As Long
    Get
      Return _SubLocation_Z
    End Get
    Set(ByVal value As Long)
      _SubLocation_Z = value
    End Set
  End Property
  Public Property Update_User_ID() As String
    Get
      Return _Update_User_ID
    End Get
    Set(ByVal value As String)
      _Update_User_ID = value
    End Set
  End Property
  Public Property UNPACK_TIME() As String
    Get
      Return _UNPACK_TIME
    End Get
    Set(ByVal value As String)
      _UNPACK_TIME = value
    End Set
  End Property
  Public Property TALLY_ENABLE() As Boolean
    Get
      Return _TALLY_ENABLE
    End Get
    Set(ByVal value As Boolean)
      _TALLY_ENABLE = value
    End Set
  End Property
  Public Property CREATE_USER_ID() As String
    Get
      Return _CREATE_USER_ID
    End Get
    Set(ByVal value As String)
      _CREATE_USER_ID = value
    End Set
  End Property

  '物件建立時執行的事件
  Public Sub New(ByVal Carrier_ID As String, ByVal Carrier_Alis As String, ByVal Carrier_Desc As String,
                 ByVal Carrier_Type As Long, ByVal Carrier_Mode As enuCarrierMode, ByVal Create_Time As String,
                 ByVal Comments As String, ByVal SubLocation_Index_X As Long, ByVal SubLocation_Index_Y As Long, ByVal SubLocation_Index_Z As Long,
                 ByVal Location_No As String, ByVal Reserved As Boolean, ByVal Locked As Boolean,
                 ByVal Locked_User As String, ByVal Locked_Reason As String, ByVal Locked_Time As String, ByVal Stage_ID As String,
                 ByVal Return_Location_No As String, ByVal Stocktaking_Time As String, ByVal Firstin_Time As String,
                 ByVal Last_Transfer_Time As String, ByVal Update_Time As String,
                 ByVal WEIGHT As Long, ByVal LENGTH As Long, ByVal WIDTH As Long, ByVal HEIGHT As Long, ByVal CARRIER_COMMON01 As String,
                 ByVal CARRIER_COMMON02 As String, ByVal CARRIER_COMMON03 As String, ByVal CARRIER_COMMON04 As String, ByVal CARRIER_COMMON05 As String,
                 ByVal CARRIER_COMMON06 As String, ByVal CARRIER_COMMON07 As String, ByVal CARRIER_COMMON08 As String, ByVal CARRIER_COMMON09 As String,
                 ByVal CARRIER_COMMON10 As String, ByVal SubLocation_X As Long, ByVal SubLocation_Y As Long, ByVal SubLocation_Z As Long, ByVal Update_User_ID As String,
                 ByVal UNPACK_TIME As String, ByVal CREATE_USER_ID As String, ByVal TALLY_ENABLE As Boolean)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(Carrier_ID)
      _gid = key
      _Carrier_ID = Carrier_ID
      _Carrier_Alis = Carrier_Alis
      _Carrier_Desc = Carrier_Desc
      _Carrier_Type = Carrier_Type
      _Carrier_Mode = Carrier_Mode
      _Create_Time = Create_Time
      _Comments = Comments
      _SubLocation_Index_X = SubLocation_Index_X
      _SubLocation_Index_Y = SubLocation_Index_Y
      _SubLocation_Index_Z = SubLocation_Index_Z
      _WEIGHT = WEIGHT
      _LENGTH = LENGTH
      _WIDTH = WIDTH
      _HEIGHT = HEIGHT
      _CARRIER_COMMON01 = CARRIER_COMMON01
      _CARRIER_COMMON02 = CARRIER_COMMON02
      _CARRIER_COMMON03 = CARRIER_COMMON03
      _CARRIER_COMMON04 = CARRIER_COMMON04
      _CARRIER_COMMON05 = CARRIER_COMMON05
      _CARRIER_COMMON06 = CARRIER_COMMON06
      _CARRIER_COMMON07 = CARRIER_COMMON07
      _CARRIER_COMMON08 = CARRIER_COMMON08
      _CARRIER_COMMON09 = CARRIER_COMMON09
      _CARRIER_COMMON10 = CARRIER_COMMON10
      _Location_No = Location_No
      _Reserved = Reserved
      _Locked = Locked
      _Locked_User = Locked_User
      _Locked_Reason = Locked_Reason
      _Locked_Time = Locked_Time
      _Stage_ID = Stage_ID
      _Return_Location_No = Return_Location_No
      _Stocktaking_Time = Stocktaking_Time
      _Firstin_Time = Firstin_Time
      _Last_Transfer_Time = Last_Transfer_Time
      _Update_Time = Update_Time
      _SubLocation_X = SubLocation_X
      _SubLocation_Y = SubLocation_Y
      _SubLocation_Z = SubLocation_Z
      _Update_User_ID = Update_User_ID
      _UNPACK_TIME = UNPACK_TIME
      _TALLY_ENABLE = TALLY_ENABLE
      _CREATE_USER_ID = CREATE_USER_ID

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
  End Sub

  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal Carrier_ID As String) As String
    Try
      Dim key As String = Carrier_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsCarrier
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String), ByRef lstQueueSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_Carrier_StatusManagement.GetInsertSQL(Me)
      Dim _strSQL As String = WMS_M_CarrierManagement.GetInsertSQL(Me)

      lstSQL.Add(strSQL)
      lstSQL.Add(_strSQL)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String), Optional ByRef lstQueueSQL As List(Of String) = Nothing) As Boolean
    Try
      Dim strSQL As String = WMS_T_Carrier_StatusManagement.GetUpdateSQL(Me)
      Dim _strSQL As String = WMS_M_CarrierManagement.GetUpdateSQL(Me)

      lstSQL.Add(strSQL)
      lstSQL.Add(_strSQL)


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Delete的SQL
  Public Function O_Add_Delete_SQLString(ByRef lstSQL As List(Of String), Optional ByRef lstQueueSQL As List(Of String) = Nothing) As Boolean
    Try
      Dim strSQL As String = WMS_T_Carrier_StatusManagement.GetDeleteSQL(Me)
      Dim _strSQL As String = WMS_M_CarrierManagement.GetDeleteSQL(Me)

      lstSQL.Add(strSQL)
      lstSQL.Add(_strSQL)


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '資料加入Dictionary

  '資料從Dictionary刪除

  '檢產Dictionary內是否有該資料




  '取得Carrier Status的Update SQL
  Public Function O_Add_CarrierStatus_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      '在資料庫Update 的資料
      Dim strSQL As String = WMS_T_Carrier_StatusManagement.GetUpdateSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function



End Class
