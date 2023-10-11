Public Class clsProduce_Info
  Private ShareName As String = "Produce_Info"
  Private ShareKey As String = ""

  Private _gid As String
  Private _Factory_No As String
  Private _Area_No As String
  Private _PO_ID As String
  Private _SKU_NO As String
  Private _Status As enuProduceStatus
  Private _Qty As Double
  Private _Qty_Process As Double
  Private _Previous_Qty_Process As Double
  Private _Qty_NG As Double
  Private _Previous_Qty_NG As Double
  Private _Previous_Area_No As String
  Private _Create_Time As String
  Private _Start_Time As String
  Private _Update_Time As String
  Private _Finish_Time As String
  Private _PO_Info1 As String
  Private _PO_Info2 As String
  Private _PO_Info3 As String
  Private _PO_Info4 As String
  Private _PO_Info5 As String
  Private _PO_Info6 As String
  Private _PO_Info7 As String
  Private _PO_Info8 As String
  Private _PO_Info9 As String
  Private _PO_Info10 As String

  Private _objHandling As clsHandlingObject

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
  Public Property PO_ID() As String
    Get
      Return _PO_ID
    End Get
    Set(ByVal value As String)
      _PO_ID = value
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
  Public Property Status() As enuProduceStatus
    Get
      Return _Status
    End Get
    Set(ByVal value As enuProduceStatus)
      _Status = value
    End Set
  End Property
  Public Property Qty() As Double
    Get
      Return _Qty
    End Get
    Set(ByVal value As Double)
      _Qty = value
    End Set
  End Property
  Public Property Qty_Process() As Double
    Get
      Return _Qty_Process
    End Get
    Set(ByVal value As Double)
      _Qty_Process = value
    End Set
  End Property
  Public Property Previous_Qty_Process() As Double
    Get
      Return _Previous_Qty_Process
    End Get
    Set(ByVal value As Double)
      _Previous_Qty_Process = value
    End Set
  End Property
  Public Property Qty_NG() As Double
    Get
      Return _Qty_NG
    End Get
    Set(ByVal value As Double)
      _Qty_NG = value
    End Set
  End Property
  Public Property Previous_Qty_NG() As Double
    Get
      Return _Previous_Qty_NG
    End Get
    Set(ByVal value As Double)
      _Previous_Qty_NG = value
    End Set
  End Property
  Public Property Previous_Area_No() As String
    Get
      Return _Previous_Area_No
    End Get
    Set(ByVal value As String)
      _Previous_Area_No = value
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
  Public Property Start_Time() As String
    Get
      Return _Start_Time
    End Get
    Set(ByVal value As String)
      _Start_Time = value
    End Set
  End Property
  Public Property Update_Time() As String
    Get
      Return _Update_Time
    End Get
    Set(ByVal value As String)
      _Update_Time = value
    End Set
  End Property
  Public Property Finish_Time() As String
    Get
      Return _Finish_Time
    End Get
    Set(ByVal value As String)
      _Finish_Time = value
    End Set
  End Property
  Public Property PO_Info1() As String
    Get
      Return _PO_Info1
    End Get
    Set(ByVal value As String)
      _PO_Info1 = value
    End Set
  End Property
  Public Property PO_Info2() As String
    Get
      Return _PO_Info2
    End Get
    Set(ByVal value As String)
      _PO_Info2 = value
    End Set
  End Property
  Public Property PO_Info3() As String
    Get
      Return _PO_Info3
    End Get
    Set(ByVal value As String)
      _PO_Info3 = value
    End Set
  End Property
  Public Property PO_Info4() As String
    Get
      Return _PO_Info4
    End Get
    Set(ByVal value As String)
      _PO_Info4 = value
    End Set
  End Property
  Public Property PO_Info5() As String
    Get
      Return _PO_Info5
    End Get
    Set(ByVal value As String)
      _PO_Info5 = value
    End Set
  End Property
  Public Property PO_Info6() As String
    Get
      Return _PO_Info6
    End Get
    Set(ByVal value As String)
      _PO_Info6 = value
    End Set
  End Property
  Public Property PO_Info7() As String
    Get
      Return _PO_Info7
    End Get
    Set(ByVal value As String)
      _PO_Info7 = value
    End Set
  End Property
  Public Property PO_Info8() As String
    Get
      Return _PO_Info8
    End Get
    Set(ByVal value As String)
      _PO_Info8 = value
    End Set
  End Property
  Public Property PO_Info9() As String
    Get
      Return _PO_Info9
    End Get
    Set(ByVal value As String)
      _PO_Info9 = value
    End Set
  End Property
  Public Property PO_Info10() As String
    Get
      Return _PO_Info10
    End Get
    Set(ByVal value As String)
      _PO_Info10 = value
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
  Public Sub New(ByVal Factory_No As String,
                 ByVal Area_No As String,
                 ByVal PO_ID As String,
                 ByVal SKU_NO As String,
                 ByVal Status As enuProduceStatus,
                 ByVal Qty As Double,
                 ByVal Qty_Process As Double,
                 ByVal Previous_Qty_Process As Double,
                 ByVal Qty_NG As Double,
                 ByVal Previous_Qty_NG As Double,
                 ByVal Previous_Area_No As String,
                 ByVal Create_Time As String,
                 ByVal Start_Time As String,
                 ByVal Update_Time As String,
                 ByVal Finish_Time As String,
                 ByVal PO_Info1 As String,
                 ByVal PO_Info2 As String,
                 ByVal PO_Info3 As String,
                 ByVal PO_Info4 As String,
                 ByVal PO_Info5 As String,
                 ByVal PO_Info6 As String,
                 ByVal PO_Info7 As String,
                 ByVal PO_Info8 As String,
                 ByVal PO_Info9 As String,
                 ByVal PO_Info10 As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(Factory_No, Area_No, PO_ID, SKU_NO)
      _gid = key
      _Factory_No = Factory_No
      _Area_No = Area_No
      _PO_ID = PO_ID
      _SKU_NO = SKU_NO
      _Status = Status
      _Qty = Qty
      _Qty_Process = Qty_Process
      _Previous_Qty_Process = Previous_Qty_Process
      _Qty_NG = Qty_NG
      _Previous_Qty_NG = Previous_Qty_NG
      _Previous_Area_No = Previous_Area_No
      _Create_Time = Create_Time
      _Start_Time = Start_Time
      _Update_Time = Update_Time
      _Finish_Time = Finish_Time
      _PO_Info1 = PO_Info1
      _PO_Info2 = PO_Info2
      _PO_Info3 = PO_Info3
      _PO_Info4 = PO_Info4
      _PO_Info5 = PO_Info5
      _PO_Info6 = PO_Info6
      _PO_Info7 = PO_Info7
      _PO_Info8 = PO_Info8
      _PO_Info9 = PO_Info9
      _PO_Info10 = PO_Info10
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
  Public Shared Function Get_Combination_Key(ByVal Factory_No As String,
                                             ByVal Area_No As String,
                                             ByVal PO_ID As String,
                                             ByVal SKU_NO As String) As String
    Try
      Dim key As String = Factory_No & LinkKey & Area_No & LinkKey & PO_ID & LinkKey & SKU_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsProduce_Info
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
        objHandling.O_Add_Produce_Info(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_Produce_Info(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_CT_PRODUCE_INFOManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_CT_PRODUCE_INFOManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_CT_PRODUCE_INFOManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '非標準的Function
  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsProduce_Info) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> gid Then
        SendMessageToLog("Key can not Update, old_Key=" & gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _Factory_No = obj.Factory_No
      _Area_No = obj.Area_No
      _PO_ID = obj.PO_ID
      _SKU_NO = obj.SKU_NO
      _Status = obj.Status
      _Qty = obj.Qty
      _Qty_Process = obj.Qty_Process
      _Previous_Qty_Process = obj.Previous_Qty_Process
      _Qty_NG = obj.Qty_NG
      _Previous_Qty_NG = obj.Previous_Qty_NG
      _Previous_Area_No = obj.Previous_Area_No
      _Create_Time = obj.Create_Time
      _Start_Time = obj.Start_Time
      _Update_Time = obj.Update_Time
      _Finish_Time = obj.Finish_Time
      _PO_Info1 = obj.PO_Info1
      _PO_Info2 = obj.PO_Info2
      _PO_Info3 = obj.PO_Info3
      _PO_Info4 = obj.PO_Info4
      _PO_Info5 = obj.PO_Info5
      _PO_Info6 = obj.PO_Info6
      _PO_Info7 = obj.PO_Info7
      _PO_Info8 = obj.PO_Info8
      _PO_Info9 = obj.PO_Info9
      _PO_Info10 = obj.PO_Info10
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
