Public Class clsPO
  Private ShareName As String = "PO"
  Private ShareKey As String = ""

  Private _gid As String
  Private _PO_ID As String
  Private _PO_Type1 As String
  Private _PO_Type2 As String
  Private _PO_Type3 As String
  Private _WO_Type As enuWOType
  Private _Priority As Long
  Private _Create_Time As String
  Private _Start_Time As String
  Private _Finish_Time As String
  Private _User_ID As String
  Private _Customer_No As String
  Private _Class_NO As String
  Private _Shipping_No As String
  Private _PO_Status As enuPOStatus
  Private _Write_Off_No As String
  Private _Auto_Bound As Boolean

  Private _H_PO_CREATE_TIME As String
  Private _H_PO_FINISH_TIME As String
  Private _H_PO_STEP_NO As Long
  Private _H_PO_ORDER_TYPE As String
  Private _H_PO1 As String
  Private _H_PO2 As String
  Private _H_PO3 As String
  Private _H_PO4 As String
  Private _H_PO5 As String
  Private _H_PO6 As String
  Private _H_PO7 As String
  Private _H_PO8 As String
  Private _H_PO9 As String
  Private _H_PO10 As String
  Private _H_PO11 As String
  Private _H_PO12 As String
  Private _H_PO13 As String
  Private _H_PO14 As String
  Private _H_PO15 As String
  Private _H_PO16 As String
  Private _H_PO17 As String
  Private _H_PO18 As String
  Private _H_PO19 As String
  Private _H_PO20 As String
  Private _SUPPLIER_NO As String
  Private _PO_KEY1 As String                    'Vito_19b16
  Private _PO_KEY2 As String                    'Vito_19b16
  Private _PO_KEY3 As String                    'Vito_19b16
  Private _PO_KEY4 As String                    'Vito_19b16
  Private _PO_KEY5 As String                    'Vito_19b16

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
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
  Public Property PO_Type1() As String
    Get
      Return _PO_Type1
    End Get
    Set(ByVal value As String)
      _PO_Type1 = value
    End Set
  End Property
  Public Property PO_Type2() As String
    Get
      Return _PO_Type2
    End Get
    Set(ByVal value As String)
      _PO_Type2 = value
    End Set
  End Property
  Public Property PO_Type3() As String
    Get
      Return _PO_Type3
    End Get
    Set(ByVal value As String)
      _PO_Type3 = value
    End Set
  End Property
  Public Property WO_Type() As enuWOType
    Get
      Return _WO_Type
    End Get
    Set(ByVal value As enuWOType)
      _WO_Type = value
    End Set
  End Property
  Public Property Priority() As Long
    Get
      Return _Priority
    End Get
    Set(ByVal value As Long)
      _Priority = value
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
  Public Property Finish_Time() As String
    Get
      Return _Finish_Time
    End Get
    Set(ByVal value As String)
      _Finish_Time = value
    End Set
  End Property
  Public Property User_ID() As String
    Get
      Return _User_ID
    End Get
    Set(ByVal value As String)
      _User_ID = value
    End Set
  End Property
  Public Property Customer_No() As String
    Get
      Return _Customer_No
    End Get
    Set(ByVal value As String)
      _Customer_No = value
    End Set
  End Property
  Public Property Class_No() As String
    Get
      Return _Class_NO
    End Get
    Set(ByVal value As String)
      _Class_NO = value
    End Set
  End Property
  Public Property Shipping_No() As String
    Get
      Return _Shipping_No
    End Get
    Set(ByVal value As String)
      _Shipping_No = value
    End Set
  End Property
  Public Property Write_Off_No() As String
    Get
      Return _Write_Off_No
    End Get
    Set(ByVal value As String)
      _Write_Off_No = value
    End Set
  End Property
  Public Property PO_Status() As enuPOStatus
    Get
      Return _PO_Status
    End Get
    Set(ByVal value As enuPOStatus)
      _PO_Status = value
    End Set
  End Property
  Public Property Auto_Bound() As Boolean
    Get
      Return _Auto_Bound
    End Get
    Set(ByVal value As Boolean)
      _Auto_Bound = value
    End Set
  End Property

  Public Property H_PO_CREATE_TIME() As String
    Get
      Return _H_PO_CREATE_TIME
    End Get
    Set(ByVal value As String)
      _H_PO_CREATE_TIME = value
    End Set
  End Property
  Public Property H_PO_FINISH_TIME() As String
    Get
      Return _H_PO_FINISH_TIME
    End Get
    Set(ByVal value As String)
      _H_PO_FINISH_TIME = value
    End Set
  End Property
  Public Property H_PO_STEP_NO() As Long
    Get
      Return _H_PO_STEP_NO
    End Get
    Set(ByVal value As Long)
      _H_PO_STEP_NO = value
    End Set
  End Property
  Public Property H_PO_ORDER_TYPE() As String
    Get
      Return _H_PO_ORDER_TYPE
    End Get
    Set(ByVal value As String)
      _H_PO_ORDER_TYPE = value
    End Set
  End Property
  Public Property H_PO1() As String
    Get
      Return _H_PO1
    End Get
    Set(ByVal value As String)
      _H_PO1 = value
    End Set
  End Property
  Public Property H_PO2() As String
    Get
      Return _H_PO2
    End Get
    Set(ByVal value As String)
      _H_PO2 = value
    End Set
  End Property
  Public Property H_PO3() As String
    Get
      Return _H_PO3
    End Get
    Set(ByVal value As String)
      _H_PO3 = value
    End Set
  End Property
  Public Property H_PO4() As String
    Get
      Return _H_PO4
    End Get
    Set(ByVal value As String)
      _H_PO4 = value
    End Set
  End Property
  Public Property H_PO5() As String
    Get
      Return _H_PO5
    End Get
    Set(ByVal value As String)
      _H_PO5 = value
    End Set
  End Property
  Public Property H_PO6() As String
    Get
      Return _H_PO6
    End Get
    Set(ByVal value As String)
      _H_PO6 = value
    End Set
  End Property
  Public Property H_PO7() As String
    Get
      Return _H_PO7
    End Get
    Set(ByVal value As String)
      _H_PO7 = value
    End Set
  End Property
  Public Property H_PO8() As String
    Get
      Return _H_PO8
    End Get
    Set(ByVal value As String)
      _H_PO8 = value
    End Set
  End Property
  Public Property H_PO9() As String
    Get
      Return _H_PO9
    End Get
    Set(ByVal value As String)
      _H_PO9 = value
    End Set
  End Property
  Public Property H_PO10() As String
    Get
      Return _H_PO10
    End Get
    Set(ByVal value As String)
      _H_PO10 = value
    End Set
  End Property
  Public Property H_PO11() As String
    Get
      Return _H_PO11
    End Get
    Set(ByVal value As String)
      _H_PO11 = value
    End Set
  End Property
  Public Property H_PO12() As String
    Get
      Return _H_PO12
    End Get
    Set(ByVal value As String)
      _H_PO12 = value
    End Set
  End Property
  Public Property H_PO13() As String
    Get
      Return _H_PO13
    End Get
    Set(ByVal value As String)
      _H_PO13 = value
    End Set
  End Property
  Public Property H_PO14() As String
    Get
      Return _H_PO14
    End Get
    Set(ByVal value As String)
      _H_PO14 = value
    End Set
  End Property
  Public Property H_PO15() As String
    Get
      Return _H_PO15
    End Get
    Set(ByVal value As String)
      _H_PO15 = value
    End Set
  End Property
  Public Property H_PO16() As String
    Get
      Return _H_PO16
    End Get
    Set(ByVal value As String)
      _H_PO16 = value
    End Set
  End Property
  Public Property H_PO17() As String
    Get
      Return _H_PO17
    End Get
    Set(ByVal value As String)
      _H_PO17 = value
    End Set
  End Property
  Public Property H_PO18() As String
    Get
      Return _H_PO18
    End Get
    Set(ByVal value As String)
      _H_PO18 = value
    End Set
  End Property
  Public Property H_PO19() As String
    Get
      Return _H_PO19
    End Get
    Set(ByVal value As String)
      _H_PO19 = value
    End Set
  End Property
  Public Property H_PO20() As String
    Get
      Return _H_PO20
    End Get
    Set(ByVal value As String)
      _H_PO20 = value
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
  Public Property PO_KEY1() As String 'Vito_19b16
    Get
      Return _PO_KEY1
    End Get
    Set(ByVal value As String)
      _PO_KEY1 = value
    End Set
  End Property
  Public Property PO_KEY2() As String 'Vito_19b16
    Get
      Return _PO_KEY2
    End Get
    Set(ByVal value As String)
      _PO_KEY2 = value
    End Set
  End Property
  Public Property PO_KEY3() As String 'Vito_19b16
    Get
      Return _PO_KEY3
    End Get
    Set(ByVal value As String)
      _PO_KEY3 = value
    End Set
  End Property
  Public Property PO_KEY4() As String 'Vito_19b16
    Get
      Return _PO_KEY4
    End Get
    Set(ByVal value As String)
      _PO_KEY4 = value
    End Set
  End Property
  Public Property PO_KEY5() As String 'Vito_19b16
    Get
      Return _PO_KEY5
    End Get
    Set(ByVal value As String)
      _PO_KEY5 = value
    End Set
  End Property

  '物件建立時執行的事件
  Public Sub New(ByVal PO_ID As String, ByVal PO_Type1 As String, ByVal PO_Type2 As String, ByVal PO_Type3 As String, ByVal WO_Type As enuWOType,
                 ByVal Priority As Long, ByVal Create_Time As String, ByVal Start_Time As String, ByVal Finish_Time As String, ByVal User_ID As String,
                 ByVal Customer_No As String, ByVal Class_NO As String, ByVal Shipping_No As String,
                 ByVal PO_Status As enuPOStatus, ByVal Write_Off_No As String, ByVal Auto_Bound As Boolean,
                 ByVal H_PO_CREATE_TIME As String, ByVal H_PO_FINISH_TIME As String, ByVal H_PO_STEP_NO As Long, ByVal H_PO_ORDER_TYPE As String, ByVal H_PO1 As String,
                 ByVal H_PO2 As String, ByVal H_PO3 As String, ByVal H_PO4 As String, ByVal H_PO5 As String, ByVal H_PO6 As String, ByVal H_PO7 As String, ByVal H_PO8 As String,
                 ByVal H_PO9 As String, ByVal H_PO10 As String, ByVal H_PO11 As String, ByVal H_PO12 As String, ByVal H_PO13 As String, ByVal H_PO14 As String, ByVal H_PO15 As String,
                 ByVal H_PO16 As String, ByVal H_PO17 As String, ByVal H_PO18 As String, ByVal H_PO19 As String, ByVal H_PO20 As String, ByVal SUPPLIER_NO As String,
                 ByVal PO_KEY1 As String, ByVal PO_KEY2 As String, ByVal PO_KEY3 As String, ByVal PO_KEY4 As String, ByVal PO_KEY5 As String) 'Vito_19b16
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PO_ID)  'Vito_19b16
      _gid = key
      _PO_ID = PO_ID
      _PO_Type1 = PO_Type1
      _PO_Type2 = PO_Type2
      _PO_Type3 = PO_Type3
      _Priority = Priority
      _Create_Time = Create_Time
      _Start_Time = Start_Time
      _Finish_Time = Finish_Time
      _User_ID = User_ID
      _Customer_No = Customer_No
      _Class_NO = Class_NO
      _Shipping_No = Shipping_No
      _PO_Status = PO_Status
      _WO_Type = WO_Type
      _Write_Off_No = Write_Off_No
      _Auto_Bound = Auto_Bound

      _H_PO_CREATE_TIME = H_PO_CREATE_TIME
      _H_PO_FINISH_TIME = H_PO_FINISH_TIME
      _H_PO_STEP_NO = H_PO_STEP_NO
      _H_PO_ORDER_TYPE = H_PO_ORDER_TYPE
      _H_PO1 = H_PO1
      _H_PO2 = H_PO2
      _H_PO3 = H_PO3
      _H_PO4 = H_PO4
      _H_PO5 = H_PO5
      _H_PO6 = H_PO6
      _H_PO7 = H_PO7
      _H_PO8 = H_PO8
      _H_PO9 = H_PO9
      _H_PO10 = H_PO10
      _H_PO11 = H_PO11
      _H_PO12 = H_PO12
      _H_PO13 = H_PO13
      _H_PO14 = H_PO14
      _H_PO15 = H_PO15
      _H_PO16 = H_PO16
      _H_PO17 = H_PO17
      _H_PO18 = H_PO18
      _H_PO19 = H_PO19
      _H_PO20 = H_PO20
      _SUPPLIER_NO = SUPPLIER_NO
      _PO_KEY1 = PO_KEY1              'Vito_19b16
      _PO_KEY2 = PO_KEY2              'Vito_19b16
      _PO_KEY3 = PO_KEY3              'Vito_19b16
      _PO_KEY4 = PO_KEY4              'Vito_19b16
      _PO_KEY5 = PO_KEY5              'Vito_19b16
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
  Public Shared Function Get_Combination_Key(ByVal PO_ID As String) As String
    Try
      Dim key As String = PO_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPO
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_POManagement.GetInsertSQL(Me)
      If strSQL IsNot Nothing AndAlso strSQL <> "" Then
        lstSQL.Add(strSQL)
        Return True
      Else
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_POManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_POManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsPO) As Boolean
    Try
      Dim key As String = obj._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _PO_ID = obj.PO_ID
      _PO_Type1 = obj.PO_Type1
      _PO_Type2 = obj.PO_Type2
      _PO_Type3 = obj.PO_Type3
      _Priority = obj.Priority
      _Create_Time = obj.Create_Time
      _Start_Time = obj.Start_Time
      _Finish_Time = obj.Finish_Time
      _User_ID = obj.User_ID
      _Customer_No = obj.Customer_No
      _Class_NO = obj.Class_No
      _Shipping_No = obj.Shipping_No
      _PO_Status = obj.PO_Status
      _WO_Type = obj.WO_Type
      _Write_Off_No = obj.Write_Off_No
      _Auto_Bound = obj.Auto_Bound
      _H_PO_CREATE_TIME = obj.H_PO_CREATE_TIME
      _H_PO_FINISH_TIME = obj.H_PO_FINISH_TIME
      _H_PO_STEP_NO = obj.H_PO_STEP_NO
      _H_PO_ORDER_TYPE = obj.H_PO_ORDER_TYPE
      _H_PO1 = obj.H_PO1
      _H_PO2 = obj.H_PO2
      _H_PO3 = obj.H_PO3
      _H_PO4 = obj.H_PO4
      _H_PO5 = obj.H_PO5
      _H_PO6 = obj.H_PO6
      _H_PO7 = obj.H_PO7
      _H_PO8 = obj.H_PO8
      _H_PO9 = obj.H_PO9
      _H_PO10 = obj.H_PO10
      _H_PO11 = obj.H_PO11
      _H_PO12 = obj.H_PO12
      _H_PO13 = obj.H_PO13
      _H_PO14 = obj.H_PO14
      _H_PO15 = obj.H_PO15
      _H_PO16 = obj.H_PO16
      _H_PO17 = obj.H_PO17
      _H_PO18 = obj.H_PO18
      _H_PO19 = obj.H_PO19
      _H_PO20 = obj.H_PO20
      _SUPPLIER_NO = obj.SUPPLIER_NO
      _PO_KEY1 = obj._PO_KEY1           'Vito_19b16
      _PO_KEY2 = obj._PO_KEY2           'Vito_19b16
      _PO_KEY3 = obj._PO_KEY3           'Vito_19b16
      _PO_KEY4 = obj._PO_KEY4           'Vito_19b16
      _PO_KEY5 = obj._PO_KEY5           'Vito_19b16
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
