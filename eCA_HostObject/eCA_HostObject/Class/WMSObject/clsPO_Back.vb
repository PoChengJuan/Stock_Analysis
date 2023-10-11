Public Class clsPO_Back
  Private ShareName As String = "PO"
  Private ShareKey As String = ""

  Private gid As String  '等於PO_ID
  Private gPO_ID As String
  Private gPO_Type_1 As String
  Private gPO_Type_2 As String
  Private gPO_Type_3 As String
  Private gPriority As Integer
  Private gCreate_Time As String
  Private gStart_Time As String
  Private gFinish_Time As String
  Private gUser_ID As String
  Private gCustomer_No As String
  Private gClass_NO As String
  Private gShipping_No As String
  Private gFactory_NO As String
  Private gDest_Area_NO As String
  Private gOWNER_NO As String
  Private gPO_STATUS As enuPOStatus
  Private gWO_Type As enuWOType
  Private gWRITE_OFF_NO As String

  Private gHOST_CUSTOMER_ID As String
  Private gHOST_CUSTOMER_COMMON1 As String
  Private gHOST_CUSTOMER_COMMON2 As String
  Private gHOST_CUSTOMER_COMMON3 As String
  Private gHOST_CUSTOMER_COMMON4 As String
  Private gHOST_CUSTOMER_COMMON5 As String
  Private gHOST_OWNER_ID As String
  Private gHOST_OWNER_COMMON1 As String
  Private gHOST_OWNER_COMMON2 As String
  Private gHOST_OWNER_COMMON3 As String
  Private gHOST_OWNER_COMMON4 As String
  Private gHOST_OWNER_COMMON5 As String
  Private gHOST_COMMON1 As String
  Private gHOST_COMMON2 As String
  Private gHOST_COMMON3 As String
  Private gHOST_COMMON4 As String
  Private gHOST_COMMON5 As String
  Private gHOST_COMMON6 As String
  Private gHOST_COMMON7 As String
  Private gHOST_COMMON8 As String
  Private gHOST_COMMON9 As String
  Private gHOST_COMMON10 As String
  Private gHOST_COMMON11 As String
  Private gHOST_COMMON12 As String
  Private gHOST_COMMON13 As String
  Private gHOST_COMMON14 As String
  Private gHOST_COMMON15 As String
  Private gHOST_COMMON16 As String
  Private gHOST_COMMON17 As String
  Private gHOST_COMMON18 As String
  Private gHOST_COMMON19 As String
  Private gHOST_COMMON20 As String
  Private gHOST_COMMENTS As String
  Private gHOST_CREATE_TIME As String
  Private gHOST_FINISH_TIME As String
  Private gHOST_STEP_NO As Integer
  Private gHOST_OrderType As enuOrderType







  Private gobjWMS As clsHandlingObject
  '1.PO_DTL
  Public gdicPO_DTL As New Concurrent.ConcurrentDictionary(Of String, clsPO_DTL_Bak)


  '物件建立時執行的事件
  Public Sub New(ByVal PO_ID As String, ByVal gPO_Type_1 As String, ByVal gPO_Type_2 As String, ByVal gPO_Type_3 As String, ByVal Priority As Integer,
                  ByVal Create_Time As String, ByVal Start_Time As String, ByVal Finish_Time As String, ByVal User_ID As String,
                  ByVal Customer_No As String, ByVal Class_NO As String, ByVal Shipping_No As String,
                  ByVal Factory_No As String, ByVal Dest_Area_NO As String, ByVal OWNER_NO As String, ByVal PO_STATUS As enuPOStatus,
                  ByVal WO_Type As enuWOType, ByVal WRITE_OFF_NO As String,
                  ByVal HOST_CREATE_TIME As String, ByVal HOST_FINISH_TIME As String, ByVal HOST_STEP_NO As Integer, ByVal HOST_OrderType As enuOrderType,
                  ByVal HOST_CUSTOMER_NO As String, ByVal HOST_CUSTOMER_ID As String,
                  ByVal HOST_CUSTOMER_COMMON1 As String, ByVal HOST_CUSTOMER_COMMON2 As String, ByVal HOST_CUSTOMER_COMMON3 As String, ByVal HOST_CUSTOMER_COMMON4 As String,
                  ByVal HOST_CUSTOMER_COMMON5 As String, ByVal HOST_OWNER_NO As String, ByVal HOST_OWNER_ID As String, ByVal HOST_OWNER_COMMON1 As String, ByVal HOST_OWNER_COMMON2 As String,
                  ByVal HOST_OWNER_COMMON3 As String, ByVal HOST_OWNER_COMMON4 As String, ByVal HOST_OWNER_COMMON5 As String, ByVal HOST_COMMON1 As String, ByVal HOST_COMMON2 As String,
                  ByVal HOST_COMMON3 As String, ByVal HOST_COMMON4 As String, ByVal HOST_COMMON5 As String, ByVal HOST_COMMON6 As String, ByVal HOST_COMMON7 As String, ByVal HOST_COMMON8 As String,
                  ByVal HOST_COMMON9 As String, ByVal HOST_COMMON10 As String, ByVal HOST_COMMON11 As String, ByVal HOST_COMMON12 As String, ByVal HOST_COMMON13 As String, ByVal HOST_COMMON14 As String,
                  ByVal HOST_COMMON15 As String, ByVal HOST_COMMON16 As String, ByVal HOST_COMMON17 As String, ByVal HOST_COMMON18 As String,
                  ByVal HOST_COMMON19 As String, ByVal HOST_COMMON20 As String, ByVal HOST_COMMENTS As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PO_ID)
      set_gid(key)
      set_PO_ID(PO_ID)
      set_PO_Type_1(gPO_Type_1)
      set_PO_Type_2(gPO_Type_2)
      set_PO_Type_3(gPO_Type_3)
      set_Priority(Priority)
      set_Create_Time(Create_Time)
      set_Start_Time(Start_Time)
      set_Finish_Time(Finish_Time)
      set_User_ID(User_ID)
      set_Customer_No(Customer_No)
      set_Class_NO(Class_NO)
      set_Shipping_No(Shipping_No)
      set_Factory(Factory_No)
      set_Dest_Area_NO(Dest_Area_NO)
      set_OWNER_NO(OWNER_NO)
      set_PO_STATUS(PO_STATUS)
      set_WO_Type(WO_Type)
      set_WRITE_OFF_NO(WRITE_OFF_NO)


      set_HOST_CUSTOMER_ID(HOST_CUSTOMER_ID)
      set_HOST_CUSTOMER_COMMON1(HOST_CUSTOMER_COMMON1)
      set_HOST_CUSTOMER_COMMON2(HOST_CUSTOMER_COMMON2)
      set_HOST_CUSTOMER_COMMON3(HOST_CUSTOMER_COMMON3)
      set_HOST_CUSTOMER_COMMON4(HOST_CUSTOMER_COMMON4)
      set_HOST_CUSTOMER_COMMON5(HOST_CUSTOMER_COMMON5)
      set_HOST_OWNER_ID(HOST_OWNER_ID)
      set_HOST_OWNER_COMMON1(HOST_OWNER_COMMON1)
      set_HOST_OWNER_COMMON2(HOST_OWNER_COMMON2)
      set_HOST_OWNER_COMMON3(HOST_OWNER_COMMON3)
      set_HOST_OWNER_COMMON4(HOST_OWNER_COMMON4)
      set_HOST_OWNER_COMMON5(HOST_OWNER_COMMON5)
      set_HOST_COMMON1(HOST_COMMON1)
      set_HOST_COMMON2(HOST_COMMON2)
      set_HOST_COMMON3(HOST_COMMON3)
      set_HOST_COMMON4(HOST_COMMON4)
      set_HOST_COMMON5(HOST_COMMON5)
      set_HOST_COMMON6(HOST_COMMON6)
      set_HOST_COMMON7(HOST_COMMON7)
      set_HOST_COMMON8(HOST_COMMON8)
      set_HOST_COMMON9(HOST_COMMON9)
      set_HOST_COMMON10(HOST_COMMON10)
      set_HOST_COMMON11(HOST_COMMON11)
      set_HOST_COMMON12(HOST_COMMON12)
      set_HOST_COMMON13(HOST_COMMON13)
      set_HOST_COMMON14(HOST_COMMON14)
      set_HOST_COMMON15(HOST_COMMON15)
      set_HOST_COMMON16(HOST_COMMON16)
      set_HOST_COMMON17(HOST_COMMON17)
      set_HOST_COMMON18(HOST_COMMON18)
      set_HOST_COMMON19(HOST_COMMON19)
      set_HOST_COMMON20(HOST_COMMON20)
      set_HOST_COMMENTS(HOST_COMMENTS)
      set_HOST_CREATE_TIME(HOST_CREATE_TIME)
      set_HOST_FINISH_TIME(HOST_FINISH_TIME)
      set_HOST_STEP_NO(HOST_STEP_NO)
      set_HOST_OrderType(HOST_OrderType)




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
    gobjWMS = Nothing
    gdicPO_DTL = Nothing
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
  Public Function Clone() As clsPO_Back
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Sub Add_Relationship(ByRef objWMS As clsHandlingObject)
    Try
      '挷定PO和WMS的關係
      If objWMS IsNot Nothing Then
        set_objWMS(objWMS)
        objWMS.O_Add_PO(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      '解除PO和WMS的關係
      If gobjWMS IsNot Nothing Then
        gobjWMS.O_Remove_PO(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_POManagement_BackUp.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_POManagement_BackUp.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_POManagement_BackUp.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '資料加入Dictionary
  '把PO_DTL加入gcolPO_DTL
  Public Function O_Add_PO_DTL(ByRef obj As clsPO_DTL_Bak) As Boolean
    Try
      Dim key As String = obj.get_gid()
      If Not gdicPO_DTL.ContainsKey(key) Then
        gdicPO_DTL.TryAdd(key, obj)
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
  '把PO_DTL從gcolPO_DTL刪除
  Public Function O_Remove_PO_DTL(ByRef obj As clsPO_DTL_Bak) As Boolean
    Try
      Dim key As String = obj.get_gid()
      If gdicPO_DTL.ContainsKey(key) Then
        gdicPO_DTL.TryRemove(key, obj)
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
  '從gcolPO_DTL取得指定的objPO_DTL
  Public Function O_Get_PO_DTL(ByVal PO_ID As String, ByVal PO_SERIAL_No As String,
                                   Optional ByRef RetObj As clsPO_DTL_Bak = Nothing) As Boolean
    Try
      Dim key As String = clsPO_DTL_Bak.Get_Combination_Key(gPO_ID, PO_SERIAL_No)
      Dim obj As clsPO_DTL_Bak
      If gdicPO_DTL.ContainsKey(key) Then
        obj = gdicPO_DTL.Item(key)
        RetObj = obj
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '檢查Dictionary內的資料
  '檢查gcolPO_DTL內是否有指定的PO_DTL
  Public Function O_Check_PO_DTL(ByVal PO_ID As String, ByVal PO_SERIAL_No As String) As Boolean
    Try
      Dim key As String = clsPO_DTL_Bak.Get_Combination_Key(gPO_ID, PO_SERIAL_No)
      If gdicPO_DTL.ContainsKey(key) Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


  '-供他人使用的GET
  '-得到gid
  Public Function get_gid() As String
    Try
      Return gid
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gPO_ID
  Public Function get_PO_ID() As String
    Try
      Return gPO_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gPO_Type_1
  Public Function get_PO_Type_1() As String
    Try
      Return gPO_Type_1
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gPO_Type2
  Public Function get_PO_Type_2() As String
    Try
      Return gPO_Type_2
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gPO_Type_3
  Public Function get_PO_Type_3() As String
    Try
      Return gPO_Type_3
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gWO_Type
  Public Function get_WO_Type() As enuWOType
    Try
      Return gWO_Type
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gPriority
  Public Function get_Priority() As Integer
    Try
      Return gPriority
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gWRITE_OFF_NO
  Public Function get_WRITE_OFF_NO() As String
    Try
      Return gWRITE_OFF_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCreate_Time
  Public Function get_Create_Time() As String
    Try
      Return gCreate_Time
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gStart_Time
  Public Function get_Start_Time() As String
    Try
      Return gStart_Time
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gFinish_Time
  Public Function get_Finish_Time() As String
    Try
      Return gFinish_Time
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gUser_ID
  Public Function get_User_ID() As String
    Try
      Return gUser_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCustomer_No
  Public Function get_Customer_No() As String
    Try
      Return gCustomer_No
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gClass_NO
  Public Function get_Class_NO() As String
    Try
      Return gClass_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gShipping_ID
  Public Function get_Shipping_No() As String
    Try
      Return gShipping_No
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gFactory
  Public Function get_Factory_NO() As String
    Try
      Return gFactory_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gDest_Area_NO
  Public Function get_Dest_Area_NO() As String
    Try
      Return gDest_Area_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOWNER_NO
  Public Function get_OWNER_NO() As String
    Try
      Return gOWNER_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gPO_STATUS
  Public Function get_PO_STATUS() As enuPOStatus
    Try
      Return gPO_STATUS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '-得到gCUSTOMER_ID
  Public Function get_HOST_CUSTOMER_ID() As String
    Try
      Return gHOST_CUSTOMER_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCUSTOMER_COMMON1
  Public Function get_HOST_CUSTOMER_COMMON1() As String
    Try
      Return gHOST_CUSTOMER_COMMON1
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCUSTOMER_COMMON2
  Public Function get_HOST_CUSTOMER_COMMON2() As String
    Try
      Return gHOST_CUSTOMER_COMMON2
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCUSTOMER_COMMON3
  Public Function get_HOST_CUSTOMER_COMMON3() As String
    Try
      Return gHOST_CUSTOMER_COMMON3
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCUSTOMER_COMMON4
  Public Function get_HOST_CUSTOMER_COMMON4() As String
    Try
      Return gHOST_CUSTOMER_COMMON4
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCUSTOMER_COMMON5
  Public Function get_HOST_CUSTOMER_COMMON5() As String
    Try
      Return gHOST_CUSTOMER_COMMON5
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOWNER_ID
  Public Function get_HOST_OWNER_ID() As String
    Try
      Return gHOST_OWNER_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOWNER_COMMON1
  Public Function get_HOST_OWNER_COMMON1() As String
    Try
      Return gHOST_OWNER_COMMON1
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOWNER_COMMON2
  Public Function get_HOST_OWNER_COMMON2() As String
    Try
      Return gHOST_OWNER_COMMON2
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOWNER_COMMON3
  Public Function get_HOST_OWNER_COMMON3() As String
    Try
      Return gHOST_OWNER_COMMON3
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOWNER_COMMON4
  Public Function get_HOST_OWNER_COMMON4() As String
    Try
      Return gHOST_OWNER_COMMON4
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOWNER_COMMON5
  Public Function get_HOST_OWNER_COMMON5() As String
    Try
      Return gHOST_OWNER_COMMON5
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON1
  Public Function get_HOST_COMMON1() As String
    Try
      Return gHOST_COMMON1
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON2
  Public Function get_HOST_COMMON2() As String
    Try
      Return gHOST_COMMON2
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON3
  Public Function get_HOST_COMMON3() As String
    Try
      Return gHOST_COMMON3
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON4
  Public Function get_HOST_COMMON4() As String
    Try
      Return gHOST_COMMON4
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON5
  Public Function get_HOST_COMMON5() As String
    Try
      Return gHOST_COMMON5
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON6
  Public Function get_HOST_COMMON6() As String
    Try
      Return gHOST_COMMON6
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON7
  Public Function get_HOST_COMMON7() As String
    Try
      Return gHOST_COMMON7
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON8
  Public Function get_HOST_COMMON8() As String
    Try
      Return gHOST_COMMON8
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON9
  Public Function get_HOST_COMMON9() As String
    Try
      Return gHOST_COMMON9
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON10
  Public Function get_HOST_COMMON10() As String
    Try
      Return gHOST_COMMON10
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON11
  Public Function get_HOST_COMMON11() As String
    Try
      Return gHOST_COMMON11
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON12
  Public Function get_HOST_COMMON12() As String
    Try
      Return gHOST_COMMON12
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON13
  Public Function get_HOST_COMMON13() As String
    Try
      Return gHOST_COMMON13
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON14
  Public Function get_HOST_COMMON14() As String
    Try
      Return gHOST_COMMON14
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON15
  Public Function get_HOST_COMMON15() As String
    Try
      Return gHOST_COMMON15
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON16
  Public Function get_HOST_COMMON16() As String
    Try
      Return gHOST_COMMON16
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON17
  Public Function get_HOST_COMMON17() As String
    Try
      Return gHOST_COMMON17
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON18
  Public Function get_HOST_COMMON18() As String
    Try
      Return gHOST_COMMON18
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON19
  Public Function get_HOST_COMMON19() As String
    Try
      Return gHOST_COMMON19
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMON20
  Public Function get_HOST_COMMON20() As String
    Try
      Return gHOST_COMMON20
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function





  '-得到gCOMMENTS
  Public Function get_HOST_COMMENTS() As String
    Try
      Return gHOST_COMMENTS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCREATE_TIME
  Public Function get_HOST_CREATE_TIME() As String
    Try
      Return gHOST_CREATE_TIME
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gFINISH_TIME
  Public Function get_HOST_FINISH_TIME() As String
    Try
      Return gHOST_FINISH_TIME
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gSTEP_NO
  Public Function get_HOST_STEP_NO() As Integer
    Try
      Return gHOST_STEP_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOrderType
  Public Function get_HOST_OrderType() As enuOrderType
    Try
      Return gHOST_OrderType
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function


  '-得到gobjWMS
  Public Function get_objWMS() As clsHandlingObject
    Try
      Return gobjWMS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function


  '=================Private Function=======================
  '-內部私人的SET
  '-設定gid
  Private Sub set_gid(ByVal key As String)
    Try
      gid = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPO_ID
  Private Sub set_PO_ID(ByVal PO_ID As String)
    Try
      gPO_ID = PO_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPO_Type1
  Private Sub set_PO_Type_1(ByVal PO_Type_1 As String)
    Try
      gPO_Type_1 = PO_Type_1
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPO_Type2
  Private Sub set_PO_Type_2(ByVal PO_Type_2 As String)
    Try
      gPO_Type_2 = PO_Type_2
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPO_Type3
  Private Sub set_PO_Type_3(ByVal PO_Type_3 As String)
    Try
      gPO_Type_3 = PO_Type_3
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gWO_Type
  Private Sub set_WO_Type(ByVal WO_Type As enuWOType)
    Try
      gWO_Type = WO_Type
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gWRITE_OFF_NO
  Private Sub set_WRITE_OFF_NO(ByVal WRITE_OFF_NO As String)
    Try
      gWRITE_OFF_NO = WRITE_OFF_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPriority
  Private Sub set_Priority(ByVal Priority As Integer)
    Try
      gPriority = Priority
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCreate_Time
  Private Sub set_Create_Time(ByVal Create_Time As String)
    Try
      gCreate_Time = Create_Time
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gStart_Time
  Private Sub set_Start_Time(ByVal Start_Time As String)
    Try
      gStart_Time = Start_Time
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gFinish_Time
  Private Sub set_Finish_Time(ByVal Finish_Time As String)
    Try
      gFinish_Time = Finish_Time
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gUser_ID
  Private Sub set_User_ID(ByVal User_ID As String)
    Try
      gUser_ID = User_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCustomer_No
  Private Sub set_Customer_No(ByVal Customer_No As String)
    Try
      gCustomer_No = Customer_No
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gClass_NO
  Private Sub set_Class_NO(ByVal Class_NO As String)
    Try
      gClass_NO = Class_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gShipping_ID
  Private Sub set_Shipping_No(ByVal Shipping_ID As String)
    Try
      gShipping_No = Shipping_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gFactory
  Private Sub set_Factory(ByVal Factory As String)
    Try
      gFactory_NO = Factory
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gDest_Area_NO
  Private Sub set_Dest_Area_NO(ByVal Dest_Area_NO As String)
    Try
      gDest_Area_NO = Dest_Area_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gOWNER_NO
  Private Sub set_OWNER_NO(ByVal OWNER_NO As String)
    Try
      gOWNER_NO = OWNER_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPO_STATUS
  Private Sub set_PO_STATUS(ByVal PO_STATUS As enuPOStatus)
    Try
      gPO_STATUS = PO_STATUS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCUSTOMER_ID
  Private Sub set_HOST_CUSTOMER_ID(ByVal key As String)
    Try
      gCustomer_No = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCUSTOMER_COMMON1
  Private Sub set_HOST_CUSTOMER_COMMON1(ByVal key As String)
    Try
      gHOST_CUSTOMER_COMMON1 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCUSTOMER_COMMON2
  Private Sub set_HOST_CUSTOMER_COMMON2(ByVal key As String)
    Try
      gHOST_CUSTOMER_COMMON2 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCUSTOMER_COMMON3
  Private Sub set_HOST_CUSTOMER_COMMON3(ByVal key As String)
    Try
      gHOST_CUSTOMER_COMMON3 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCUSTOMER_COMMON4
  Private Sub set_HOST_CUSTOMER_COMMON4(ByVal key As String)
    Try
      gHOST_CUSTOMER_COMMON4 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCUSTOMER_COMMON5
  Private Sub set_HOST_CUSTOMER_COMMON5(ByVal key As String)
    Try
      gHOST_CUSTOMER_COMMON5 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gOWNER_ID
  Private Sub set_HOST_OWNER_ID(ByVal key As String)
    Try
      gHOST_OWNER_ID = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gOWNER_COMMON1
  Private Sub set_HOST_OWNER_COMMON1(ByVal key As String)
    Try
      gHOST_OWNER_COMMON1 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gOWNER_COMMON2
  Private Sub set_HOST_OWNER_COMMON2(ByVal key As String)
    Try
      gHOST_OWNER_COMMON2 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gOWNER_COMMON3
  Private Sub set_HOST_OWNER_COMMON3(ByVal key As String)
    Try
      gHOST_OWNER_COMMON3 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gOWNER_COMMON4
  Private Sub set_HOST_OWNER_COMMON4(ByVal key As String)
    Try
      gHOST_OWNER_COMMON4 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gOWNER_COMMON5
  Private Sub set_HOST_OWNER_COMMON5(ByVal key As String)
    Try
      gHOST_OWNER_COMMON5 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON1
  Private Sub set_HOST_COMMON1(ByVal key As String)
    Try
      gHOST_COMMON1 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON2
  Private Sub set_HOST_COMMON2(ByVal key As String)
    Try
      gHOST_COMMON2 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON3
  Private Sub set_HOST_COMMON3(ByVal key As String)
    Try
      gHOST_COMMON3 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON4
  Private Sub set_HOST_COMMON4(ByVal key As String)
    Try
      gHOST_COMMON4 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON5
  Private Sub set_HOST_COMMON5(ByVal key As String)
    Try
      gHOST_COMMON5 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON6
  Private Sub set_HOST_COMMON6(ByVal key As String)
    Try
      gHOST_COMMON6 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON7
  Private Sub set_HOST_COMMON7(ByVal key As String)
    Try
      gHOST_COMMON7 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON8
  Private Sub set_HOST_COMMON8(ByVal key As String)
    Try
      gHOST_COMMON8 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON9
  Private Sub set_HOST_COMMON9(ByVal key As String)
    Try
      gHOST_COMMON9 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON10
  Private Sub set_HOST_COMMON10(ByVal key As String)
    Try
      gHOST_COMMON10 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON11
  Private Sub set_HOST_COMMON11(ByVal key As String)
    Try
      gHOST_COMMON11 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON12
  Private Sub set_HOST_COMMON12(ByVal key As String)
    Try
      gHOST_COMMON12 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON13
  Private Sub set_HOST_COMMON13(ByVal key As String)
    Try
      gHOST_COMMON13 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON14
  Private Sub set_HOST_COMMON14(ByVal key As String)
    Try
      gHOST_COMMON14 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON15
  Private Sub set_HOST_COMMON15(ByVal key As String)
    Try
      gHOST_COMMON15 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON16
  Private Sub set_HOST_COMMON16(ByVal key As String)
    Try
      gHOST_COMMON16 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON17
  Private Sub set_HOST_COMMON17(ByVal key As String)
    Try
      gHOST_COMMON17 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON18
  Private Sub set_HOST_COMMON18(ByVal key As String)
    Try
      gHOST_COMMON18 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON19
  Private Sub set_HOST_COMMON19(ByVal key As String)
    Try
      gHOST_COMMON19 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMON20
  Private Sub set_HOST_COMMON20(ByVal key As String)
    Try
      gHOST_COMMON20 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub




  '-設定gCOMMENTS
  Private Sub set_HOST_COMMENTS(ByVal key As String)
    Try
      gHOST_COMMENTS = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCREATE_TIME
  Private Sub set_HOST_CREATE_TIME(ByVal key As String)
    Try
      gHOST_CREATE_TIME = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gFINISH_TIME
  Private Sub set_HOST_FINISH_TIME(ByVal key As String)
    Try
      gHOST_FINISH_TIME = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gSTEP_NO
  Private Sub set_HOST_STEP_NO(ByVal key As Integer)
    Try
      gHOST_STEP_NO = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gOrderType
  Private Sub set_HOST_OrderType(ByVal key As enuOrderType)
    Try
      gHOST_OrderType = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub








  '-設定gobjWMS
  Private Sub set_objWMS(ByVal objWMS As clsHandlingObject)
    Try
      gobjWMS = objWMS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub


  '非標準的Function
  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef objPO As clsPO_Back) As Boolean
    Try
      Dim key As String = objPO.get_gid()
      If key <> get_gid() Then
        SendMessageToLog("Key can not Update, old_Key=" & get_gid() & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      set_PO_ID(objPO.get_PO_ID)
      set_PO_Type_1(objPO.get_PO_Type_1)
      set_PO_Type_2(objPO.get_PO_Type_2)
      set_PO_Type_3(objPO.get_PO_Type_3)
      set_Priority(objPO.get_Priority)
      set_Create_Time(objPO.get_Create_Time)
      set_Start_Time(objPO.get_Start_Time)
      set_Finish_Time(objPO.get_Finish_Time)
      set_User_ID(objPO.get_User_ID)
      set_Customer_No(objPO.get_Customer_No)
      set_Class_NO(objPO.get_Class_NO)
      set_Shipping_No(objPO.get_Shipping_No)
      set_Factory(objPO.get_Factory_NO)
      set_Dest_Area_NO(objPO.get_Dest_Area_NO)



      set_HOST_CUSTOMER_ID(objPO.get_HOST_CUSTOMER_ID)
      set_HOST_CUSTOMER_COMMON1(objPO.get_HOST_CUSTOMER_COMMON1)
      set_HOST_CUSTOMER_COMMON2(objPO.get_HOST_CUSTOMER_COMMON2)
      set_HOST_CUSTOMER_COMMON3(objPO.get_HOST_CUSTOMER_COMMON3)
      set_HOST_CUSTOMER_COMMON4(objPO.get_HOST_CUSTOMER_COMMON4)
      set_HOST_CUSTOMER_COMMON5(objPO.get_HOST_CUSTOMER_COMMON5)
      set_HOST_OWNER_ID(objPO.get_HOST_OWNER_ID)
      set_HOST_OWNER_COMMON1(objPO.get_HOST_OWNER_COMMON1)
      set_HOST_OWNER_COMMON2(objPO.get_HOST_OWNER_COMMON2)
      set_HOST_OWNER_COMMON3(objPO.get_HOST_OWNER_COMMON3)
      set_HOST_OWNER_COMMON4(objPO.get_HOST_OWNER_COMMON4)
      set_HOST_OWNER_COMMON5(objPO.get_HOST_OWNER_COMMON5)
      set_HOST_COMMON1(objPO.get_HOST_COMMON1)
      set_HOST_COMMON2(objPO.get_HOST_COMMON2)
      set_HOST_COMMON3(objPO.get_HOST_COMMON3)
      set_HOST_COMMON4(objPO.get_HOST_COMMON4)
      set_HOST_COMMON5(objPO.get_HOST_COMMON5)
      set_HOST_COMMON6(objPO.get_HOST_COMMON6)
      set_HOST_COMMON7(objPO.get_HOST_COMMON7)
      set_HOST_COMMON8(objPO.get_HOST_COMMON8)
      set_HOST_COMMON9(objPO.get_HOST_COMMON9)
      set_HOST_COMMON10(objPO.get_HOST_COMMON10)
      set_HOST_COMMON11(objPO.get_HOST_COMMON11)
      set_HOST_COMMON12(objPO.get_HOST_COMMON12)
      set_HOST_COMMON13(objPO.get_HOST_COMMON13)
      set_HOST_COMMON14(objPO.get_HOST_COMMON14)
      set_HOST_COMMON15(objPO.get_HOST_COMMON15)
      set_HOST_COMMENTS(objPO.get_HOST_COMMENTS)
      set_HOST_CREATE_TIME(objPO.get_HOST_CREATE_TIME)
      set_HOST_FINISH_TIME(objPO.get_HOST_FINISH_TIME)
      set_HOST_STEP_NO(objPO.get_HOST_STEP_NO)
      set_HOST_OrderType(objPO.get_HOST_OrderType)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function Update_To_Memory_ByMOCAB01_SENDWORKDATA(ByVal TA001 As String, ByVal TA030 As String, ByVal TA007 As String, ByVal TA013 As String, ByVal TA021 As String, ByVal TA063 As String) As Boolean
    Try
      set_PO_Type_1(TA001)
      set_PO_Type_2(TA030)
      set_HOST_COMMON1(TA007)
      set_HOST_COMMON2(TA013)
      set_HOST_COMMON3(TA021)
      set_HOST_COMMON4(TA063)



      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory_ByMOCAB04_SENDWORKIDCHANGEDATA(ByVal TO003 As String, ByVal TO004 As String, ByVal TO034 As String) As Boolean
    Try


      set_HOST_COMMON4(TO034)
      set_HOST_COMMON5(TO003)
      set_HOST_COMMON6(TO004)



      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try

  End Function
  Public Function Update_To_Memory_ByMOCAB02_SENDPICKINGDATA(ByVal TC003 As String, ByVal TC009 As String, ByVal TC001 As String, ByVal TC004 As String, ByVal TC005 As String) As Boolean
    Try
      set_PO_Type_1(TC001)
      set_Factory(TC004)
      set_Dest_Area_NO(TC005)
      set_HOST_COMMON1(TC003)
      set_HOST_COMMON2(TC009)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory_ByCOPAB01_SENDSELLDATA(ByVal TH001 As String, ByVal TG023 As String) As Boolean
    Try
      set_PO_Type_1(TH001)
      set_HOST_COMMON1(TG023)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory_ByPURAB01_SENDPURCHASEDATA(ByVal TH001 As String, ByVal TG013 As String) As Boolean
    Try
      set_PO_Type_1(TH001)
      set_HOST_COMMON1(TG013)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory_ByINVAB03_SENDINVENTORYDATA(ByVal TE004 As String) As Boolean
    Try
      set_HOST_COMMON1(TE004)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory_ByINVAB01_SENDTRANSACTIONDATA(ByVal TA001 As String, ByVal TA009 As String, ByVal TA008 As String, ByVal TA006 As String, ByVal TA003 As String) As Boolean
    Try
      set_PO_Type_1(TA001)
      set_PO_Type_2(TA009)
      set_Factory(TA008)
      set_HOST_COMMON1(TA006)
      set_HOST_COMMON2(TA003)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_POType_Priority_UserID_CustomerNo_ClassNo_ShippingNo_FactoryNo_DestAreaNo_To_Memory(ByVal PO_Type As enuWOType,
                                    ByVal Priority As Long, ByVal User_ID As String, ByVal Customer_No As String, ByVal Class_No As String,
                                    ByVal Shipping_No As String, ByVal Factory_No As String, ByVal Dest_Area_No As String) As Boolean
    Try
      'set_PO_Type(PO_Type)
      set_Priority(Priority)
      set_User_ID(User_ID)
      set_Customer_No(Customer_No)
      set_Class_NO(Class_No)
      set_Shipping_No(Shipping_No)
      set_Factory(Factory_No)
      set_Dest_Area_NO(Dest_Area_No)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function Update_PO_FINISH_TIME() As Boolean
    Try
      set_Finish_Time(ModuleHelpFunc.GetNewTime_DBFormat)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '新增資料到DB
  Public Function O_Insert_PO_To_DB(ByRef objWMS As clsHandlingObject) As Boolean
    Try
      '一定要寫成功，才更新記憶體的狀態			
      If WMS_T_POManagement_BackUp.AddWMS_T_POData(Me) = True Then
        '建立梆定
        Add_Relationship(objWMS)
        Return True
      Else
        SendMessageToLog("Insert PO to DB Failed ,TableName = " & WMS_T_POManagement_BackUp.TableName & " ,key=" & get_gid(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '刪除
  Public Function O_Delete_PO_To_DB() As Boolean
    Try
      '一定要寫成功，才更新記憶體的狀態			
      If WMS_T_POManagement_BackUp.DeleteWMS_T_POData(Me) = True Then
        '解除梆定
        Remove_Relationship()
        Return True
      Else
        SendMessageToLog("Delete PO to DB Failed ,TableName = " & WMS_T_POManagement_BackUp.TableName & " ,key=" & get_gid(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '更新
  Public Function O_Update_CreateTime_To_DB(ByVal CreateTime As String) As Boolean
    Try
      '一定要寫成功，才更新記憶體的狀態
      Dim Update_Time As String = GetNewTime_ByDataTimeFormat(DBTimeFormat)
      Dim New_objProcessFlow As clsPO_Back = Clone()
      New_objProcessFlow.set_Create_Time(CreateTime)

      '在資料庫Update 的資料
      If WMS_T_POManagement_BackUp.UpdateWMS_T_POData(New_objProcessFlow) = True Then
        set_Create_Time(CreateTime)
        Return True
      Else
        SendMessageToLog("Update CreateTime to DB Failed ,TableName = " & WMS_T_POManagement_BackUp.TableName & " ,key=" & get_gid(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Class
