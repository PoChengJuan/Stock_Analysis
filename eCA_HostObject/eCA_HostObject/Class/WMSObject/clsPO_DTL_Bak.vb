Public Class clsPO_DTL_Bak
  Private ShareName As String = "PO_DTL"
  Private ShareKey As String = ""

  Private gid As String  '等於PO_ID
  Private gPO_ID As String '訂單編號
  Private gPO_LINE_NO As String
  Private gPO_SERIAL_NO As String '訂單明細編號
  Private gSKU_No As String '貨品編號
  Private gLOT_NO As String '批號
  Private gQTY As Double '需求量
  Private gQTY_PROCESS As Double '已挷定數量
  Private gQTY_FINISH As Double '已入庫/出庫數量
  Private gCOMMENTS As String '備註
  Private gPACKAGE_ID As String '箱ID/包裝ID
  Private gITEM_COMMON1 As String '條件1
  Private gITEM_COMMON2 As String '條件2
  Private gITEM_COMMON3 As String '條件3
  Private gITEM_COMMON4 As String '條件4
  Private gITEM_COMMON5 As String '條件5
  Private gITEM_COMMON6 As String '條件6
  Private gITEM_COMMON7 As String '條件7
  Private gITEM_COMMON8 As String '條件8
  Private gITEM_COMMON9 As String '條件9
  Private gITEM_COMMON10 As String '條件10

  Private gHOST_FACTORY_NO As String
  Private gHOST_AREA_NO As String
  Private gHOST_PO_TYPE As Integer
  Private gHOST_OWNER_NO As String
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
  Private gHOST_COMMON21 As String
  Private gHOST_COMMON22 As String
  Private gHOST_COMMON23 As String
  Private gHOST_COMMON24 As String
  Private gHOST_COMMON25 As String
  Private gHOST_COMMON26 As String
  Private gHOST_COMMON27 As String
  Private gHOST_COMMON28 As String
  Private gHOST_COMMON29 As String
  Private gHOST_COMMON30 As String
  Private gHOST_COMMON31 As String
  Private gHOST_COMMON32 As String
  Private gHOST_COMMON33 As String
  Private gHOST_COMMON34 As String
  Private gHOST_COMMON35 As String
  Private gHOST_COMMON36 As String
  Private gHOST_COMMON37 As String
  Private gHOST_COMMON38 As String
  Private gHOST_COMMON39 As String
  Private gHOST_COMMON40 As String
  Private gHOST_STEP_NO As enuStepNo
  Private gHOST_MOVE_TYPE As String
  Private gHOST_FINISH_TIME As String
  Private gHOST_BILLING_DATE As String
  Private gHOST_CREATE_TIME As String



  Private gobjWMS As clsHandlingObject
  Private gobjPO As clsPO_Back


  '物件建立時執行的事件
  Public Sub New(ByVal PO_ID As String, ByVal PO_LINE_NO As String, ByVal PO_SERIAL_NO As String, ByVal SKU_No As String, ByVal LOT_NO As String,
                  ByVal QTY As Double, ByVal QTY_PROCESS As Double, ByVal QTY_FINISH As Double,
                  ByVal COMMENTS As String, ByVal PACKAGE_ID As String, ByVal ITEM_COMMON1 As String, ByVal ITEM_COMMON2 As String,
                  ByVal ITEM_COMMON3 As String, ByVal ITEM_COMMON4 As String, ByVal ITEM_COMMON5 As String, ByVal ITEM_COMMON6 As String,
                  ByVal ITEM_COMMON7 As String, ByVal ITEM_COMMON8 As String, ByVal ITEM_COMMON9 As String, ByVal ITEM_COMMON10 As String,
                  ByVal HOST_FACTORY_NO As String, ByVal HOST_AREA_NO As String, ByVal HOST_STEP_NO As String, ByVal HOST_MOVE_TYPE As String, ByVal HOST_FINISH_TIME As String, ByVal HOST_BILLING_DATE As String, ByVal HOST_CREATE_TIME As String,
                  ByVal HOST_OWNER_NO As String, ByVal HOST_COMMON1 As String, ByVal HOST_COMMON2 As String,
                  ByVal HOST_COMMON3 As String, ByVal HOST_COMMON4 As String, ByVal HOST_COMMON5 As String, ByVal HOST_COMMON6 As String, ByVal HOST_COMMON7 As String, ByVal HOST_COMMON8 As String,
                  ByVal HOST_COMMON9 As String, ByVal HOST_COMMON10 As String, ByVal HOST_COMMON11 As String, ByVal HOST_COMMON12 As String, ByVal HOST_COMMON13 As String, ByVal HOST_COMMON14 As String,
                  ByVal HOST_COMMON15 As String, ByVal HOST_COMMON16 As String, ByVal HOST_COMMON17 As String, ByVal HOST_COMMON18 As String, ByVal HOST_COMMON19 As String, ByVal HOST_COMMON20 As String,
                  ByVal HOST_COMMON21 As String, ByVal HOST_COMMON22 As String, ByVal HOST_COMMON23 As String, ByVal HOST_COMMON24 As String, ByVal HOST_COMMON25 As String,
                  ByVal HOST_COMMON26 As String, ByVal HOST_COMMON27 As String, ByVal HOST_COMMON28 As String, ByVal HOST_COMMON29 As String, ByVal HOST_COMMON30 As String,
                  ByVal HOST_COMMON31 As String, ByVal HOST_COMMON32 As String, ByVal HOST_COMMON33 As String, ByVal HOST_COMMON34 As String, ByVal HOST_COMMON35 As String,
                  ByVal HOST_COMMON36 As String, ByVal HOST_COMMON37 As String, ByVal HOST_COMMON38 As String, ByVal HOST_COMMON39 As String, ByVal HOST_COMMON40 As String
                  )
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PO_ID, PO_SERIAL_NO)
      set_gid(key)
      set_PO_ID(PO_ID)
      set_PO_LINE_NO(PO_LINE_NO)
      set_PO_SERIAL_NO(PO_SERIAL_NO)
      set_SKU_NO(SKU_No)
      set_LOT_NO(LOT_NO)
      set_QTY(QTY)
      set_QTY_PROCESS(QTY_PROCESS)
      set_QTY_FINISH(QTY_FINISH)
      set_COMMENTS(COMMENTS)
      set_PACKAGE_ID(PACKAGE_ID)
      set_ITEM_COMMON1(ITEM_COMMON1)
      set_ITEM_COMMON2(ITEM_COMMON2)
      set_ITEM_COMMON3(ITEM_COMMON3)
      set_ITEM_COMMON4(ITEM_COMMON4)
      set_ITEM_COMMON5(ITEM_COMMON5)
      set_ITEM_COMMON6(ITEM_COMMON6)
      set_ITEM_COMMON7(ITEM_COMMON7)
      set_ITEM_COMMON8(ITEM_COMMON8)
      set_ITEM_COMMON9(ITEM_COMMON9)
      set_ITEM_COMMON10(ITEM_COMMON10)

      set_HOST_FACTORY_NO(HOST_FACTORY_NO)
      set_HOST_AREA_NO(HOST_AREA_NO)

      set_HOST_OWNER_NO(HOST_OWNER_NO)
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
      set_HOST_COMMON21(HOST_COMMON21)
      set_HOST_COMMON22(HOST_COMMON22)
      set_HOST_COMMON23(HOST_COMMON23)
      set_HOST_COMMON24(HOST_COMMON24)
      set_HOST_COMMON25(HOST_COMMON25)
      set_HOST_COMMON26(HOST_COMMON26)
      set_HOST_COMMON27(HOST_COMMON27)
      set_HOST_COMMON28(HOST_COMMON28)
      set_HOST_COMMON29(HOST_COMMON29)
      set_HOST_COMMON30(HOST_COMMON30)
      set_HOST_COMMON31(HOST_COMMON31)
      set_HOST_COMMON32(HOST_COMMON32)
      set_HOST_COMMON33(HOST_COMMON33)
      set_HOST_COMMON34(HOST_COMMON34)
      set_HOST_COMMON35(HOST_COMMON35)
      set_HOST_COMMON36(HOST_COMMON36)
      set_HOST_COMMON37(HOST_COMMON37)
      set_HOST_COMMON38(HOST_COMMON38)
      set_HOST_COMMON39(HOST_COMMON39)
      set_HOST_COMMON40(HOST_COMMON40)
      set_HOST_STEP_NO(HOST_STEP_NO)
      set_HOST_MOVE_TYPE(HOST_MOVE_TYPE)
      set_HOST_FINISH_TIME(HOST_FINISH_TIME)
      set_HOST_BILLING_DATE(HOST_BILLING_DATE)
      set_HOST_CREATE_TIME(HOST_CREATE_TIME)

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
    gobjPO = Nothing

  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal PO_ID As String, ByVal PO_SERIAL_NO As String) As String
    Try
      Dim key As String = PO_ID & "_" & PO_SERIAL_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPO_DTL_Bak
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Sub Add_Relationship(ByRef objWMS As clsHandlingObject)
    Try
      '挷定PO_DTL和WMS的關係
      If objWMS IsNot Nothing Then
        set_objWMS(objWMS)
        objWMS.O_Add_PO_DTL(Me)
        '挷定PO_DTL和PO的關係
        Dim objPO As clsPO_Back = Nothing
        If objWMS.O_Get_PO(gPO_ID, objPO) Then
          If objPO IsNot Nothing Then
            set_objPO(objPO)
            objPO.O_Add_PO_DTL(Me)
          End If
        End If
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      '解除PO_DTL和WMS的關係
      If gobjWMS IsNot Nothing Then
        gobjWMS.O_Remove_PO_DTL(Me)
      End If
      '解除PO_DTL和PO的關係
      If gobjPO IsNot Nothing Then
        gobjPO.O_Remove_PO_DTL(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_PO_DTLManagement_BackUp.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_PO_DTLManagement_BackUp.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_PO_DTLManagement_BackUp.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function



  '資料加入Dictionary


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
  '-得到gPO_LINE_NO
  Public Function get_PO_LINE_NO() As String
    Try
      Return gPO_LINE_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gPO_SERIAL_NO
  Public Function get_PO_SERIAL_NO() As String
    Try
      Return gPO_SERIAL_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gLOT_NO
  Public Function get_LOT_NO() As String
    Try
      Return gLOT_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gQTY
  Public Function get_QTY() As String
    Try
      Return gQTY
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gQTY_PROCESS
  Public Function get_QTY_PROCESS() As String
    Try
      Return gQTY_PROCESS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gQTY_FINISH
  Public Function get_QTY_FINISH() As String
    Try
      Return gQTY_FINISH
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gPACKAGE_ID
  Public Function get_PACKAGE_ID() As String
    Try
      Return gPACKAGE_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMENTS
  Public Function get_COMMENTS() As String
    Try
      Return gCOMMENTS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function '
  '-得到gITEM_COMMON1
  Public Function get_ITEM_COMMON1() As String
    Try
      Return gITEM_COMMON1
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON2
  Public Function get_ITEM_COMMON2() As String
    Try
      Return gITEM_COMMON2
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON3
  Public Function get_ITEM_COMMON3() As String
    Try
      Return gITEM_COMMON3
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON4
  Public Function get_ITEM_COMMON4() As String
    Try
      Return gITEM_COMMON4
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON5
  Public Function get_ITEM_COMMON5() As String
    Try
      Return gITEM_COMMON5
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON6
  Public Function get_ITEM_COMMON6() As String
    Try
      Return gITEM_COMMON6
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON7
  Public Function get_ITEM_COMMON7() As String
    Try
      Return gITEM_COMMON7
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON8
  Public Function get_ITEM_COMMON8() As String
    Try
      Return gITEM_COMMON8
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON9
  Public Function get_ITEM_COMMON9() As String
    Try
      Return gITEM_COMMON9
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gITEM_COMMON10
  Public Function get_ITEM_COMMON10() As String
    Try
      Return gITEM_COMMON10
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gFACTORY_NO
  Public Function get_HOST_FACTORY_NO() As String
    Try
      Return gHOST_FACTORY_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gAREA_NO
  Public Function get_HOST_AREA_NO() As String
    Try
      Return gHOST_AREA_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gSKU_NO
  Public Function get_SKU_NO() As String
    Try
      Return gSKU_No
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gOWNER_NO
  Public Function get_HOST_OWNER_NO() As String
    Try
      Return gHOST_OWNER_NO
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
  '-得到gHOST_COMMON4
  Public Function get_HOST_COMMON4() As String
    Try
      Return gHOST_COMMON4
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON5
  Public Function get_HOST_COMMON5() As String
    Try
      Return gHOST_COMMON5
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON6
  Public Function get_HOST_COMMON6() As String
    Try
      Return gHOST_COMMON6
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON7
  Public Function get_HOST_COMMON7() As String
    Try
      Return gHOST_COMMON7
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON8
  Public Function get_HOST_COMMON8() As String
    Try
      Return gHOST_COMMON8
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON9
  Public Function get_HOST_COMMON9() As String
    Try
      Return gHOST_COMMON9
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON10
  Public Function get_HOST_COMMON10() As String
    Try
      Return gHOST_COMMON10
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON11
  Public Function get_HOST_COMMON11() As String
    Try
      Return gHOST_COMMON11
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON12
  Public Function get_HOST_COMMON12() As String
    Try
      Return gHOST_COMMON12
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON13
  Public Function get_HOST_COMMON13() As String
    Try
      Return gHOST_COMMON13
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON14
  Public Function get_HOST_COMMON14() As String
    Try
      Return gHOST_COMMON14
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON15
  Public Function get_HOST_COMMON15() As String
    Try
      Return gHOST_COMMON15
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON16
  Public Function get_HOST_COMMON16() As String
    Try
      Return gHOST_COMMON16
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON17
  Public Function get_HOST_COMMON17() As String
    Try
      Return gHOST_COMMON17
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON18
  Public Function get_HOST_COMMON18() As String
    Try
      Return gHOST_COMMON18
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON19
  Public Function get_HOST_COMMON19() As String
    Try
      Return gHOST_COMMON19
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON20
  Public Function get_HOST_COMMON20() As String
    Try
      Return gHOST_COMMON20
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON21
  Public Function get_HOST_COMMON21() As String
    Try
      Return gHOST_COMMON21
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON22
  Public Function get_HOST_COMMON22() As String
    Try
      Return gHOST_COMMON22
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON23
  Public Function get_HOST_COMMON23() As String
    Try
      Return gHOST_COMMON23
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON24
  Public Function get_HOST_COMMON24() As String
    Try
      Return gHOST_COMMON24
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON25
  Public Function get_HOST_COMMON25() As String
    Try
      Return gHOST_COMMON25
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON26
  Public Function get_HOST_COMMON26() As String
    Try
      Return gHOST_COMMON26
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON27
  Public Function get_HOST_COMMON27() As String
    Try
      Return gHOST_COMMON27
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON28
  Public Function get_HOST_COMMON28() As String
    Try
      Return gHOST_COMMON28
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON29
  Public Function get_HOST_COMMON29() As String
    Try
      Return gHOST_COMMON29
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON30
  Public Function get_HOST_COMMON30() As String
    Try
      Return gHOST_COMMON30
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON31
  Public Function get_HOST_COMMON31() As String
    Try
      Return gHOST_COMMON31
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON32
  Public Function get_HOST_COMMON32() As String
    Try
      Return gHOST_COMMON32
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON33
  Public Function get_HOST_COMMON33() As String
    Try
      Return gHOST_COMMON33
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON34
  Public Function get_HOST_COMMON34() As String
    Try
      Return gHOST_COMMON34
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON35
  Public Function get_HOST_COMMON35() As String
    Try
      Return gHOST_COMMON35
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON36
  Public Function get_HOST_COMMON36() As String
    Try
      Return gHOST_COMMON36
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON37
  Public Function get_HOST_COMMON37() As String
    Try
      Return gHOST_COMMON37
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON38
  Public Function get_HOST_COMMON38() As String
    Try
      Return gHOST_COMMON38
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON39
  Public Function get_HOST_COMMON39() As String
    Try
      Return gHOST_COMMON39
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_COMMON40
  Public Function get_HOST_COMMON40() As String
    Try
      Return gHOST_COMMON40
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gCOMMENTS
  Public Function get_HOST_COMMENTS() As String
    Try
      Return gCOMMENTS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_STEP_NO
  Public Function get_HOST_STEP_NO() As String
    Try
      Return gHOST_STEP_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_MOVE_TYPE
  Public Function get_HOST_MOVE_TYPE() As String
    Try
      Return gHOST_MOVE_TYPE
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_FINISH_TIME
  Public Function get_HOST_FINISH_TIME() As String
    Try
      Return gHOST_FINISH_TIME
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_BILLING_DATE
  Public Function get_HOST_BILLING_DATE() As String
    Try
      Return gHOST_BILLING_DATE
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-得到gHOST_CREATE_TIME
  Public Function get_HOST_CREATE_TIME() As String
    Try
      Return gHOST_CREATE_TIME
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
  '-得到gobjPO
  Public Function get_objPO() As clsPO_Back
    Try
      Return gobjPO
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
  '-設定gPO_LINE_NO
  Private Sub set_PO_LINE_NO(ByVal PO_LINE_NO As String)
    Try
      gPO_LINE_NO = PO_LINE_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPO_SERIAL_NO
  Private Sub set_PO_SERIAL_NO(ByVal PO_SERIAL_NO As String)
    Try
      gPO_SERIAL_NO = PO_SERIAL_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gLOT_NO
  Private Sub set_LOT_NO(ByVal LOT_NO As String)
    Try
      gLOT_NO = LOT_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gQTY
  Private Sub set_QTY(ByVal QTY As String)
    Try
      gQTY = QTY
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gQTY_PROCESS
  Private Sub set_QTY_PROCESS(ByVal QTY_PROCESS As String)
    Try
      gQTY_PROCESS = QTY_PROCESS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gQTY_FINISH
  Private Sub set_QTY_FINISH(ByVal QTY_FINISH As String)
    Try
      gQTY_FINISH = QTY_FINISH
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPACKAGE_ID
  Private Sub set_PACKAGE_ID(ByVal PACKAGE_ID As String)
    Try
      gPACKAGE_ID = PACKAGE_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON1
  Private Sub set_ITEM_COMMON1(ByVal ITEM_COMMON1 As String)
    Try
      gITEM_COMMON1 = ITEM_COMMON1
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON2
  Private Sub set_ITEM_COMMON2(ByVal ITEM_COMMON2 As String)
    Try
      gITEM_COMMON2 = ITEM_COMMON2
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON3
  Private Sub set_ITEM_COMMON3(ByVal ITEM_COMMON3 As String)
    Try
      gITEM_COMMON3 = ITEM_COMMON3
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON4
  Private Sub set_ITEM_COMMON4(ByVal ITEM_COMMON4 As String)
    Try
      gITEM_COMMON4 = ITEM_COMMON4
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON5
  Private Sub set_ITEM_COMMON5(ByVal ITEM_COMMON5 As String)
    Try
      gITEM_COMMON5 = ITEM_COMMON5
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON6
  Private Sub set_ITEM_COMMON6(ByVal ITEM_COMMON6 As String)
    Try
      gITEM_COMMON6 = ITEM_COMMON6
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON7
  Private Sub set_ITEM_COMMON7(ByVal ITEM_COMMON7 As String)
    Try
      gITEM_COMMON7 = ITEM_COMMON7
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON8
  Private Sub set_ITEM_COMMON8(ByVal ITEM_COMMON8 As String)
    Try
      gITEM_COMMON8 = ITEM_COMMON8
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON9
  Private Sub set_ITEM_COMMON9(ByVal ITEM_COMMON9 As String)
    Try
      gITEM_COMMON9 = ITEM_COMMON9
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gITEM_COMMON10
  Private Sub set_ITEM_COMMON10(ByVal ITEM_COMMON10 As String)
    Try
      gITEM_COMMON10 = ITEM_COMMON10
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_FACTORY_NO
  Private Sub set_HOST_FACTORY_NO(ByVal key As String)
    Try
      gHOST_FACTORY_NO = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_AREA_NO
  Private Sub set_HOST_AREA_NO(ByVal key As String)
    Try
      gHOST_AREA_NO = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gSKU_NO
  Private Sub set_SKU_NO(ByVal key As String)
    Try
      gSKU_No = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_OWNER_NO
  Private Sub set_HOST_OWNER_NO(ByVal key As String)
    Try
      gHOST_OWNER_NO = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON1
  Private Sub set_HOST_COMMON1(ByVal key As String)
    Try
      gHOST_COMMON1 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON2
  Private Sub set_HOST_COMMON2(ByVal key As String)
    Try
      gHOST_COMMON2 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON3
  Private Sub set_HOST_COMMON3(ByVal key As String)
    Try
      gHOST_COMMON3 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON4
  Private Sub set_HOST_COMMON4(ByVal key As String)
    Try
      gHOST_COMMON4 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON5
  Private Sub set_HOST_COMMON5(ByVal key As String)
    Try
      gHOST_COMMON5 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON6
  Private Sub set_HOST_COMMON6(ByVal key As String)
    Try
      gHOST_COMMON6 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON7
  Private Sub set_HOST_COMMON7(ByVal key As String)
    Try
      gHOST_COMMON7 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON8
  Private Sub set_HOST_COMMON8(ByVal key As String)
    Try
      gHOST_COMMON8 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON9
  Private Sub set_HOST_COMMON9(ByVal key As String)
    Try
      gHOST_COMMON9 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON10
  Private Sub set_HOST_COMMON10(ByVal key As String)
    Try
      gHOST_COMMON10 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON11
  Private Sub set_HOST_COMMON11(ByVal key As String)
    Try
      gHOST_COMMON11 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON12
  Private Sub set_HOST_COMMON12(ByVal key As String)
    Try
      gHOST_COMMON12 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON13
  Private Sub set_HOST_COMMON13(ByVal key As String)
    Try
      gHOST_COMMON13 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON14
  Private Sub set_HOST_COMMON14(ByVal key As String)
    Try
      gHOST_COMMON14 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON15
  Private Sub set_HOST_COMMON15(ByVal key As String)
    Try
      gHOST_COMMON15 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON16
  Private Sub set_HOST_COMMON16(ByVal key As String)
    Try
      gHOST_COMMON16 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON17
  Private Sub set_HOST_COMMON17(ByVal key As String)
    Try
      gHOST_COMMON17 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON18
  Private Sub set_HOST_COMMON18(ByVal key As String)
    Try
      gHOST_COMMON18 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON19
  Private Sub set_HOST_COMMON19(ByVal key As String)
    Try
      gHOST_COMMON19 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON20
  Private Sub set_HOST_COMMON20(ByVal key As String)
    Try
      gHOST_COMMON20 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON21
  Private Sub set_HOST_COMMON21(ByVal key As String)
    Try
      gHOST_COMMON21 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON22
  Private Sub set_HOST_COMMON22(ByVal key As String)
    Try
      gHOST_COMMON22 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON23
  Private Sub set_HOST_COMMON23(ByVal key As String)
    Try
      gHOST_COMMON23 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON24
  Private Sub set_HOST_COMMON24(ByVal key As String)
    Try
      gHOST_COMMON24 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON25
  Private Sub set_HOST_COMMON25(ByVal key As String)
    Try
      gHOST_COMMON25 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON26
  Private Sub set_HOST_COMMON26(ByVal key As String)
    Try
      gHOST_COMMON26 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON27
  Private Sub set_HOST_COMMON27(ByVal key As String)
    Try
      gHOST_COMMON27 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON28
  Private Sub set_HOST_COMMON28(ByVal key As String)
    Try
      gHOST_COMMON28 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON29
  Private Sub set_HOST_COMMON29(ByVal key As String)
    Try
      gHOST_COMMON29 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON30
  Private Sub set_HOST_COMMON30(ByVal key As String)
    Try
      gHOST_COMMON30 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON31
  Private Sub set_HOST_COMMON31(ByVal key As String)
    Try
      gHOST_COMMON31 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON32
  Private Sub set_HOST_COMMON32(ByVal key As String)
    Try
      gHOST_COMMON32 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON33
  Private Sub set_HOST_COMMON33(ByVal key As String)
    Try
      gHOST_COMMON33 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON34
  Private Sub set_HOST_COMMON34(ByVal key As String)
    Try
      gHOST_COMMON34 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON35
  Private Sub set_HOST_COMMON35(ByVal key As String)
    Try
      gHOST_COMMON35 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON36
  Private Sub set_HOST_COMMON36(ByVal key As String)
    Try
      gHOST_COMMON36 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON37
  Private Sub set_HOST_COMMON37(ByVal key As String)
    Try
      gHOST_COMMON37 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON38
  Private Sub set_HOST_COMMON38(ByVal key As String)
    Try
      gHOST_COMMON38 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON39
  Private Sub set_HOST_COMMON39(ByVal key As String)
    Try
      gHOST_COMMON39 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_COMMON40
  Private Sub set_HOST_COMMON40(ByVal key As String)
    Try
      gHOST_COMMON40 = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '-設定gCOMMENTS
  Private Sub set_COMMENTS(ByVal key As String)
    Try
      gCOMMENTS = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_STEP_NO
  Private Sub set_HOST_STEP_NO(ByVal key As String)
    Try
      gHOST_STEP_NO = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_MOVE_TYPE
  Private Sub set_HOST_MOVE_TYPE(ByVal key As String)
    Try
      gHOST_MOVE_TYPE = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_FINISH_TIME
  Private Sub set_HOST_FINISH_TIME(ByVal key As String)
    Try
      gHOST_FINISH_TIME = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_BILLING_DATE
  Private Sub set_HOST_BILLING_DATE(ByVal key As String)
    Try
      gHOST_BILLING_DATE = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gHOST_CREATE_TIME
  Private Sub set_HOST_CREATE_TIME(ByVal key As String)
    Try
      gHOST_CREATE_TIME = key
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
  '-設定gobjPO
  Private Sub set_objPO(ByVal objPO As clsPO_Back)
    Try
      gobjPO = objPO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub


  '非標準的Function
  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef objPO_DTL As clsPO_DTL_Bak) As Boolean
    Try
      Dim key As String = objPO_DTL.get_gid()
      If key <> get_gid() Then
        SendMessageToLog("Key can not Update, old_Key=" & get_gid() & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      set_PO_ID(objPO_DTL.get_PO_ID)
      set_PO_SERIAL_NO(objPO_DTL.get_PO_SERIAL_NO)
      set_SKU_NO(objPO_DTL.get_SKU_NO)
      set_LOT_NO(objPO_DTL.get_LOT_NO)
      set_QTY(objPO_DTL.get_QTY)
      set_QTY_PROCESS(objPO_DTL.get_QTY_PROCESS)
      set_QTY_FINISH(objPO_DTL.get_QTY_FINISH)
      set_COMMENTS(objPO_DTL.get_COMMENTS)
      set_PACKAGE_ID(objPO_DTL.get_PACKAGE_ID)
      set_ITEM_COMMON1(objPO_DTL.get_ITEM_COMMON1)
      set_ITEM_COMMON2(objPO_DTL.get_ITEM_COMMON2)
      set_ITEM_COMMON3(objPO_DTL.get_ITEM_COMMON3)
      set_ITEM_COMMON4(objPO_DTL.get_ITEM_COMMON4)
      set_ITEM_COMMON5(objPO_DTL.get_ITEM_COMMON5)
      set_ITEM_COMMON6(objPO_DTL.get_ITEM_COMMON6)
      set_ITEM_COMMON7(objPO_DTL.get_ITEM_COMMON7)
      set_ITEM_COMMON8(objPO_DTL.get_ITEM_COMMON8)
      set_ITEM_COMMON9(objPO_DTL.get_ITEM_COMMON9)
      set_ITEM_COMMON10(objPO_DTL.get_ITEM_COMMON10)


      set_HOST_FACTORY_NO(objPO_DTL.get_HOST_FACTORY_NO)
      set_HOST_AREA_NO(objPO_DTL.get_HOST_AREA_NO)
      'set_HOST_PO_TYPE(objPO_DTL.get_HOST_PO_TYPE)
      set_HOST_OWNER_NO(objPO_DTL.get_HOST_OWNER_NO)
      set_HOST_COMMON1(objPO_DTL.get_HOST_COMMON1)
      set_HOST_COMMON2(objPO_DTL.get_HOST_COMMON2)
      set_HOST_COMMON3(objPO_DTL.get_HOST_COMMON3)
      set_HOST_COMMON4(objPO_DTL.get_HOST_COMMON4)
      set_HOST_COMMON5(objPO_DTL.get_HOST_COMMON5)
      set_HOST_COMMON6(objPO_DTL.get_HOST_COMMON6)
      set_HOST_COMMON7(objPO_DTL.get_HOST_COMMON7)
      set_HOST_COMMON8(objPO_DTL.get_HOST_COMMON8)
      set_HOST_COMMON9(objPO_DTL.get_HOST_COMMON9)
      set_HOST_COMMON10(objPO_DTL.get_HOST_COMMON10)
      set_HOST_COMMON11(objPO_DTL.get_HOST_COMMON11)
      set_HOST_COMMON12(objPO_DTL.get_HOST_COMMON12)
      set_HOST_COMMON13(objPO_DTL.get_HOST_COMMON13)
      set_HOST_COMMON14(objPO_DTL.get_HOST_COMMON14)
      set_HOST_COMMON15(objPO_DTL.get_HOST_COMMON15)
      set_HOST_COMMON16(objPO_DTL.get_HOST_COMMON16)
      set_HOST_COMMON17(objPO_DTL.get_HOST_COMMON17)
      set_HOST_COMMON18(objPO_DTL.get_HOST_COMMON18)
      set_HOST_COMMON19(objPO_DTL.get_HOST_COMMON19)
      set_HOST_COMMON20(objPO_DTL.get_HOST_COMMON20)
      set_HOST_COMMON21(objPO_DTL.get_HOST_COMMON21)
      set_HOST_COMMON22(objPO_DTL.get_HOST_COMMON22)
      set_HOST_COMMON23(objPO_DTL.get_HOST_COMMON23)
      set_HOST_COMMON24(objPO_DTL.get_HOST_COMMON24)
      set_HOST_COMMON25(objPO_DTL.get_HOST_COMMON25)
      set_HOST_COMMON26(objPO_DTL.get_HOST_COMMON26)
      set_HOST_COMMON27(objPO_DTL.get_HOST_COMMON27)
      set_HOST_COMMON28(objPO_DTL.get_HOST_COMMON28)
      set_HOST_COMMON29(objPO_DTL.get_HOST_COMMON29)
      set_HOST_COMMON30(objPO_DTL.get_HOST_COMMON30)
      set_HOST_COMMON31(objPO_DTL.get_HOST_COMMON31)
      set_HOST_COMMON32(objPO_DTL.get_HOST_COMMON32)
      set_HOST_COMMON33(objPO_DTL.get_HOST_COMMON33)
      set_HOST_COMMON34(objPO_DTL.get_HOST_COMMON34)
      set_HOST_COMMON35(objPO_DTL.get_HOST_COMMON35)
      set_HOST_COMMON36(objPO_DTL.get_HOST_COMMON36)
      set_HOST_COMMON37(objPO_DTL.get_HOST_COMMON37)
      set_HOST_COMMON38(objPO_DTL.get_HOST_COMMON38)
      set_HOST_COMMON39(objPO_DTL.get_HOST_COMMON39)
      set_HOST_COMMON40(objPO_DTL.get_HOST_COMMON40)
      set_HOST_STEP_NO(objPO_DTL.get_HOST_STEP_NO)
      set_HOST_MOVE_TYPE(objPO_DTL.get_HOST_MOVE_TYPE)
      set_HOST_FINISH_TIME(objPO_DTL.get_HOST_FINISH_TIME)
      set_HOST_BILLING_DATE(objPO_DTL.get_HOST_BILLING_DATE)
      set_HOST_CREATE_TIME(objPO_DTL.get_HOST_CREATE_TIME)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try

  End Function

  Public Function Update_To_Memory_ByModule_MOCAB01_SENDWORKDATA(ByVal TA006 As String, ByVal TA063 As String, ByVal TA015 As String, ByVal TA007 As String, ByVal TA011 As String, ByVal TA013 As String, ByVal TA021 As String, ByVal TA030 As String) As Boolean
    Try

      set_SKU_NO(TA006)
      set_LOT_NO(TA063)
      set_QTY(TA015)
      set_HOST_COMMON1(TA007)
      set_HOST_COMMON2(TA011)
      set_HOST_COMMON3(TA013)
      set_HOST_COMMON4(TA021)
      set_HOST_COMMON5(TA030)
      set_HOST_COMMON6(TA063)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try

  End Function
  Public Function Update_To_Memory_ByMOCAB04_SENDWORKIDCHANGEDATA(ByVal TO003 As String, ByVal TO004 As String, ByVal TO017 As String, ByVal TO023 As String, ByVal TO034 As String, ByVal TO041 As String) As Boolean
    Try

      set_LOT_NO(TO034)
      set_QTY(TO017)
      set_HOST_COMMON3(TO041)
      set_HOST_COMMON4(TO023)
      set_HOST_COMMON6(TO034)
      set_HOST_COMMON7(TO003)
      set_HOST_COMMON8(TO004)


      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try

  End Function
  '新增資料到DB
  Public Function O_Insert_PODTL_To_DB(ByRef objWMS As clsHandlingObject) As Boolean
    Try
      '一定要寫成功，才更新記憶體的狀態			
      If WMS_T_PO_DTLManagement_BackUp.AddWMS_T_PO_DTLData(Me) = True Then
        '建立梆定
        Add_Relationship(objWMS)
        Return True
      Else
        SendMessageToLog("Insert PODTL to DB Failed ,TableName = " & WMS_T_PO_DTLManagement_BackUp.TableName & " ,key=" & get_gid(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '刪除
  Public Function O_Delete_PODTL_To_DB() As Boolean
    Try
      '一定要寫成功，才更新記憶體的狀態			
      If WMS_T_PO_DTLManagement_BackUp.DeleteWMS_T_PO_DTLData(Me) = True Then
        '解除梆定
        Remove_Relationship()
        Return True
      Else
        SendMessageToLog("Delete PODTL to DB Failed ,TableName = " & WMS_T_PO_DTLManagement_BackUp.TableName & " ,key=" & get_gid(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '更新
  Public Function O_Update_SKUID_To_DB(ByVal SKUID As String) As Boolean
    Try
      '一定要寫成功，才更新記憶體的狀態
      Dim Update_Time As String = GetNewTime_ByDataTimeFormat(DBTimeFormat)
      Dim New_objProcessFlow As clsPO_DTL_Bak = Clone()
      New_objProcessFlow.set_SKU_NO(SKUID)

      '在資料庫Update 的資料
      If WMS_T_PO_DTLManagement_BackUp.UpdateWMS_T_PO_DTLData(New_objProcessFlow) = True Then
        set_SKU_NO(SKUID)
        Return True
      Else
        SendMessageToLog("Update SKUID to DB Failed ,TableName = " & WMS_T_PO_DTLManagement_BackUp.TableName & " ,key=" & get_gid(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '更新Qty_Process
  Public Function O_Update_Qty_Process(ByVal Qty_Process As Double) As Boolean
    Try
      set_QTY_PROCESS(Qty_Process)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '更新Qty_Finish
  Public Function O_Update_Qty_Finish(ByVal Qty_Finish As Double) As Boolean
    Try
      set_QTY_FINISH(Qty_Finish)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
