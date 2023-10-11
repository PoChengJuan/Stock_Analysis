Imports eCA_HostObject

''' <summary>
''' 20181119
''' V1.0.0
''' Mark
''' 處理寫入歷史資料的部份
''' </summary>
Public Class Module_Process_History
  ''' <summary>
  ''' 20181119
  ''' Mark
  ''' 把Line變更的Status寫入History
  ''' </summary>
  ''' <param name="ret_lstQueueQSL"></param>
  ''' <param name="dicUpdateLine"></param>
  Public Shared Sub Get_LineStatus_HIST_SQL(ByRef ret_lstQueueQSL As List(Of String),
                                            ByRef dicUpdateLine As Dictionary(Of String, clsLine_Status))
    Try
      Dim Now_Time As String = GetNewTime_DBFormat()
      For Each objNewLine As clsLine_Status In dicUpdateLine.Values
        '取得原來的LineStatus
        Dim Factory_No As String = objNewLine.Factory_No
        Dim Area_No As String = objNewLine.Area_No
        Dim Device_No As String = objNewLine.Device_No
        Dim Unit_ID As String = objNewLine.Unit_ID
        Dim To_Status As enuLineStatus = objNewLine.Status
        Dim From_Status As enuLineStatus = enuLineStatus.None
        Dim objLine As clsLine_Status = Nothing
        If gMain.objHandling.O_Get_CLine(Factory_No, Area_No, Device_No, Unit_ID, objLine) = True Then
          From_Status = objLine.Status
        End If
        '組成Line_Hist
        Dim objLineHist As New clsLine_Status_Hist(Factory_No, Area_No, Device_No, Unit_ID, From_Status, To_Status, Now_Time)
        objLineHist.O_Add_Insert_SQLString(ret_lstQueueQSL)
      Next
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  ''' <summary>
  ''' 20181119
  ''' Mark
  ''' 把LineInfo移除的寫入History
  ''' </summary>
  ''' <param name="ret_lstQueueQSL"></param>
  ''' <param name="dicDeleteLineIfno"></param>
  Public Shared Sub Get_LineInfo_HIST_SQL(ByRef ret_lstQueueQSL As List(Of String),
                                          ByVal Remove_User As String,
                                          ByRef dicDeleteLineIfno As Dictionary(Of String, clsLineInfo))
    Try
      Dim Now_Time As String = GetNewTime_DBFormat()
      For Each objDeleteLineInfo As clsLineInfo In dicDeleteLineIfno.Values
        '取得原來的LineStatus
        Dim Factory_No As String = objDeleteLineInfo.Factory_No
        Dim Area_No As String = objDeleteLineInfo.Area_No
        Dim Device_No As String = objDeleteLineInfo.Device_No
        Dim Unit_ID As String = objDeleteLineInfo.Unit_ID
        Dim Occur_Time As String = objDeleteLineInfo.Occur_Time
				Dim Message As String = objDeleteLineInfo.Maintenance_Message

				Dim MAINTENANCE_ID As String = objDeleteLineInfo.MAINTENANCE_ID
				Dim FUCTION_ID As String = objDeleteLineInfo.FUCTION_ID
				Dim OPERATOR_USER As String = Remove_User
				Dim COMMENTS As String = ""
				'組成Line_Hist
				Dim objLineHist As New clsLineInfo_Hist(Factory_No, Area_No, Device_No, Unit_ID, Occur_Time, Message, Remove_User, Now_Time, MAINTENANCE_ID,
																								FUCTION_ID, OPERATOR_USER, COMMENTS)
				objLineHist.O_Add_Insert_SQLString(ret_lstQueueQSL)
      Next
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub


    Public Shared Sub Get_Account_HIST_SQL(ByRef ret_lstQueueQSL As List(Of String),
                                            ByRef dicDeleteAccount As Dictionary(Of String, clsWMS_CT_ACCOUNT_REPORT))
        Try
            Dim Now_Time As String = GetNewTime_DBFormat()
            For Each objDeleteAccount In dicDeleteAccount.Values
                '取得原來的LineStatus
                Dim SKU_NO = objDeleteAccount.SKU_NO
                Dim LOT_NO = objDeleteAccount.LOT_NO
                Dim ITEM_COMMON1 = objDeleteAccount.ITEM_COMMON1
                Dim ITEM_COMMON2 = objDeleteAccount.ITEM_COMMON2
                Dim ITEM_COMMON3 = objDeleteAccount.ITEM_COMMON3
                Dim ITEM_COMMON4 = objDeleteAccount.ITEM_COMMON4
                Dim ITEM_COMMON5 = objDeleteAccount.ITEM_COMMON5
                Dim ITEM_COMMON6 = objDeleteAccount.ITEM_COMMON6
                Dim ITEM_COMMON7 = objDeleteAccount.ITEM_COMMON7
                Dim ITEM_COMMON8 = objDeleteAccount.ITEM_COMMON8
                Dim ITEM_COMMON9 = objDeleteAccount.ITEM_COMMON9
                Dim ITEM_COMMON10 = objDeleteAccount.ITEM_COMMON10
                Dim SORT_ITEM_COMMON1 = objDeleteAccount.SORT_ITEM_COMMON1
                Dim SORT_ITEM_COMMON2 = objDeleteAccount.SORT_ITEM_COMMON2
                Dim SORT_ITEM_COMMON3 = objDeleteAccount.SORT_ITEM_COMMON3
                Dim SORT_ITEM_COMMON4 = objDeleteAccount.SORT_ITEM_COMMON4
                Dim SORT_ITEM_COMMON5 = objDeleteAccount.SORT_ITEM_COMMON5
                Dim OWNER_NO = objDeleteAccount.OWNER_NO
                Dim SUB_OWNER_NO = objDeleteAccount.SUB_OWNER_NO
                Dim STORAGE_TYPE = objDeleteAccount.STORAGE_TYPE
                Dim BND = objDeleteAccount.BND
                Dim QC_STATUS = objDeleteAccount.QC_STATUS
                Dim WMS_STOCK_QTY = objDeleteAccount.WMS_STOCK_QTY
                Dim ERP_SYSTEM = objDeleteAccount.ERP_SYSTEM
                Dim ERP_STOCK_QTY = objDeleteAccount.ERP_STOCK_QTY
                Dim QUANTITY_VARIANCE = objDeleteAccount.QUANTITY_VARIANCE
                Dim CREATE_TIME = objDeleteAccount.CREATE_TIME
                Dim ACC_COMMON1 = objDeleteAccount.ACC_COMMON1
                Dim ACC_COMMON2 = objDeleteAccount.ACC_COMMON2
                Dim ACC_COMMON3 = objDeleteAccount.ACC_COMMON3
                Dim ACC_COMMON4 = objDeleteAccount.ACC_COMMON4
                Dim ACC_COMMON5 = objDeleteAccount.ACC_COMMON5
                Dim HIST_TIME = Now_Time
                '組成Line_Hist
                Dim objLineHist As New clsWMS_CH_ACCOUNT_REPORT(SKU_NO, LOT_NO, ITEM_COMMON1, ITEM_COMMON2, ITEM_COMMON3, ITEM_COMMON4, ITEM_COMMON5, ITEM_COMMON6, ITEM_COMMON7, ITEM_COMMON8, ITEM_COMMON9, ITEM_COMMON10, SORT_ITEM_COMMON1, SORT_ITEM_COMMON2, SORT_ITEM_COMMON3, SORT_ITEM_COMMON4, SORT_ITEM_COMMON5, OWNER_NO, SUB_OWNER_NO, STORAGE_TYPE, BND, QC_STATUS, WMS_STOCK_QTY, ERP_SYSTEM, ERP_STOCK_QTY, QUANTITY_VARIANCE, CREATE_TIME, ACC_COMMON1, ACC_COMMON2, ACC_COMMON3, ACC_COMMON4, ACC_COMMON5, HIST_TIME)
                objLineHist.O_Add_Insert_SQLString(ret_lstQueueQSL)
            Next
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End Try
    End Sub

End Class
