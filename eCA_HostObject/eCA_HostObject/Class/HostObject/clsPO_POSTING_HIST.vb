﻿Public Class clsPO_POSTING_HIST
  Private ShareName As String = "WMS_T_PO_POSTING_HIST"
  Private ShareKey As String = ""
  Private _gid As String
  Private _PO_ID As String '訂單編號

  Private _PO_LINE_NO As String '訂單明細編號(上傳時使用)

  Private _WO_ID As String '工单编号

  Private _SORT_ITEM_COMMON1 As String '寄售库存标记

  Private _SORT_ITEM_COMMON2 As String '基本计量单位

  Private _SORT_ITEM_COMMON3 As String

  Private _SORT_ITEM_COMMON4 As String

  Private _SORT_ITEM_COMMON5 As String

  Private _QTY As Double '此次过账数量

  Private _UUID As String '过账命令编号

  Private _CREATE_TIME As String '建立时间

  Private _UPDATE_TIME As String '更新时间

  Private _RESULT As Double '0=过账成功, 1=过帐失败

  Private _RESULT_MESSAGE As String '过帐失败时填入原因

  Private _H_POP1 As String '

  Private _H_POP2 As String '

  Private _H_POP3 As String '

  Private _H_POP4 As String '

  Private _H_POP5 As String '

  Private _HIST_TIME As String

  Private _SKU_NO As String

  Private _CLOSE_USER_ID As String

  Private _START_TRANSFER_TIME As String

  Private _FINISH_TRANSFER_TIME As String

  Private _ORDER_TYPE As enuOrderType

  Private _PO_SERIAL_NO As String

  Private _TKNUM As String

  Private _LOT_NO As String

  Private _OWNER As String

  Private _SUBOWNER As String

  Public Property SKU_NO() As String
    Get
      Return _SKU_NO
    End Get
    Set(ByVal value As String)
      _SKU_NO = value
    End Set
  End Property
  Public Property HIST_TIME() As String
    Get
      Return _HIST_TIME
    End Get
    Set(ByVal value As String)
      _HIST_TIME = value
    End Set
  End Property

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
  Public Property PO_LINE_NO() As String
    Get
      Return _PO_LINE_NO
    End Get
    Set(ByVal value As String)
      _PO_LINE_NO = value
    End Set
  End Property
  Public Property WO_ID() As String
    Get
      Return _WO_ID
    End Get
    Set(ByVal value As String)
      _WO_ID = value
    End Set
  End Property
  Public Property SORT_ITEM_COMMON1() As String
    Get
      Return _SORT_ITEM_COMMON1
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON1 = value
    End Set
  End Property
  Public Property SORT_ITEM_COMMON2() As String
    Get
      Return _SORT_ITEM_COMMON2
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON2 = value
    End Set
  End Property

  Public Property SORT_ITEM_COMMON5() As String
    Get
      Return _SORT_ITEM_COMMON5
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON5 = value
    End Set
  End Property

  Public Property SORT_ITEM_COMMON4() As String
    Get
      Return _SORT_ITEM_COMMON4
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON4 = value
    End Set
  End Property

  Public Property SORT_ITEM_COMMON3() As String
    Get
      Return _SORT_ITEM_COMMON3
    End Get
    Set(ByVal value As String)
      _SORT_ITEM_COMMON3 = value
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
  Public Property UUID() As String
    Get
      Return _UUID
    End Get
    Set(ByVal value As String)
      _UUID = value
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
  Public Property RESULT() As Double
    Get
      Return _RESULT
    End Get
    Set(ByVal value As Double)
      _RESULT = value
    End Set
  End Property
  Public Property RESULT_MESSAGE() As String
    Get
      Return _RESULT_MESSAGE
    End Get
    Set(ByVal value As String)
      _RESULT_MESSAGE = value
    End Set
  End Property
  Public Property H_POP1() As String
    Get
      Return _H_POP1
    End Get
    Set(ByVal value As String)
      _H_POP1 = value
    End Set
  End Property
  Public Property H_POP2() As String
    Get
      Return _H_POP2
    End Get
    Set(ByVal value As String)
      _H_POP2 = value
    End Set
  End Property
  Public Property H_POP3() As String
    Get
      Return _H_POP3
    End Get
    Set(ByVal value As String)
      _H_POP3 = value
    End Set
  End Property
  Public Property H_POP4() As String
    Get
      Return _H_POP4
    End Get
    Set(ByVal value As String)
      _H_POP4 = value
    End Set
  End Property
  Public Property H_POP5() As String
    Get
      Return _H_POP5
    End Get
    Set(ByVal value As String)
      _H_POP5 = value
    End Set
  End Property
  Public Property CLOSE_USER_ID() As String
    Get
      Return _CLOSE_USER_ID
    End Get
    Set(ByVal value As String)
      _CLOSE_USER_ID = value
    End Set
  End Property
  Public Property START_TRANSFER_TIME() As String
    Get
      Return _START_TRANSFER_TIME
    End Get
    Set(ByVal value As String)
      _START_TRANSFER_TIME = value
    End Set
  End Property
  Public Property FINISH_TRANSFER_TIME() As String
    Get
      Return _FINISH_TRANSFER_TIME
    End Get
    Set(ByVal value As String)
      _FINISH_TRANSFER_TIME = value
    End Set
  End Property
  Public Property ORDER_TYPE() As enuOrderType
    Get
      Return _ORDER_TYPE
    End Get
    Set(ByVal value As enuOrderType)
      _ORDER_TYPE = value
    End Set
  End Property
  Public Property PO_SERIAL_NO() As String
    Get
      Return _PO_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _PO_SERIAL_NO = value
    End Set
  End Property
  Public Property TKNUM() As String
    Get
      Return _TKNUM
    End Get
    Set(ByVal value As String)
      _TKNUM = value
    End Set
  End Property
  Public Property LOT_NO() As String
    Get
      Return _LOT_NO
    End Get
    Set(ByVal value As String)
      _LOT_NO = value
    End Set
  End Property
  Public Property OWNER() As String
    Get
      Return _OWNER
    End Get
    Set(ByVal value As String)
      _OWNER = value
    End Set
  End Property
  Public Property SUBOWNER() As String
    Get
      Return _SUBOWNER
    End Get
    Set(ByVal value As String)
      _SUBOWNER = value
    End Set
  End Property

  Public Sub New(ByVal PO_ID As String, ByVal PO_LINE_NO As String, ByVal WO_ID As String, ByVal SORT_ITEM_COMMON1 As String, ByVal SORT_ITEM_COMMON2 As String,
               ByVal SORT_ITEM_COMMON3 As String, ByVal SORT_ITEM_COMMON4 As String, ByVal SORT_ITEM_COMMON5 As String, ByVal QTY As Double, ByVal UUID As String,
               ByVal CREATE_TIME As String, ByVal UPDATE_TIME As String, ByVal RESULT As Double, ByVal RESULT_MESSAGE As String, ByVal H_POP1 As String, ByVal H_POP2 As String, ByVal H_POP3 As String,
               ByVal H_POP4 As String, ByVal H_POP5 As String, ByVal HIST_TIME As String, ByVal SKU_NO As String, ByVal CLOSE_USER_ID As String, ByVal START_TRANSFER_TIME As String, ByVal FINISH_TRANSFER_TIME As String,
               ByVal ORDER_TYPE As enuOrderType, ByVal PO_SERIAL_NO As String, ByVal TKNUM As String, ByVal LOT_NO As String, ByVal OWNER As String, ByVal SUBOWNER As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key()
      _gid = key
      _TKNUM = TKNUM
      _PO_ID = PO_ID
      _PO_LINE_NO = PO_LINE_NO
      _PO_SERIAL_NO = PO_SERIAL_NO
      _WO_ID = WO_ID
      _SKU_NO = SKU_NO
      _SORT_ITEM_COMMON1 = SORT_ITEM_COMMON1
      _SORT_ITEM_COMMON2 = SORT_ITEM_COMMON2
      _SORT_ITEM_COMMON3 = SORT_ITEM_COMMON3
      _SORT_ITEM_COMMON4 = SORT_ITEM_COMMON4
      _SORT_ITEM_COMMON5 = SORT_ITEM_COMMON5
      _QTY = QTY
      _UUID = UUID
      _CREATE_TIME = CREATE_TIME
      _UPDATE_TIME = UPDATE_TIME
      _RESULT = RESULT
      _RESULT_MESSAGE = RESULT_MESSAGE
      _H_POP1 = H_POP1
      _H_POP2 = H_POP2
      _H_POP3 = H_POP3
      _H_POP4 = H_POP4
      _H_POP5 = H_POP5
      _HIST_TIME = HIST_TIME
      _CLOSE_USER_ID = CLOSE_USER_ID
      _START_TRANSFER_TIME = START_TRANSFER_TIME
      _FINISH_TRANSFER_TIME = FINISH_TRANSFER_TIME
      _ORDER_TYPE = ORDER_TYPE
      _LOT_NO = LOT_NO
      _OWNER = OWNER
      _SUBOWNER = SUBOWNER
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
  Public Shared Function Get_Combination_Key() '(ByVal PO_ID As String, ByVal PO_LINE_NO As String, ByVal WO_ID As String, ByVal SORT_ITEM_COMMON1 As String, ByVal SORT_ITEM_COMMON2 As String, ByVal QTY As Double) As String
    Try
      Dim key As String = "" ' PO_ID & "_" & PO_LINE_NO & "_" & WO_ID & "_" & SORT_ITEM_COMMON1 & "_" & SORT_ITEM_COMMON2 & "_" & QTY.ToString
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPO_POSTING
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_H_PO_POSTING_HISTManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
