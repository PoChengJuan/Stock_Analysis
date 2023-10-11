Imports System.Collections.Concurrent
Public Class clsPO_MERGE
  Private ShareName As String = "PO_MERGE"
  Private ShareKey As String = ""
  Private _gid As String
  Private _PO_ID As String '訂單編號
  Private _PO_SERIAL_NO As String '訂單明細編號
  Private _WO_ID As String '工單編號
  Private _WO_SERIAL_NO As String '工單明細編號
  Private _QTY As Decimal '需求量
  Private _QTY_PROCESS As Decimal '已挷定數量(已經配置了棧板上料品的數量)
  Private _QTY_FINISH As Decimal '已完成數量(已經完成入出庫的數量)
  Private _CLOSE_UUID As String '結單時WMS發送給HostHandler的UUID(不用)
  Private _COMMENTS As String '備註
  Private _CREATE_TIME As String '建立時間
  Public objWMS As clsHandlingObject
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
  Public Property PO_SERIAL_NO() As String
    Get
      Return _PO_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _PO_SERIAL_NO = value
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
  Public Property WO_SERIAL_NO() As String
    Get
      Return _WO_SERIAL_NO
    End Get
    Set(ByVal value As String)
      _WO_SERIAL_NO = value
    End Set
  End Property
  Public Property QTY() As Decimal
    Get
      Return _QTY
    End Get
    Set(ByVal value As Decimal)
      _QTY = value
    End Set
  End Property
  Public Property QTY_PROCESS() As Decimal
    Get
      Return _QTY_PROCESS
    End Get
    Set(ByVal value As Decimal)
      _QTY_PROCESS = value
    End Set
  End Property
  Public Property QTY_FINISH() As Decimal
    Get
      Return _QTY_FINISH
    End Get
    Set(ByVal value As Decimal)
      _QTY_FINISH = value
    End Set
  End Property
  Public Property CLOSE_UUID() As String
    Get
      Return _CLOSE_UUID
    End Get
    Set(ByVal value As String)
      _CLOSE_UUID = value
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

  Public Sub New(ByVal PO_ID As String, ByVal PO_SERIAL_NO As String, ByVal WO_ID As String, ByVal WO_SERIAL_NO As String, ByVal QTY As Decimal, ByVal QTY_PROCESS As Decimal, ByVal QTY_FINISH As Decimal, ByVal CLOSE_UUID As String, ByVal COMMENTS As String, ByVal CREATE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PO_ID, PO_SERIAL_NO, WO_ID, WO_SERIAL_NO)
      _gid = key
      _PO_ID = PO_ID
      _PO_SERIAL_NO = PO_SERIAL_NO
      _WO_ID = WO_ID
      _WO_SERIAL_NO = WO_SERIAL_NO
      _QTY = QTY
      _QTY_PROCESS = QTY_PROCESS
      _QTY_FINISH = QTY_FINISH
      _CLOSE_UUID = CLOSE_UUID
      _COMMENTS = COMMENTS
      _CREATE_TIME = CREATE_TIME
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
  Public Shared Function Get_Combination_Key(ByVal PO_ID As String, ByVal PO_SERIAL_NO As String, ByVal WO_ID As String, ByVal WO_SERIAL_NO As String) As String
    Try
      Dim key As String = PO_ID & LinkKey & PO_SERIAL_NO & LinkKey & WO_ID & LinkKey & WO_SERIAL_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPO_MERGE
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '新增記憶體內容
  'Public Sub Add_Relationship(ByRef objWMS As clsHandlingObject)
  '  Try
  '    '綁定PO_MERGE和WMS的關係
  '    If objWMS IsNot Nothing Then
  '      Me.objWMS = objWMS
  '      '此處如有更改，須自行修改
  '      objWMS.O_Add_PO_MERGE(Me)
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub
  ''移除記憶體內容
  'Public Sub Remove_Relationship()
  '  Try
  '    If Me.objWMS IsNot Nothing Then
  '      Me.objWMS.O_Remove_PO_MERGE(Me)
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_PO_MERGEManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_PO_MERGEManagement.GetUpdateSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  'Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
  '  Try
  '    Dim objPO_MERGE As clsPO_MERGE = Nothing
  '    Dim dicChangeColumnValue As New Dictionary(Of String, String)
  '    If O_Get_UpdateColumnValue(objPO_MERGE, Me, dicChangeColumnValue) = True Then
  '      Dim strSQL As String = WMS_T_PO_MERGEManagement.GetUpdateSQLForChangeValue(Me, dicChangeColumnValue)
  '      If strSQL <> "" Then
  '        lstSQL.Add(strSQL)
  '      End If
  '    Else
  '      SendMessageToLog("O_Get_UpdateColumnValue Faled", eCALogTool.ILogTool.enuTrcLevel.lvError)
  '      '失敗先用原來的方式
  '      Dim strSQL As String = WMS_T_PO_MERGEManagement.GetUpdateSQL(Me)
  '      lstSQL.Add(strSQL)
  '    End If
  '    Return True
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return False
  '  End Try
  'End Function
  '取得要Delete的SQL
  Public Function O_Add_Delete_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_PO_MERGEManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_T_PO_MERGE As clsPO_MERGE) As Boolean
    Try
      Dim key As String = objWMS_T_PO_MERGE._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _PO_ID = objWMS_T_PO_MERGE.PO_ID
      _PO_SERIAL_NO = objWMS_T_PO_MERGE.PO_SERIAL_NO
      _WO_ID = objWMS_T_PO_MERGE.WO_ID
      _WO_SERIAL_NO = objWMS_T_PO_MERGE.WO_SERIAL_NO
      _QTY = objWMS_T_PO_MERGE.QTY
      _QTY_PROCESS = objWMS_T_PO_MERGE.QTY_PROCESS
      _QTY_FINISH = objWMS_T_PO_MERGE.QTY_FINISH
      _CLOSE_UUID = objWMS_T_PO_MERGE.CLOSE_UUID
      _COMMENTS = objWMS_T_PO_MERGE.COMMENTS
      _CREATE_TIME = objWMS_T_PO_MERGE.CREATE_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
