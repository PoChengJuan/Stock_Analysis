﻿Public Class clsClassProduction_HIST
  Private ShareName As String = "ClassProduction_HIST"
  Private ShareKey As String = ""
  Private _gid As String
  Private _FACTORY_NO As String '廠別

  Private _AREA_NO As String '區域編號

  Private _DEVICE_NO As String 'WCS編號

  Private _UNIT_ID As String '設備編號

  Private _CLASS_NO As String '班別編號

  Private _QTY_TOTAL As Double '總生產數量

  Private _QTY_PROCESS As Double '目前生產數量

  Private _QTY_MODIFY As Double '人員手動調整的數量

  Private _QTY_NG As Double 'NG的數量

  Private _HIST_TIME As String '寫入歷史時間

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property FACTORY_NO() As String
    Get
      Return _FACTORY_NO
    End Get
    Set(ByVal value As String)
      _FACTORY_NO = value
    End Set
  End Property
  Public Property AREA_NO() As String
    Get
      Return _AREA_NO
    End Get
    Set(ByVal value As String)
      _AREA_NO = value
    End Set
  End Property
  Public Property DEVICE_NO() As String
    Get
      Return _DEVICE_NO
    End Get
    Set(ByVal value As String)
      _DEVICE_NO = value
    End Set
  End Property
  Public Property UNIT_ID() As String
    Get
      Return _UNIT_ID
    End Get
    Set(ByVal value As String)
      _UNIT_ID = value
    End Set
  End Property
  Public Property CLASS_NO() As String
    Get
      Return _CLASS_NO
    End Get
    Set(ByVal value As String)
      _CLASS_NO = value
    End Set
  End Property
  Public Property QTY_TOTAL() As Double
    Get
      Return _QTY_TOTAL
    End Get
    Set(ByVal value As Double)
      _QTY_TOTAL = value
    End Set
  End Property
  Public Property QTY_PROCESS() As Double
    Get
      Return _QTY_PROCESS
    End Get
    Set(ByVal value As Double)
      _QTY_PROCESS = value
    End Set
  End Property
  Public Property QTY_MODIFY() As Double
    Get
      Return _QTY_MODIFY
    End Get
    Set(ByVal value As Double)
      _QTY_MODIFY = value
    End Set
  End Property
  Public Property QTY_NG() As Double
    Get
      Return _QTY_NG
    End Get
    Set(ByVal value As Double)
      _QTY_NG = value
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

  Public Sub New(ByVal FACTORY_NO As String,
                 ByVal AREA_NO As String,
                 ByVal DEVICE_NO As String,
                 ByVal UNIT_ID As String,
                 ByVal CLASS_NO As String,
                 ByVal QTY_TOTAL As Double,
                 ByVal QTY_PROCESS As Double,
                 ByVal QTY_MODIFY As Double,
                 ByVal QTY_NG As Double,
                 ByVal HIST_TIME As String)
    MyBase.New()
    Try

      _FACTORY_NO = FACTORY_NO
      _AREA_NO = AREA_NO
      _DEVICE_NO = DEVICE_NO
      _UNIT_ID = UNIT_ID
      _CLASS_NO = CLASS_NO
      _QTY_TOTAL = QTY_TOTAL
      _QTY_PROCESS = QTY_PROCESS
      _QTY_MODIFY = QTY_MODIFY
      _QTY_NG = QTY_NG
      _HIST_TIME = HIST_TIME
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
  Public Function Clone() As clsClassProduction_HIST
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
      Dim strSQL As String = WMS_CH_CLASS_PRODUCTION_HISTManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Class
