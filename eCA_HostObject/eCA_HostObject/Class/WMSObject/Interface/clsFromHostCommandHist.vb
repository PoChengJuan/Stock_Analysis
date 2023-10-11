Public Class clsFromHostCommandHist
  Private ShareName As String = "FromHostCommandHist"
  Private ShareKey As String = ""

  Private _UUID As String
  Private _Send_System As enuSystemType
  Private _Receive_System As enuSystemType
  Private _Function_ID As String
  Private _SEQ As Long
  Private _User_ID As String
  Private _Create_Time As String
  Private _Message As String
  Private _Result As String
  Private _Result_Message As String
  Private _Wait_UUID As String
  Private _Hist_Time As String

  Public Property UUID() As String
    Get
      Return _UUID
    End Get
    Set(ByVal value As String)
      _UUID = value
    End Set
  End Property
  Public Property Send_System() As enuSystemType
    Get
      Return _Send_System
    End Get
    Set(ByVal value As enuSystemType)
      _Send_System = value
    End Set
  End Property
  Public Property Receive_System() As enuSystemType
    Get
      Return _Receive_System
    End Get
    Set(ByVal value As enuSystemType)
      _Receive_System = value
    End Set
  End Property
  Public Property Function_ID() As String
    Get
      Return _Function_ID
    End Get
    Set(ByVal value As String)
      _Function_ID = value
    End Set
  End Property
  Public Property SEQ() As Long
    Get
      Return _SEQ
    End Get
    Set(ByVal value As Long)
      _SEQ = value
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
  Public Property Create_Time() As String
    Get
      Return _Create_Time
    End Get
    Set(ByVal value As String)
      _Create_Time = value
    End Set
  End Property
  Public Property Message() As String
    Get
      Return _Message
    End Get
    Set(ByVal value As String)
      _Message = value
    End Set
  End Property
  Public Property Result() As String
    Get
      Return _Result
    End Get
    Set(ByVal value As String)
      _Result = value
    End Set
  End Property
  Public Property Result_Message() As String
    Get
      Return _Result_Message
    End Get
    Set(ByVal value As String)
      _Result_Message = value
    End Set
  End Property
  Public Property Wait_UUID() As String
    Get
      Return _Wait_UUID
    End Get
    Set(ByVal value As String)
      _Wait_UUID = value
    End Set
  End Property
  Public Property Hist_Time() As String
    Get
      Return _Hist_Time
    End Get
    Set(ByVal value As String)
      _Hist_Time = value
    End Set
  End Property

  Public Sub New(ByVal UUID As String,
                 ByVal Send_System As enuSystemType,
                 ByVal Receive_System As enuSystemType,
                 ByVal Function_ID As String,
                 ByVal SEQ As Long,
                 ByVal User_ID As String,
                 ByVal Create_Time As String,
                 ByVal Message As String,
                 ByVal Result As String,
                 ByVal Result_Message As String,
                 ByVal Wait_UUID As String,
                 ByVal Hist_Time As String)
    MyBase.New()
    Try
      _UUID = UUID
      _Send_System = Send_System
      _Receive_System = Receive_System
      _Function_ID = Function_ID
      _SEQ = SEQ
      _User_ID = User_ID
      _Create_Time = Create_Time
      _Message = Message
      _Result = Result
      _Result_Message = Result_Message
      _Wait_UUID = Wait_UUID
      _Hist_Time = Hist_Time
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
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = HOST_H_Command_HistManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
