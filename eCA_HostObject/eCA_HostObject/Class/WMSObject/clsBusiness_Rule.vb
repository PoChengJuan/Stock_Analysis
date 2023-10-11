Public Class clsBusiness_Rule

  Private ShareName As String = "Business_Rule"
  Private ShareKey As String = ""

  Private _gid As String
  Private _Rule_No As enuBusinessRuleNO '參數編號(對應程式的列舉編號)
  Private _Rule_Name As String '參數名稱
  Private _Rule_Type1 As String '參數大分類
  Private _Rule_Type2 As String '參數小分類
  Private _Rule_Value As String '參數值
  Private _Enable As Boolean '是否啟用0:禁用1:啟用
  Private _Update_Time As String '更新時間
  Private _Rule_Desc As String '更新時間
  Private _User_Set_Enable As Boolean '是否可讓使用者設定
  Private _objHandling As clsHandlingObject

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property Rule_No() As enuBusinessRuleNO
    Get
      Return _Rule_No
    End Get
    Set(ByVal value As enuBusinessRuleNO)
      _Rule_No = value
    End Set
  End Property
  Public Property Rule_Name() As String
    Get
      Return _Rule_Name
    End Get
    Set(ByVal value As String)
      _Rule_Name = value
    End Set
  End Property
  Public Property Rule_Type1() As String
    Get
      Return _Rule_Type1
    End Get
    Set(ByVal value As String)
      _Rule_Type1 = value
    End Set
  End Property
  Public Property Rule_Type2() As String
    Get
      Return _Rule_Type2
    End Get
    Set(ByVal value As String)
      _Rule_Type2 = value
    End Set
  End Property
  Public Property Rule_Value() As String
    Get
      Return _Rule_Value
    End Get
    Set(ByVal value As String)
      _Rule_Value = value
    End Set
  End Property
  Public Property Enable() As Boolean
    Get
      Return _Enable
    End Get
    Set(ByVal value As Boolean)
      _Enable = value
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
  Public Property Rule_Desc() As String
    Get
      Return _Rule_Desc
    End Get
    Set(ByVal value As String)
      _Rule_Desc = value
    End Set
  End Property
  Public Property User_Set_Enable() As Boolean
    Get
      Return _User_Set_Enable
    End Get
    Set(ByVal value As Boolean)
      _User_Set_Enable = value
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
  Public Sub New(ByVal Rule_No As enuBusinessRuleNO, ByVal Rule_Name As String, ByVal Rule_Type1 As String, ByVal Rule_Type2 As String,
                 ByVal Rule_Value As String, ByVal Enable As Boolean, ByVal Update_Time As String, ByVal Rule_Desc As String, ByVal User_Set_Enable As Boolean)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(Rule_No)
      _gid = key
      _Rule_No = Rule_No
      _Rule_Name = Rule_Name
      _Rule_Type1 = Rule_Type1
      _Rule_Type2 = Rule_Type2
      _Rule_Value = Rule_Value
      _Enable = Enable
      _Update_Time = Update_Time
      _Rule_Desc = Rule_Desc
      _User_Set_Enable = User_Set_Enable
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
  Public Shared Function Get_Combination_Key(ByVal Rule_No As enuBusinessRuleNO) As String
    Try
      Dim key As String = Rule_No
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsClass
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Sub Add_Relationship(ByRef objHandling As clsHandlingObject)
    Try
      '挷定Customer和WMS的關係
      If objHandling IsNot Nothing Then
        _objHandling = objHandling
        objHandling.O_Add_Business_Rule(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      '解除Block和WMS的關係
      If _objHandling IsNot Nothing Then
        _objHandling.O_Remove_Business_Rule(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_M_Business_RuleManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_M_Business_RuleManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_M_Business_RuleManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef obj As clsBusiness_Rule) As Boolean
    Try
      Dim key As String = obj.gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _Rule_No = obj.Rule_No
      _Rule_Name = obj.Rule_Name
      _Rule_Type1 = obj.Rule_Type1
      _Rule_Type2 = obj.Rule_Type2
      _Rule_Value = obj.Rule_Value
      _Enable = obj.Enable
      _Update_Time = obj.Update_Time
      _Rule_Desc = obj.Rule_Desc
      _User_Set_Enable = obj.User_Set_Enable
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
