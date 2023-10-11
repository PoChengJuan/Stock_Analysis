Imports eCA_HostObject

''' <summary>
''' 20181117
''' V1.0.0
''' Mark
''' HostHandler的主Class
''' </summary>
Public Class HostMain
  '----Common----
  '根目錄
  Public RootPath As String
  'Log目錄
  Public LogPath As String
  'Export目录
  Public ExportPath As String

  Public RefreshAliveTime As Integer = 3000

  '----DBTool----
  Public DB As Dictionary(Of String, eCA_DBTool.clsDBTool)

  '----LogTool----
  Public LogTool As eCALogTool._ILogTool

  '---Trace Log ---
  Public TraceLevel As Byte

  '------Obj主體------
  Public objHandling As clsHandlingObject

  '~~~~~~~~~~~~~~~~~~~~記錄執行緒的Count~~~~~~~~~~~~~~~~~~'
  Public int_tGUIDBHandle As Integer = 0
  Public int_tMCSDBHandle As Integer = 0
  Public int_tWMSDBHandle As Integer = 0

#Region "MQ"
  '----RabbitMQTool----
  Public RabbitMQ As clsMQ = Nothing
  Public dicWMS_TO_HANDLING_Queue As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)
  Public dicWMS_TO_HANDLING_Queue_S As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)  '處理RESULT
  Public dicWMS_TO_HANDLING_Queue_R As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)  '處理WAIT_UUID
  Public dicMCS_TO_HANDLING_Queue As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)
  Public dicMCS_TO_HANDLING_Queue_S As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)  '處理RESULT
  Public dicMCS_TO_HANDLING_Queue_R As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)  '處理WAIT_UUID
  Public dicGUI_TO_HANDLING_Queue As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)
  Public dicGUI_TO_HANDLING_Queue_S As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)  '處理RESULT
  Public dicGUI_TO_HANDLING_Queue_R As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)  '處理WAIT_UUID
  Public dicNS_TO_HANDLING_Queue As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)
  Public dicNS_TO_HANDLING_Queue_S As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)  '處理RESULT
  Public dicNS_TO_HANDLING_Queue_R As New Dictionary(Of String, RabbitMQ.Client.Events.BasicDeliverEventArgs)  '處理WAIT_UUID
#End Region


  Public Sub New()
    DB = New Dictionary(Of String, eCA_DBTool.clsDBTool)
    LogTool = New eCALogTool.CLogTool
    'ASRSCtrl = New ASRSController(LogTool)
    'ftpIN = New FTPClient
    'ftpOUT = New FTPClient
  End Sub
  Private Sub Class_Terminate_Renamed()
    Try
      DB = Nothing
      LogTool = Nothing
      objHandling = Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Protected Overrides Sub Finalize()
    Class_Terminate_Renamed()
    MyBase.Finalize()
  End Sub
  ''' <summary>
  ''' 初始化objHandling
  ''' </summary>
  ''' <param name="RetMsg"></param>
  ''' <returns></returns>
  Public Function Init_objHandling(ByRef RetMsg As String) As Boolean
    Try
      objHandling = New clsHandlingObject(DB, LogTool, RetMsg)
      If RetMsg = "" Then
        Return True
      Else
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 讀取資料庫的設定參數
  ''' </summary>
  ''' <param name="RetMsg"></param>
  ''' <returns></returns>
  Public Function Init_SystemValue(ByRef RetMsg As String) As Boolean
    Try

#Region "系統間的交握介面(9000~9100)"
      '確認個系統間的交握使用的介面，並執行相對應的操作
      '確認GUIToHandlingInterface
      Dim objBusiness_GUIToHandlingInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.GUIToHandlingInterfaceType = enuHandlingInterfaceType.DB  'GUI對HostHandler的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.GUItoHostHandlerInterface, objBusiness_GUIToHandlingInterfaceType) = True Then
        If objBusiness_GUIToHandlingInterfaceType IsNot Nothing AndAlso objBusiness_GUIToHandlingInterfaceType.Enable = True Then
          If CheckValueInEnum(Of enuHandlingInterfaceType)(objBusiness_GUIToHandlingInterfaceType.Rule_Value) = True Then
            ModuleDeclaration.GUIToHandlingInterfaceType = CInt(objBusiness_GUIToHandlingInterfaceType.Rule_Value)
          Else
            SendMessageToLog("{0} not defined, RULE_NO =" & objBusiness_GUIToHandlingInterfaceType.Rule_No & ", RULE_Name =" & objBusiness_GUIToHandlingInterfaceType.Rule_Name & ", RULE_Value =" & objBusiness_GUIToHandlingInterfaceType.Rule_Value, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        End If
      End If

      '確認MCStoHostHandlerInterface
      Dim objBusiness_MCSToHandlingInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.MCSToHandlingInterfaceType = enuHandlingInterfaceType.DB  'MCS對HostHandler的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.MCStoHostHandlerInterface, objBusiness_MCSToHandlingInterfaceType) = True Then
        If objBusiness_MCSToHandlingInterfaceType IsNot Nothing AndAlso objBusiness_MCSToHandlingInterfaceType.Enable = True Then
          If CheckValueInEnum(Of enuHandlingInterfaceType)(objBusiness_MCSToHandlingInterfaceType.Rule_Value) = True Then
            ModuleDeclaration.MCSToHandlingInterfaceType = CInt(objBusiness_MCSToHandlingInterfaceType.Rule_Value)
          Else
            SendMessageToLog("{0} not defined, RULE_NO =" & objBusiness_MCSToHandlingInterfaceType.Rule_No & ", RULE_Name =" & objBusiness_MCSToHandlingInterfaceType.Rule_Name & ", RULE_Value =" & objBusiness_MCSToHandlingInterfaceType.Rule_Value, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        End If
      End If

      '確認WMStoHostHandlerInterface
      Dim objBusiness_WMSToHandlingInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.WMSToHandlingInterfaceType = enuHandlingInterfaceType.DB  'WMS對HostHandler的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.WMStoHostHandlerInterface, objBusiness_WMSToHandlingInterfaceType) = True Then
        If objBusiness_WMSToHandlingInterfaceType IsNot Nothing AndAlso objBusiness_WMSToHandlingInterfaceType.Enable = True Then
          If CheckValueInEnum(Of enuHandlingInterfaceType)(objBusiness_WMSToHandlingInterfaceType.Rule_Value) = True Then
            ModuleDeclaration.WMSToHandlingInterfaceType = CInt(objBusiness_WMSToHandlingInterfaceType.Rule_Value)
          Else
            SendMessageToLog("{0} not defined, RULE_NO =" & objBusiness_WMSToHandlingInterfaceType.Rule_No & ", RULE_Name =" & objBusiness_WMSToHandlingInterfaceType.Rule_Name & ", RULE_Value =" & objBusiness_WMSToHandlingInterfaceType.Rule_Value, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        End If
      End If

      '確認NStoHostHandlerInterface
      Dim objBusiness_NSToHandlingInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.NSToHandlingInterfaceType = enuHandlingInterfaceType.DB  'NS對HostHandler的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.NStoHostHandlerInterface, objBusiness_NSToHandlingInterfaceType) = True Then
        If objBusiness_NSToHandlingInterfaceType IsNot Nothing AndAlso objBusiness_NSToHandlingInterfaceType.Enable = True Then
          If CheckValueInEnum(Of enuHandlingInterfaceType)(objBusiness_NSToHandlingInterfaceType.Rule_Value) = True Then
            ModuleDeclaration.NSToHandlingInterfaceType = CInt(objBusiness_NSToHandlingInterfaceType.Rule_Value)
          Else
            SendMessageToLog("{0} not defined, RULE_NO =" & objBusiness_NSToHandlingInterfaceType.Rule_No & ", RULE_Name =" & objBusiness_NSToHandlingInterfaceType.Rule_Name & ", RULE_Value =" & objBusiness_NSToHandlingInterfaceType.Rule_Value, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        End If
      End If

      '確認HandlingToMCSInterface
      Dim objBusiness_HandlingToMCSInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.HandlingToMCSInterfaceType = enuHandlingInterfaceType.DB  'HostHandler對MCS的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.HostHandlerToMCSInterface, objBusiness_HandlingToMCSInterfaceType) = True Then
        If objBusiness_HandlingToMCSInterfaceType IsNot Nothing AndAlso objBusiness_HandlingToMCSInterfaceType.Enable = True Then
          If CheckValueInEnum(Of enuHandlingInterfaceType)(objBusiness_HandlingToMCSInterfaceType.Rule_Value) = True Then
            ModuleDeclaration.HandlingToMCSInterfaceType = CInt(objBusiness_HandlingToMCSInterfaceType.Rule_Value)
          Else
            SendMessageToLog("{0} not defined, RULE_NO =" & objBusiness_HandlingToMCSInterfaceType.Rule_No & ", RULE_Name =" & objBusiness_HandlingToMCSInterfaceType.Rule_Name & ", RULE_Value =" & objBusiness_HandlingToMCSInterfaceType.Rule_Value, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        End If
      End If

      '確認HandlingToGUIInterface
      Dim objBusiness_HandlingToGUIInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.HandlingToGUIInterfaceType = enuHandlingInterfaceType.DB  'HostHandler對GUI的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.HostHandlerToGUIInterface, objBusiness_HandlingToGUIInterfaceType) = True Then
        If objBusiness_HandlingToGUIInterfaceType IsNot Nothing AndAlso objBusiness_HandlingToGUIInterfaceType.Enable = True Then
          If CheckValueInEnum(Of enuHandlingInterfaceType)(objBusiness_HandlingToGUIInterfaceType.Rule_Value) = True Then
            ModuleDeclaration.HandlingToGUIInterfaceType = CInt(objBusiness_HandlingToGUIInterfaceType.Rule_Value)
          Else
            SendMessageToLog("{0} not defined, RULE_NO =" & objBusiness_HandlingToGUIInterfaceType.Rule_No & ", RULE_Name =" & objBusiness_HandlingToGUIInterfaceType.Rule_Name & ", RULE_Value =" & objBusiness_HandlingToGUIInterfaceType.Rule_Value, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        End If
      End If

      '確認HandlingToWMSInterface
      Dim objBusiness_HandlingToWMSInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.HandlingToWMSInterfaceType = enuHandlingInterfaceType.DB  'HostHandler對WMS的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.HostHandlerToWMSInterface, objBusiness_HandlingToWMSInterfaceType) = True Then
        If objBusiness_HandlingToWMSInterfaceType IsNot Nothing AndAlso objBusiness_HandlingToWMSInterfaceType.Enable = True Then
          If CheckValueInEnum(Of enuHandlingInterfaceType)(objBusiness_HandlingToWMSInterfaceType.Rule_Value) = True Then
            ModuleDeclaration.HandlingToWMSInterfaceType = CInt(objBusiness_HandlingToWMSInterfaceType.Rule_Value)
          Else
            SendMessageToLog("{0} not defined, RULE_NO =" & objBusiness_HandlingToWMSInterfaceType.Rule_No & ", RULE_Name =" & objBusiness_HandlingToWMSInterfaceType.Rule_Name & ", RULE_Value =" & objBusiness_HandlingToMCSInterfaceType.Rule_Value, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        End If
      End If

      '確認HandlingToNSInterface
      Dim objBusiness_HandlingToNSInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.HandlingToNSInterfaceType = enuHandlingInterfaceType.DB  'HostHandler對NS的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.HostHandlerToNSInterface, objBusiness_HandlingToNSInterfaceType) = True Then
        If objBusiness_HandlingToNSInterfaceType IsNot Nothing AndAlso objBusiness_HandlingToNSInterfaceType.Enable = True Then
          If CheckValueInEnum(Of enuHandlingInterfaceType)(objBusiness_HandlingToNSInterfaceType.Rule_Value) = True Then
            ModuleDeclaration.HandlingToNSInterfaceType = CInt(objBusiness_HandlingToNSInterfaceType.Rule_Value)
          Else
            SendMessageToLog("{0} not defined, RULE_NO =" & objBusiness_HandlingToNSInterfaceType.Rule_No & ", RULE_Name =" & objBusiness_HandlingToNSInterfaceType.Rule_Name & ", RULE_Value =" & objBusiness_HandlingToNSInterfaceType.Rule_Value, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        End If
      End If

      '確認HandlingToNSInterface
      Dim objBusiness_HostToHandlingInterfaceType As clsBusiness_Rule = Nothing
      ModuleDeclaration.HostToHandlingInterfaceType = enuHandlingInterfaceType.DB  'HostHandler對NS的交握方式(重置)
      If gMain.objHandling.O_Get_Business_Rule(enuBusinessRuleNO.HostToHostHandlerInterface, objBusiness_HostToHandlingInterfaceType) = True Then
        If objBusiness_HostToHandlingInterfaceType IsNot Nothing AndAlso objBusiness_HostToHandlingInterfaceType.Enable = True Then
          '這裡不檢查ENUM的存在，因為外層之後會使用AND的作法比對BIT，讓一個物件可以判斷多重條件
          ModuleDeclaration.HostToHandlingInterfaceType = CInt(objBusiness_HostToHandlingInterfaceType.Rule_Value)
        End If
      End If
#End Region

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 寫Log
  ''' </summary>
  ''' <param name="message"></param>
  ''' <param name="messageLevel"></param>
  ''' <returns></returns>
  Public Function SendMessageToLog(ByVal message As String, ByVal messageLevel As eCALogTool.ILogTool.enuTrcLevel) As Boolean
    Try
      If LogTool IsNot Nothing Then
        LogTool.TraceLog(String.Format("Message:{0}", message), , (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod.Name, messageLevel)
      End If
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 寫Log
  ''' </summary>
  ''' <param name="message"></param>
  ''' <param name="messageLevel"></param>
  ''' <param name="frameNum"></param>
  ''' <returns></returns>
  Public Function SendMessageToLog(ByVal message As String, ByVal messageLevel As eCALogTool.ILogTool.enuTrcLevel, ByVal frameNum As Integer) As Boolean
    Try
      If LogTool IsNot Nothing Then
        LogTool.TraceLog(String.Format("Message:{0}", message), , (New System.Diagnostics.StackTrace).GetFrame(frameNum).GetMethod.Name, messageLevel)
      End If
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
End Class
