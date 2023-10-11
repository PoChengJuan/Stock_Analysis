
Imports eCA_TransactionMessage
''' <summary>
''' 20181109
''' V1.0.0
''' Mark
''' SubMain，主程式啟動式先載入的Module
''' </summary>
''' 
Module ModuleMain
  'HeartBeat Start===============================================================  
  'Dim HB As New FormHeartBeat(9001, 1000)
  'HeartBeat End ===============================================================
  Public ThreadReceiveGUIMessage As Threading.Thread
  Public ThreadReceiveMCSMessage As Threading.Thread
  Public ThreadReceiveWMSMessage As Threading.Thread
  Public ThreadReceiveNSMessage As Threading.Thread
  Public ThreadTransmitWMSMessage As Threading.Thread   'Vito_19b18
  Public ThreadWriteQueueLog As Threading.Thread
  Public ThreadAutoExcute As Threading.Thread
  Public ThreadAutoCheck As Threading.Thread
  Public ThreadAutoReport As Threading.Thread
  Public ThreadHeartBeat As Threading.Thread  'Vito_12b30
  Public ThreadDownLoadMainFile As Threading.Thread

  Public gPO_CheckThread As Threading.Thread
  Public gHttpListenerThread As Threading.Thread
  Public WebSuccessResult = "" '-webservice成功的回覆
  Public WebServiceTimeOut = 1 '-預設120秒
  Public ExcutePO As New Dictionary(Of String, String) '紀錄已執行的PO 在ERP拋單時進行確認後刪除 20190628
  Public HeartBeatAddr As String = ""     'Vito_12b30

  Public FactoryId As String = ""
  Public Companyid As String = ""

  Public ClientCredentialsUserName = "" '-UserName
  Public ClientCredentialsPassword = "" '-Password
  Public bln_SendAccountMail As Boolean = False

  Public DownLoad_Interval_SKU As Long = 1  '撈取ERP料品主檔的時間間隔，預設180 = 3分鐘(180秒)
  Public DownLoad_Interval_PO As Long = 180   '撈取ERP單據主檔的時間間隔，預設180 = 3分鐘(180秒)

  Public HeartBeat As eCAHeartBeat.clsHeartBeat = New eCAHeartBeat.clsHeartBeat(eCAHeartBeat.EnumDeclaration.enuConnectMode.Server) 'Vito_12b30
  ''' <summary>
  ''' 程式啟動後第一個載入的Function
  ''' </summary>
  Public Sub Main()
    Try
      ''確認該程式名稱是否已經被執行了
      'If UBound(Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName)) > 0 Then
      '  MsgBox(Process.GetCurrentProcess.ProcessName & " is running ! Path[" & Application.StartupPath & "]", vbCritical, "Error")
      '  End
      'End If
      Dim xmlStr As String = ""
      Dim tmp_objtest As MSG_Test = Nothing
      tmp_objtest = ParseXmlStringToClass(Of MSG_Test)(xmlStr.ToString)
      '顯示程式在進行Initialize
      FormMsg.SetFormMessage("Initialize HOST program, please wait...")
      FormMsg.Show()
      TimeDelay(1)
      '建立gMain物件
      gMain = New HostMain
      FormMsg.SetFormMessage("Get parameters from Configuration file, please wait...")
      '讀取Config
      If Not I_Load_Config() Then End
      '建立objHandling物件(截入DB的資料建立物件)
      Dim RetMsg As String = ""
      If gMain.Init_objHandling(RetMsg) = False Then
        MsgBox(RetMsg)
        End
      End If
      '載入部份需要預先從記憶體載入WMS變數資訊
      If gMain.Init_SystemValue(RetMsg) = False Then
        MsgBox(RetMsg)
        End
      End If
      '關閉Init Message
      FormMsg.Close()
      '開程FromMain
      Application.Run(FormMain)
    Catch ex As Exception
      MsgBox(ex.Message.ToString)
      End
    End Try
  End Sub
  Private Function I_Load_Config() As Boolean
    Try
      SendMessageToLog("Load Config...", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      ' Dim strSection As String
      If My.Computer.FileSystem.FileExists(ConfigPath) Then
        Dim ConfigXML As XDocument = XDocument.Load(ConfigPath)

        'Public
        Dim Common As IEnumerable =
         From c
         In ConfigXML.<Config>.<Public>

        For Each Pub As XElement In Common
          'RootPath
          gMain.RootPath = Pub.<RootPath>.Value
          If Right(gMain.RootPath, 1) <> "\" Then
            gMain.RootPath = gMain.RootPath & "\"
          End If

          'LogPath
          gMain.LogPath = Pub.<LogPath>.Value
          If Right(gMain.LogPath, 1) <> "\" Then
            gMain.LogPath = gMain.LogPath & "\"
          End If

          'ExportPath
          gMain.ExportPath = Pub.<ExportPath>.Value
          If Right(gMain.ExportPath, 1) <> "\" Then
            gMain.ExportPath = gMain.ExportPath & "\"
          End If
          If gMain.ExportPath = "" Then
            gMain.ExportPath = gMain.LogPath
          End If

          Dim hasDrive As Integer = 0
          Dim linq = From driver In IO.Directory.GetLogicalDrives Where Left(driver, 1).ToLower() = Left(gMain.RootPath, 1).ToLower()

          For Each driveletter As String In linq
            hasDrive = 1
            Exit For
          Next
          If hasDrive = 0 Then
            MsgBox("Please check the RootPath in Config.xml.")
            Return False
          End If

          'TraceLevel
          gMain.TraceLevel = Pub.<TraceLv>.Value
          'MaxViewLine
          gMain.LogTool.MaxViewLine = Pub.<MaxViewLine>.Value
          'ViewLV
          gMain.LogTool.ViewLV = Pub.<ViewLV>.Value

          gMain.RefreshAliveTime = IIf(Pub.<RefreshAliveTime>.Value Is Nothing OrElse Pub.<RefreshAliveTime>.Value = 0, 5000, Pub.<RefreshAliveTime>.Value)

          'WCFHostIpPort
          WCFHostIpPort = Pub.<WCFHostIpPort>.Value

          AutoReportSleepTime = Pub.<AutoReportSleepTime>.Value

          WebServiceTimeOut = Pub.<WebServiceTimeOut>.Value

          FROM_WMS_API_URL = Pub.<FROM_WMS_API_URL>.Value
          TO_WMS_API_URL = Pub.<TO_WMS_API_URL>.Value
          FROM_GUI_API_URL = Pub.<FROM_GUI_API_URL>.Value
          TO_GUI_API_URL = Pub.<TO_GUI_API_URL>.Value
          FROM_MCS_API_URL = Pub.<FROM_MCS_API_URL>.Value
          TO_MCS_API_URL = Pub.<TO_MCS_API_URL>.Value
          FROM_NS_API_URL = Pub.<FROM_NS_API_URL>.Value
          TO_NS_API_URL = Pub.<TO_NS_API_URL>.Value
          HOSTIpPort = Pub.<HOSTIpPort>.Value

          '對應頂新ERP的功能開關，非0就是要走頂新的功能
          f_CommonERPSwitch = IntegerConvertToBoolean(Pub.<CommonERPSwitch>.Value)

          'Dim _lstReportTime = Pub.<ReportTime>.Value
          'ReportTime = _lstReportTime.Split(",").ToList

          HeartBeatAddr = Pub.<HeartBeatPort>.Value                                                               'Vito_12b30

          FactoryId = Pub.<FactoryId>.Value
          Companyid = Pub.<Companyid>.Value

          DownLoad_Interval_SKU = Pub.<DownLoad_Interval_SKU>.Value
          DownLoad_Interval_PO = Pub.<DownLoad_Interval_PO>.Value
        Next

        Dim HeartBeatIP_Port As String() = Split(HeartBeatAddr, ":")                                              'Vito_12b30
        Dim ret_HeartBeat As String = ""                                                                          'Vito_12b30
        'If HeartBeat.CreateHeartBeat(HeartBeatIP_Port(0), HeartBeatIP_Port(1), ret_HeartBeat) = False Then        'Vito_12b30
        '  SendMessageToLog(ret_HeartBeat, eCALogTool.ILogTool.enuTrcLevel.lvError)                                'Vito_12b30
        'End If                                                                                                    'Vito_12b30

        Dim DB As IEnumerable =
         From c
         In ConfigXML.<Config>.<DBList>

        For Each DBList As XElement In DB
          For Each DBInfo As XElement In DBList.Nodes
            'DBTool记录挡案资料夹
            DBTool_Name = "Host_" & DBInfo.<DBName>.Value & "_" & DBInfo.<DBType>.Value & "_" & DBInfo.<UID>.Value

            Dim objDBInfo As New eCA_DBTool.clsDBTool
            Dim key = DBInfo.<DBEnum>.Value '(1:WMS/2:MCS)
            If key = 1 Then
              objDBInfo = New eCA_DBTool.clsDBTool(DBTool_Name)
            ElseIf key = 2 Then
              objDBInfo = New eCA_DBTool.clsDBTool(DBTool_Name)
            End If

            'DBType
            objDBInfo.m_nDBType = DBInfo.<DBType>.Value

            'DBServer
            objDBInfo.m_szDBServer = DBInfo.<DBServer>.Value

            'DBName
            objDBInfo.m_szDBName = DBInfo.<DBName>.Value

            'UID
            objDBInfo.m_szDBUID = DBInfo.<UID>.Value

            'PWD
            objDBInfo.m_szDBPWD = DBInfo.<PWD>.Value

            gMain.DB.Add(key, objDBInfo)
          Next
        Next

        '抓出RabbitMQList的資料
        Dim RabbitMQ As IEnumerable = From c In ConfigXML.<Config>.<MQList>
        For Each MQList As XElement In RabbitMQ
          '取得每一個DB連線設定的資料
          For Each MQInfo As XElement In MQList.Nodes
            'RabbitMQ設定檔
            Dim Enable = MQInfo.<Enable>.Value
            Dim strRabbitMQIp = MQInfo.<Ip>.Value
            Dim strRabbitMQPort = MQInfo.<Port>.Value
            Dim objRabbitMQInfo As New clsMQ(Enable, strRabbitMQIp, strRabbitMQPort)

            gMain.RabbitMQ = objRabbitMQInfo
          Next
        Next

        With gMain
          .LogTool.APName = My.Application.Info.AssemblyName
          .LogTool.LogPath = .RootPath & .LogPath
          .LogTool.TraceLevel = .TraceLevel
          '.LogTool.InitialEnd = True
        End With
      Else
        MsgBox("Cannot find config file: " & ConfigPath)
      End If
      SendMessageToLog("Load Config Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Return True
    Catch ex As Exception
      MsgBox(ex.ToString, MsgBoxStyle.Question)
      Return False
    End Try
  End Function

End Module
