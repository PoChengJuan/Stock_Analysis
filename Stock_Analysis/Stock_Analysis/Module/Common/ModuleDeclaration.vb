Public Module ModuleDeclaration
    '----Main----
    Public gMain As HostMain
    Public objDBLock As New Object
    Public Refreshflag As Boolean = False
    Public SystemDeviceID As String = "HostHandling"

    Public Handling_CycleStop As Integer = 0 '-WMS階段性停止Flag

    Public ConfigPath As String = Application.StartupPath & "\Config.xml"

    Public Const DBShunKangGoodsTimeFormat As String = "yyMMddHHmmssfff"
    Public Const DBFullTimeFormat As String = "yyyy/MM/dd HH:mm:ss.fff"
    Public Const DBTimeFormat As String = "yyyy/MM/dd HH:mm:ss"
    Public Const DBDayFormat As String = "yyyy/MM/dd"
    Public Const DBMonthFormat As String = "yyyy/MM"
    Public Const DBDate_IDFormat As String = "yyMMdd"
    Public Const DBDate_IDFormat_TaiYin As String = "yyyyMMdd"
  Public Const DBDate_IDFormat_yyyyMMdd As String = "yyyyMMdd"

  Public Const DBOnlyTimeFormat As String = "HH:mm:ss"
    Public Const OneDaySeconds As Integer = 86400 '
    Public Const DBOnlyHMFormat As String = "HH:mm"

    Public MCSRpcURL As String
    Public AmatMCSRpcURL As String
    Public DBTool_Name As String 'DBTool记录挡案资料夹
#Region "WCF用參數"
    Public WCFHostIpPort As String
    Public WMS_Service_Address As String 'WMS的WSDL URL
    Public GUI_Service_Address As String 'GUI的WSDL URL
    Public MCS_Service_Address As String 'MCS的WSDL URL
    Public NS_Service_Address As String 'NS的WSDL URL
#End Region
#Region "WEBAPI用參數"
    Public FROM_WMS_API_URL As String '接收WMS的API的URL
    Public TO_WMS_API_URL As String 'WMS的WEBAPI URL
    Public FROM_GUI_API_URL As String '接收GUI的API的URL
    Public TO_GUI_API_URL As String 'GUI的WEBAPI URL
    Public FROM_MCS_API_URL As String '接收MCS的API的URL
    Public TO_MCS_API_URL As String 'MCS的WEBAPI URL
    Public FROM_NS_API_URL As String '接收NS的API的URL
    Public TO_NS_API_URL As String 'NS的WEBAPI URL
    Public HOSTIpPort As String
#End Region

    '上報設定使用
    Public ReportTime As New List(Of String) '-上報時間

    Public f_CommonERPSwitch As Boolean = False '通用型功能開關，通常是對頂新使用，放在config做成開關
    Public f_AutoDownLoad As Boolean = False  '對上位系統用DB溝通
    '客製化欄位
    Public AutoReportSleepTime As String = "60000" 'USI自動過帳間隔時間

    'ERP 成功回填的resule
    Public ERP_Result = "success"

    '預設使用DB進行交握
    '異質系統和Handling的交握方式
    Public GUIToHandlingInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'GUI對Handling的交握方式
    Public MCSToHandlingInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'MCS對Handling的交握方式
    Public WMSToHandlingInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'WMS對Handling的交握方式
    Public WMSToHandlingInterfaceType_Result As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'WMS對Handling的交握方式 (回覆) Vito_19b18
    Public NSToHandlingInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'NS對Handling的交握方式
    Public HostToHandlingInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  '上位系統對Handling的交握方式
    'Handling和異質系統的交握方式
    Public HandlingToGUIInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'Handling對GUI的交握方式
    Public HandlingToMCSInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'Handling對MCS的交握方式
    Public HandlingToWMSInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'Handling對WMS的交握方式
    Public HandlingToNSInterfaceType As enuHandlingInterfaceType = enuHandlingInterfaceType.DB  'Handling對NS的交握方式

    'Receive Message Action
    Public Msg_Action_Create As String = "Create"
    Public Msg_Action_Modify As String = "Modify"
    Public Msg_Action_Delete As String = "Delete"
    Public Const Msg_Direction_Primary As String = "Primary"
    Public Const Msg_Direction_Secondary As String = "Secondary"

    Public Const Msg_Send_Type_JSON As String = "JSON"
    Public Const Msg_Send_Type_XML As String = "XML"

    Public Const Msg_Application_JSON As String = "application/json"
    Public Const Msg_Application_XML As String = "application/xml"
    Public Const Msg_Application_TEXT As String = "application/text"

    Public Const ASRSDevice As String = "A01"
End Module
