Public Enum enuBusinessRuleNO
  '自動執行類
  AutoPOToWO = 1      '建立PO單時是否自動轉成工單
  AutoCloseWO = 2     '命令完成時檢查是可以自動結單(Value填入可結單的單據類型，用,隔開1:入庫單/2:出庫單)
  AutoCreateStocktaking = 3     '自動生成盤點單，填入的值為自動生成盤點單的時間

  '客製化類
  Super_Customize = 21    '是否根據廠別進行超級客製
  Customize_WO_ID = 22    '建立工單時，是否使用客製化的方式產生WO_ID
  Report_Time = 23        '定義上報的時間點 用,隔開

  '入出庫類
  MWAutoInbound = 41  '平置倉是否自動完成入庫命令
  MWAutoOutbound = 42 '平置倉是否自動完成出庫命令
  ReceiptEmptyCarrier = 43      '待入庫編輯時是否只能使用空棧板進行編輯
  '出庫類
  RestackPackageIDCheck = 61    '出庫分揀時是否檢查Package_ID
  RestackPackSmaller = 62       '出庫分揀時是否支援可以減少量的方式(減取小於一半的)
  '搬送命令類
  MergeTransferCommand = 81     '相同的Source和Dest是否使用同一Command

  '異質系統對WMS的交握方式
  GUIToWMSInterface = 101         'GUI對WMS的交握方式
  MCSToWMSInterface = 102         'MCS對WMS的交握方式
  'WMS對異質系統的交握方式        
  WMSToGUIInterface = 111          'WMS對GUI的交握方式
  WMSToMCSInterface = 112         'WMS對MCS的交握方式

  FullCarrierQtyForMaterial = 1001  '滿桶入庫時的每桶貨品數量
  FullCarrierQtyStackMethod1 = 1002 '堆疊入庫時歐規棧板的貨品數量
  FullCarrierQtyStackMethod2 = 1003 '堆疊入庫時美規棧板的貨品數量

  'HOST交握用對應DB:WMS_M_BUSINESS_RULE裡面的NO
  HostHandlerToWMSInterface = 9001 'HostHandler對WMS的交握方式
  HostHandlerToGUIInterface = 9002 'HostHandler對GUI的交握方式
  HostHandlerToMCSInterface = 9003 'HostHandler對MCS的交握方式
  HostHandlerToNSInterface = 9004 'HostHandler對NS的交握方式

  WMStoHostHandlerInterface = 9011 'WMS對HostHandler的交握方式
  GUItoHostHandlerInterface = 9012 'GUI對HostHandler的交握方式
  MCStoHostHandlerInterface = 9013 'MCS對HostHandler的交握方式
  NStoHostHandlerInterface = 9014 'NS對HostHandler的交握方式
  HostToHostHandlerInterface = 9015 'Host對HostHandler的交握方式
End Enum
