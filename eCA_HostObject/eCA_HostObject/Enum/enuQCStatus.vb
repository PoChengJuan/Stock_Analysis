Public Enum enuQCStatus
  '"QC判定狀態
  NULL = 0
  OK = 1
  NG = 2
  NA = 3


End Enum

Public Enum enuWO_DTL_DC_Status
  Queued = 0            '未處理
  AssignDAS = 1         '指派給播種站(WO_ID)
  AssignDASFinish = 2   '全部的資料都寫給DAS系統了
  DASFinish = 3         '播種站執行完成
  AssignDPS = 11        '指派給摘取站
  DPSFinish = 12        '摘取執行完成
End Enum
''' <summary>
''' 進行QC方式的定義(WMS_T_WO_DTL)
''' </summary>
Public Enum enuQCMethod
  '"QC方式
  Null = 0        '未進行設定 
  NoQC = 1        '不進行QC
  AnyWayQC = 2    '進行QC
End Enum
''' <summary>
''' 加工類型(WMS_T_WO_DTL)
''' </summary>
Public Enum enuWorkingType
  Null = 0        '無
  Split = 1       '拆解
  Combination = 2 '組合
End Enum