Imports System.Xml
Imports System.IO
Imports System.Globalization
Imports eCA_HostObject
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel

Module ModuleHelpFunc

  '-寫Log的Functoin
  '-修改日期：2018/06/19 修改人：Mark
  Public Function SendMessageToLog(ByVal message As String, ByVal messageLevel As eCALogTool.ILogTool.enuTrcLevel) As Integer
    Try
      If gMain IsNot Nothing Then
        gMain.SendMessageToLog(message, messageLevel, 2)
      End If
      Return 0
    Catch ex As Exception
      Return -1
    End Try
  End Function

  '-取得現在的時間，並自動轉成傳入的指定格式
  '-修改日期：2018/06/19 修改人：Mark
  Public Function GetNewTime_ByDataTimeFormat(ByVal DateTimeFormat As String) As String
    Try
      Return Now.ToString(DateTimeFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

  ''-取得UUID(流水號)
  ''-修改日期：2018/07/03 修改人：Mark
  'Public Function GetUUID(ByVal UUID_No As String) As String
  '  Try
  '    Dim UUID As String = ""
  '    If gMain.objWMS IsNot Nothing Then
  '      Dim objUUID As eCA_WMSObject.clsUUID = Nothing
  '      If gMain.objWMS.O_Get_UUID(UUID_No, objUUID) = True Then
  '        UUID = objUUID.Get_NewUUID()
  '      End If
  '    End If
  '    Return UUID
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return ""
  '  End Try
  'End Function






  '-延遲時間，單位為秒。 
  '-修改日期：2018/06/04 修改人：xxx
  Public Sub TimeDelay(ByVal lngInput As Single)
    Try
      Dim lngDelay As Single
      lngDelay = DateAndTime.Timer
      Do
        System.Windows.Forms.Application.DoEvents()
        '-判斷時間是否超時
      Loop While Not O_TimeOut(lngDelay, lngInput) '-判斷延遲是否完成

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '-判斷時間是否超時，輸入開始時間與時間長度，若達到該時間長度則回傳True 
  '-修改日期：2018/06/04 修改人：xxx
  Public Function O_TimeOut(ByRef nStartTime As Single, ByRef nTimeLeng As Single) As Boolean
    Try
      '-判斷時間是否已超過設定長度
      '-判斷是否有timeout
      If (((DateAndTime.Timer() + System.Math.Abs(nStartTime - 86400)) * 1000) Mod 86400000) / 1000 > nTimeLeng Then
        Return True
      Else
        Return False
      End If

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '-修改Insert字串中包含特殊符號的部份
  '-將字串單引號改為雙引號，用於寫入資料庫欄位時的修正。 
  '-修改日期：2018/06/04 修改人：xxx
  Public Function ModifyStringApostrophe(ByVal name As String) As String
    Try
      name = name.Replace("'", "''")
      name = name.Replace("&", "'||'&'||'")
      Return name

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '-修改查詢的SQL字串中包含特殊符號的部份
  '-將字串單引號改為雙引號，用於寫入資料庫欄位時的修正。 
  '-修改日期：2018/06/28 修改人：xxx
  'Public Function ModifyStringFilter(ByVal Sql As String) As String
  '  Try
  '    Sql = Sql.Replace("[", "[[]") '// 這句話一定要在下面兩個語句之前，否則作為轉義符的方括號會被當作數據被再次處理 
  '    Sql = Sql.Replace("_", "[_]")
  '    Sql = Sql.Replace("%", "[%]")
  '    Return Sql

  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return Nothing
  '  End Try
  'End Function


  '-檢查字串是否為數字
  '-修改日期：2018/06/04 修改人：xxx
  Public Function IntegerCheck(ByVal InputInteger As String) As Boolean
    Try
      If InputInteger <> "" Then
        Return IsNumeric(InputInteger) '-回傳是否為數字
      Else
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '-檢查字串是否為正數
  '-修改日期：2018/06/04 修改人：xxx
  Public Function IntegerCheckPositive(ByVal InputInteger As String) As Boolean
    Try
      If InputInteger <> "" Then
        If IsNumeric(InputInteger) Then '-確定是數字
          If InputInteger > 0 Then
            Return True '-是數字且大於0回傳 真
          End If
        End If
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '-檢查字串是否為負數
  '-修改日期：2018/06/04 修改人：xxx
  Public Function IntegerCheckNegative(ByVal InputInteger As String) As Boolean
    Try
      If InputInteger <> "" Then
        If IsNumeric(InputInteger) Then '-確定是數字
          If InputInteger < 0 Then
            Return True '-是數字且小於0回傳 真
          End If
        End If
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '-將list用","組成字串 
  '-修改日期：2018/06/04 修改人：xxx
  Public Function CombineListToString(ByVal lstWork As List(Of String)) As String
    Try
      Dim _combine = ""
      '-加入新項目後加上","逗號
      For Each item In lstWork
        _combine += item & ","
      Next

      Return _combine.TrimEnd(",") '-移除最末端的","
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

  '-回傳差異的秒數，輸入兩組時間，回傳其差異的秒數
  '-修改日期：2018/06/04 修改人：xxx
  Public Function SubTractTime_Second(ByVal nowTime As String, ByVal beforeTime As String) As Integer
    Try
      '-有時間格式限制
      Dim _nowTime As DateTime = DateTime.Parse(nowTime)
      Dim _beforeTime As DateTime = DateTime.Parse(beforeTime)
      Dim ts As TimeSpan = _nowTime.Subtract(_beforeTime)

      Return ts.TotalSeconds
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return 0
    End Try
  End Function '-回傳差異幾秒


  '-回傳差異的天數，輸入兩組時間，回傳其差異的天數
  '-修改日期：2018/06/04 修改人：xxx
  Public Function SubTractTime_Day(ByVal nowTime As String, ByVal beforeTime As String) As Integer
    Try
      '-有時間格式限制
      Dim _nowTime As DateTime = DateTime.Parse(nowTime)
      Dim _beforeTime As DateTime = DateTime.Parse(beforeTime)
      Dim ts As TimeSpan = _nowTime.Subtract(_beforeTime)

      Return ts.TotalDays
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return 0
    End Try
  End Function '-回傳差異幾天
  Public Function AddTractTime_Hour(ByVal nowTime As String, ByVal hour As Double) As String
    Try
      '-有時間格式限制
      Dim _nowTime As DateTime = DateTime.Parse(nowTime)

      Return _nowTime.AddHours(hour).ToString(DBTimeFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function '-回傳增加後的時間
  Public Function ParseTimeForDay(ByVal _time As String, ByVal timeFormat As String) As String
    Try
      '-輸入字串與時間格式，例如：20180531 010101","yyyyMMdd HHmmss"
      Dim parsed As DateTime '-轉時間格式
      If DateTime.TryParseExact(_time, timeFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then
        Return parsed.ToString(DBDayFormat)
      Else
        Return ""
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function '傳時間格式近來,以及其format  EX:"yyyyMMddHHmmss"   一律回傳 DBDayFormat
  Public Function ParseTimeForTime(ByVal _time As String, ByVal timeFormat As String) As String
    Try
      '-輸入字串與時間格式，例如：20180531 010101","yyyyMMdd HHmmss"
      Dim parsed As DateTime '-轉時間格式
      If DateTime.TryParseExact(_time, timeFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then
        Return parsed.ToString(DBOnlyTimeFormat)
      Else
        Return ""
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function '傳時間格式近來,以及其format  EX:"yyyyMMddHHmmss"   一律回傳 DBOnlyTimeFormat


  '-時間與對應的時間格式，自動轉成正常格式
  '-輸入字串與時間格式，例如：20180531 010101","yyyyMMdd HHmmss" 輸出為：2018/05/31 01:01:01 一律回傳 DBTimeFormat的格式
  '-修改日期：2018/06/04 修改人：xxx
  Public Function ParseTime(ByVal _time As String, ByVal timeFormat As String) As String
    Try
      '-輸入字串與時間格式，例如：20180531 010101","yyyyMMdd HHmmss"
      Dim parsed As DateTime '-轉時間格式
      If DateTime.TryParseExact(_time, timeFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then
        Return parsed.ToString(DBTimeFormat)
      Else
        Return ""
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function '傳時間格式近來,以及其format  EX:"yyyyMMddHHmmss"   一律回傳 DBTimeFormat的格式


  '-取得現在的時間，並自動轉成DB寫入的格式 yyyy/MM/dd HH:mm:ss
  Public Function GetNewTime_DBFormat() As String
    Try
      Return Now.ToString(DBTimeFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function GetNewTime_DBFormat_yyyymmdd() As String
    Try
      Return Now.ToString(DBDate_IDFormat_yyyyMMdd)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  '-取得現在的時間，並自動轉成DB寫入的格式 yyMMdd
  Public Function GetNewDate_DBFormat() As String
    Try
      Return Now.ToString(DBDate_IDFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  '-取得現在的時間，並自動轉成DB寫入的格式 yyMMdd
  Public Function GetNewDate_DBFormat_yyyyMMdd() As String
    Try
      Return Now.ToString(DBDayFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  '-取得現在的時間，並自動轉成特定的格式  YYMMddHHmmssfff
  Public Function GetNewTime_ShunKangFormat() As String
    Try
      Return Now.ToString(DBShunKangGoodsTimeFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  '-把傳入的Result和ResultMessage組成Json 用於回覆WebService
  '-修改日期：2018/06/04 修改人：xxx
  Public Function CombinationJson(ByVal Result As String, ByVal ResultMessage As String) As String
    Try
      '-WebService的回覆字串組合
      Dim JsonString As String = String.Format("{1}{0}Result{0}:{0}{3}{0},{0}content{0}:[{1}{0}ResultMessage{0}:{0}{4}{0}{2}]{2}", Chr(34), "{", "}", Result, ResultMessage)
      Return JsonString
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function '把傳入的Result和ResultMessage組成Json

  '-把傳入的Result和ResultMessage組成Json 用於回報WebService
  '-此為回傳GetPortCarrrierID所使用的
  '-修改日期：2018/06/04 修改人：xxx
  Public Function CombinationJsonByGetPortCarrrierID(ByVal Result As String, ByVal ResultMessage As String, ByVal CarrierID As String) As String
    Try
      '-WebService的回覆字串組合
      Dim JsonString As String = String.Format("{1}{0}Result{0}:{0}{3}{0},{0}CarrierID{0}:{0}{5}{0},{0}content{0}:[{1}{0}ResultMessage{0}:{0}{4}{0}{2}]{2}", Chr(34), "{", "}", Result, ResultMessage, CarrierID)
      Return JsonString
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function '把傳入的Result和ResultMessage組成Json 'GetPortCarrrierID專用回傳


  '-確認質是否在enum內
  '-修改日期：2018/06/28 修改人：xxx
  Public Function CheckValueInEnum(Of T)(ByVal value As String) As Boolean
    Try
      '-檢查是否為數字
      If IsNumeric(value) = False Then
        SendMessageToLog(value & "非數字", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      '-檢查是否在列舉中
      If [Enum].GetName(GetType(T), Convert.ToInt32(value)) = Nothing Then
        SendMessageToLog("列舉：" & GetType(T).Name & " 沒有 對應的value：" & value, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      'IntegerConvertToBoolean(0)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '-確認內容是否在enum內
  '-修改日期：20190215 修改人：Jerry
  Public Function CheckNameInEnum(Of T)(ByVal neme As String) As Boolean
    Try
      '-檢查是否為數字
      If neme = "" Then
        SendMessageToLog(neme & " 為空", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      '-檢查是否在列舉中
      If [Enum].IsDefined(GetType(T), neme) = Nothing Then
        SendMessageToLog("列舉沒有對應的 name：" & neme, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      'IntegerConvertToBoolean(0)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '-處理數字轉Boolean
  '-修改日期：2018/06/28 修改人：xxx
  Public Function IntegerConvertToBoolean(ByVal value As String) As Boolean
    Try
      '-檢查是否為數字
      If IsNumeric(value) = False Then
        SendMessageToLog(value & "非數字", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      '-若不為0則為真 '-鬆的規範
      If Convert.ToInt32(value) = 0 Then
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''' <summary>
  ''' 輸入SystemStatus與List的時間點(sample:08:00:00) 回傳是否到點 '如資料有缺則建立 08:00:00
  ''' 只負責回傳到點 是否更新則由後續觸發的部分處理
  ''' </summary>
  ''' <param name="SystemStatus"></param>
  ''' <returns></returns>
  Public Function CkeckClockOn(ByVal SystemStatus As enuSystemStatus, ByVal lstReportTime As List(Of String)) As Boolean
    Try
      '檢查當前時間
      Dim Now_Time = DateTime.Now.ToString(DBTimeFormat)
      Dim lstSQL As New List(Of String)
      Dim objLastReportTime As clsSystemStatus = Nothing
      '取得上次上報時間
      If gMain.objHandling.gdicSystemStatus.TryGetValue(clsSystemStatus.Get_Combination_Key(SystemStatus), objLastReportTime) = False Then
        '沒資料 建立資料 '僅建立 下次再觸發
        Dim STATUS_NO = SystemStatus
        Dim STATUS_NAME = SystemStatus.ToString
        Dim STATUS_VALUE = "08:00:00"
        Dim UPDATE_TIME = Now_Time
        Dim STATUS_MODE = enuStatusMode.HostHandler
        Dim STATUS_TYPE1 = ""
        Dim STATUS_TYPE2 = ""
        Dim STATUS_TYPE3 = ""
        Dim STATUS_DESC = SystemStatus.ToString & " 到點觸發。"
        Dim objNewSystemStatus As New clsSystemStatus(STATUS_NO, STATUS_NAME, STATUS_VALUE, UPDATE_TIME, STATUS_MODE, STATUS_TYPE1, STATUS_TYPE2, STATUS_TYPE3, STATUS_DESC)
        If objNewSystemStatus.O_Add_Insert_SQLString(lstSQL) = False Then
          SendMessageToLog("Get Insert SystemStatus Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        Else
          If Common_DBManagement.AddQueued(lstSQL) = False Then
            SendMessageToLog("Insert SystemStatus DB Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
            Return False
          End If
          objNewSystemStatus.Add_Relationship(gMain.objHandling)
        End If
        Return False
      Else
        '有資料 檢查是否到點
        '檢查這個小時是否需上報
        Dim bln_Check = False
        For Each obj In ReportTime
          If Hour(obj) = Hour(Now_Time) Then '這個小時需上報
            bln_Check = True
            Exit For
          End If
        Next
        '如果需上報 檢查上次是哪個時間
        If bln_Check Then
          If DateAndTime.Year(objLastReportTime.UPDATE_TIME) < DateAndTime.Year(Now_Time) Or
            DateAndTime.Month(objLastReportTime.UPDATE_TIME) < DateAndTime.Month(Now_Time) Or
            DateAndTime.Day(objLastReportTime.UPDATE_TIME) < DateAndTime.Day(Now_Time) Or
            Hour(objLastReportTime.UPDATE_TIME) < Hour(Now_Time) Then
            objLastReportTime.UPDATE_TIME = Now_Time
            If objLastReportTime.O_Add_Update_SQLString(lstSQL) = False Then
              SendMessageToLog("Get Update SystemStatus SQL Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If
            If Common_DBManagement.BatchUpdate(lstSQL) = False Then
              SendMessageToLog("HostHandler Update DB Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              Return False
            End If
            Return True
          End If
        End If
      End If

      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''' <summary>
  ''' 輸入SystemStatus與List的時間點(sample:08:00:00) 回傳是否到點 '如資料有缺則建立 08:00:00
  ''' 只負責回傳到點 是否更新則由後續觸發的部分處理
  ''' </summary>
  ''' <param name="SKU"></param>
  ''' <param name="Ori_QTY"></param>
  ''' <param name="PackeUnit"></param>
  ''' <param name="QTY"></param>
  ''' <param name="ret_strResultMsg"></param>
  ''' <returns></returns>
  Public Function SetQTYByPackeUnit(ByVal SKU As String, ByVal Ori_QTY As Long, ByRef PackeUnit As String, ByRef QTY As Long, ByRef ret_strResultMsg As String) As Boolean
    Try
      'Dim ResultQTY = QTY
      Dim dicPackeUnit As New Dictionary(Of String, clsMPackeUnit)
      '確認是否有此包裝名稱
      gMain.objHandling.O_GetDB_dicPackeUnitByPackeUnit(PackeUnit, dicPackeUnit)
      If dicPackeUnit.Any Then

      Else  '無需轉換
        Return True
      End If

      '取得此包裝結構
      Dim dicSKUPackeStructure As New Dictionary(Of String, clsMSKUPackeStructure)
      gMain.objHandling.O_GetDB_dicPackeUnitByPackeUnit(SKU, dicSKUPackeStructure)
      If dicSKUPackeStructure.Any Then
        For Each objSKUPackeStructure In dicSKUPackeStructure.Values
          If objSKUPackeStructure.PACKE_UNIT = PackeUnit Then
            QTY = Ori_QTY * objSKUPackeStructure.QTY
            PackeUnit = objSKUPackeStructure.SUB_PACKE_UNIT
          Else
            ret_strResultMsg = "找不到對應的換算單位：" & PackeUnit
            Return False
          End If
        Next
      Else
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function GetCellData(ByVal Sheet As ISheet, ByVal row As Integer, ByVal cell As Integer) As String
    '為了與實際行/列數一致，取值時要扣回來
    row = row - 1
    cell = cell - 1
    Dim _Row = Sheet.GetRow(row)
    'Dim a = _Row.GetCell(cell).GetType
    If _Row.GetCell(cell).CellType = CellType.Formula Then
      _Row.GetCell(cell).SetCellType(CellType.String)
      Return _Row.GetCell(cell).StringCellValue
    End If
    If _Row.GetCell(cell).ToString = "" Then
      Return Nothing
    Else
      Return _Row.GetCell(cell).ToString
    End If
  End Function
End Module
