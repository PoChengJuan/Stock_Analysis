Imports System.Xml
Imports System.IO
Imports System.Globalization
Public Module ModuleHelpFunc


  Public gLogTool As eCALogTool._ILogTool

  '-寫Log的Functoin
  '-修改日期：2018/06/19 修改人：Mark
  Public Function SendMessageToLog(ByVal message As String, ByVal messageLevel As eCALogTool.ILogTool.enuTrcLevel) As Boolean
    Try
      If gLogTool Is Nothing Then
        gLogTool = clsHandlingObject.gLogTool
      End If
      If gLogTool IsNot Nothing Then
        gLogTool.TraceLog(String.Format("Message:{0}", message), , (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod.Name, messageLevel)
        Return True
      End If
      Return False
    Catch ex As Exception
      Return False
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

    ''' <summary>
    ''' 取得系統的流水號
    ''' </summary>
    ''' <returns></returns>
    Public Function Get_System_GUID() As String
        Try
            Return System.Guid.NewGuid.ToString
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return ""
        End Try
    End Function

    ''-延遲時間，單位為秒。 
    ''-修改日期：2018/06/04 修改人：xxx
    'Public Sub TimeDelay(ByVal lngInput As Single)
    '  Try
    '    Dim lngDelay As Single
    '    lngDelay = DateAndTime.Timer
    '    Do
    '      System.Windows.Forms.Application.DoEvents()
    '      '-判斷時間是否超時
    '    Loop While Not O_TimeOut(lngDelay, lngInput) '-判斷延遲是否完成

    '  Catch ex As Exception
    '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    '  End Try
    'End Sub

    '-判斷時間是否超時，輸入開始時間與時間長度，若達到該時間長度則回傳True 
    '-修改日期：2018/06/04 修改人：Jerry
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
  '-修改日期：2018/06/04 修改人：Jerry
  Public Function ModifyStringApostrophe(ByVal name As String) As String
    Try
      name = name.Replace("'", "''")
      name = name.Replace(" & ", "'||'&'||'")
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
  '    Sql = Sql.Replace(LinkKry, "[_]")
  '    Sql = Sql.Replace("%", "[%]")
  '    Return Sql

  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return Nothing
  '  End Try
  'End Function


  '-檢查字串是否為數字
  '-修改日期：2018/06/04 修改人：Jerry
  Public Function IntegerCheck(ByVal InputInteger As String) As Boolean
    Try
      Return IsNumeric(InputInteger) '-回傳是否為數字

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

  '-檢查字串是否為負數
  '-修改日期：2018/06/04 修改人：Jerry
  Public Function IntegerCheckNegative(ByVal InputInteger As String) As Boolean
    Try
      If IsNumeric(InputInteger) Then '-確定是數字
        If InputInteger < 0 Then Return True '-是數字且小於0回傳 真
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

  '-將list用","組成字串 
  '-修改日期：2018/06/04 修改人：Jerry
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
  '-修改日期：2018/06/04 修改人：Jerry
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
  '-修改日期：2018/06/04 修改人：Jerry
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


  '-時間與對應的時間格式，自動轉成正常格式
  '-輸入字串與時間格式，例如：20180531 010101","yyyyMMdd HHmmss" 輸出為：2018/05/31 01:01:01 一律回傳 DBTimeFormat的格式
  '-修改日期：2018/06/04 修改人：Jerry
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

  '-取得現在的時間，並自動轉成DB寫入的格式 yyyy/MM/dd 
  Public Function GetNewDate_DBFormat() As String
    Try
      Return Now.ToString(DBDateFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

  '-取得現在的時間，並自動轉成DB寫入的格式 yyyy/MM/dd HH:mm:ss
  Public Function GetNewTime_DBFormat() As String
    Try
      Return Now.ToString(DBTimeFormat)
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
  '-修改日期：2018/06/28 修改人：Jerry
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
      IntegerConvertToBoolean(0)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '-處理數字轉Boolean
  '-修改日期：2018/06/28 修改人：Jerry
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

  '-處理Boolean轉數字
  '-修改日期：2018/06/29 修改人：Jerry
  Public Function BooleanConvertToInteger(ByVal bool As Boolean) As Integer
    Try
      If bool Then
        Return 1
      Else
        Return 0
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return 0
    End Try
  End Function

  '-處理 ='' 轉 is null
  '-修改日期：2018/06/29 修改人：Jerry
  Public Function SQLCorrect(ByVal SQL As List(Of String), ByRef NewSQL As List(Of String)) As Boolean
    Try
      Dim str As String
      For Each str In SQL
        Dim NewStr = ""
        Dim whereFlag = False
        For Each splitstr In str.Split(" ")
          If whereFlag Then '轉
            NewStr += splitstr.Replace("=''", " is null ") & " "
          Else '不用轉
            NewStr += splitstr & " "
          End If
          '找到where where後的=''轉成 not null
          If splitstr = "WHERE" Or splitstr = "Where" Or splitstr = "where" Then
            whereFlag = True
          End If
        Next
        NewSQL.Add(NewStr)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '-處理 ='' 轉 is null
  '-修改日期：2018/06/29 修改人：Jerry
  Public Function SQLCorrect(ByVal SQL As String, ByRef NewSQL As String) As Boolean
    Try
      Dim NewStr = ""
      Dim whereFlag = False
      For Each splitstr In SQL.Split(" ")
        If whereFlag Then '轉
          NewStr += splitstr.Replace("=''", " is null ") & " "
        Else '不用轉
          NewStr += splitstr & " "
        End If
        '找到where where後的=''轉成 not null
        If splitstr.ToUpper() = "WHERE" Then
          whereFlag = True
        End If
      Next




      NewSQL = NewStr

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''-------------------- 時間相關 --------------------''

  ''' <summary>
  ''' 回傳增加後的時間
  ''' </summary>
  ''' <param name="nowTime"></param>
  ''' <param name="hour"></param>
  ''' <returns></returns>
  Public Function AddTractTime_Hour(ByVal nowTime As String, ByVal hour As Double) As String
    Try
      '-有時間格式限制
      Dim _nowTime As DateTime = DateTime.Parse(nowTime)

      Return _nowTime.AddHours(hour).ToString(DBTimeFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function

  ''' <summary>
  ''' 處理 ='' 轉 is null(針對Oracle和SQLServer有不同的調整方式)
  ''' </summary>
  ''' <param name="SQL"></param>
  ''' <param name="NewSQL"></param>
  ''' <returns></returns>
  Public Function SQLCorrect(ByRef DB_Type As Short, ByVal SQL As String, ByRef NewSQL As String) As Boolean
    Try
      '只進行Where條件部份的轉換
      Dim NewStr = ""
      Dim whereFlag = False
      Select Case DB_Type
        Case 0  'Oracle
          '把WITH(NOLOCK)取代掉
          SQL = SQL.Replace("WITH(NOLOCK)", "")
          '把Where條件中有=''和<>''的取代掉
          For Each splitstr In SQL.Split(" ")
            If whereFlag Then '轉
              If splitstr.IndexOf("=''") <> -1 Then
                NewStr += splitstr.Replace("=''", " is null") & " "
              ElseIf splitstr.IndexOf("<>''") <> -1 Then
                NewStr += splitstr.Replace("<>''", " is not null") & " "
              Else
                NewStr += splitstr & " "
              End If
            Else '不用轉
              NewStr += splitstr & " "
            End If
            '找到where where後的=''轉成 not null
            If splitstr.ToUpper() = "WHERE" Then
              whereFlag = True
            End If
          Next
        Case 1  'SQL Server
          'SQL Server不用進行轉換
          NewStr = SQL
      End Select
      NewSQL = NewStr
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_objCommandReportByUUID(ByRef dic As Dictionary(Of String, clsCommandReport),
                                              ByVal UUID As String,
                                              ByRef ret_obj As clsCommandReport) As Boolean
    Try
      Dim tmp_dic = dic.ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
      Dim ret_dic = tmp_dic.Where(Function(obj)
                                    If obj.Value.UUID <> UUID Then
                                      Return False
                                    End If
                                    Return True
                                  End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        ret_obj = ret_dic.First.Value
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  Public Function O_Get_objHOST_T_COMMAND_REPORTByUUID(ByRef dic As Dictionary(Of String, clsHOST_T_COMMAND_REPORT),
                                              ByVal UUID As String,
                                              ByRef ret_obj As clsHOST_T_COMMAND_REPORT) As Boolean
    Try
      Dim tmp_dic = dic.ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
      Dim ret_dic = tmp_dic.Where(Function(obj)
                                    If obj.Value.UUID <> UUID Then
                                      Return False
                                    End If
                                    Return True
                                  End Function).ToDictionary(Function(obj) obj.Key, Function(obj) obj.Value)
      If ret_dic.Any Then
        ret_obj = ret_dic.First.Value
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Module



