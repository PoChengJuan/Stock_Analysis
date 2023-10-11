Public Class clsUUID
  Private ShareName As String = "UUID"
  Private ShareKey As String = ""

  Private gid As String
  Private gUUID_NO As String '編號名稱
  Private gUUID_SEQ As Double '流水號
  Private gIDLENGTH As Double '流水號長度
  Private gAPPEND As String '編碼規則
  Private gCOMMENTS As String '備註
  Private gRESETABLE As Boolean '超過上限是否自動Reset
  Private gUPDATE_DATE As String '更新日期
  Private objNewUUIDLock As New Object


  'Private gobjWMS As clsWMSObject

  'Public Property RESETABLE() As Boolean
  '  Get
  '    Return _RESETABLE
  '  End Get
  '  Set(ByVal value As Boolean)
  '    _RESETABLE = value
  '  End Set
  'End Property
  'Public Property UPDATE_DATE() As String
  '  Get
  '    Return _UPDATE_DATE
  '  End Get
  '  Set(ByVal value As String)
  '    _UPDATE_DATE = value
  '  End Set
  'End Property

  '物件建立時執行的事件
  Public Sub New(ByVal UUID_NO As String, ByVal UUID_SEQ As Double, ByVal IDLENGTH As Double, ByVal APPEND As String, ByVal COMMENTS As String, ByVal RESETABLE As Boolean, ByVal UPDATE_DATE As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(UUID_NO)
      set_gid(key)
      set_UUID_NO(UUID_NO)
      set_UUID_SEQ(UUID_SEQ)
      set_IDLENGTH(IDLENGTH)
      set_APPEND(APPEND)
      set_COMMENTS(COMMENTS)
      set_RESETABLE(RESETABLE)
      set_UPDATE_DATE(UPDATE_DATE)
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
    'gobjWMS = Nothing
  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal UUID_NO As String) As String
    Try
      Dim key As String = UUID_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  'Public Sub Add_Relationship(ByRef objWMS As clsWMSObject)
  '  Try
  '    '挷定Customer和WMS的關係
  '    If objWMS IsNot Nothing Then
  '      set_objWMS(objWMS)
  '      objWMS.O_Add_UUID(Me)
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub
  'Public Sub Remove_Relationship()
  '  Try
  '    '解除Block和WMS的關係
  '    If gobjWMS IsNot Nothing Then
  '      gobjWMS.O_Remove_UUID(Me)
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub


  '資料加入Dictionary

  '把Device加入gcolDevice

  '資料從Dictionary刪除

  '取得Dictionary內的資料


  '-供他人使用的GET
  '-取得gid
  Public Function get_gid() As String
    Try
      Return gid
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gUUID_NO
  Public Function get_UUID_NO() As String
    Try
      Return gUUID_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gUUID_SEQ
  Public Function get_UUID_SEQ() As String
    Try
      Return gUUID_SEQ
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gIDLENGTH
  Public Function get_IDLENGTH() As String
    Try
      Return gIDLENGTH
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gAPPEND
  Public Function get_APPEND() As String
    Try
      Return gAPPEND
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gCOMMENTS
  Public Function get_COMMENTS() As String
    Try
      Return gCOMMENTS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Function get_RESETABLE() As String
    Try
      Return BooleanConvertToInteger(gRESETABLE)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Function get_UPDATE_DATE() As String
    Try
      Return gUPDATE_DATE
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  ''-得到gobjWMS
  'Public Function get_objWMS() As clsWMSObject
  '  Try
  '    Return gobjWMS
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return Nothing
  '  End Try
  'End Function


  '=================Private Function=======================
  '-內部私人的SET
  '-設定gid
  Private Sub set_gid(ByVal key As String)
    Try
      gid = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gUUID_NO
  Private Sub set_UUID_NO(ByVal key As String)
    Try
      gUUID_NO = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gUUID_SEQ
  Private Sub set_UUID_SEQ(ByVal key As String)
    Try
      gUUID_SEQ = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gIDLENGTH
  Private Sub set_IDLENGTH(ByVal key As String)
    Try
      gIDLENGTH = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gAPPEND
  Private Sub set_APPEND(ByVal key As String)
    Try
      gAPPEND = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMENTS
  Private Sub set_RESETABLE(ByVal key As String)
    Try
      gRESETABLE = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMENTS
  Private Sub set_UPDATE_DATE(ByVal key As String)
    Try
      gUPDATE_DATE = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gCOMMENTS
  Private Sub set_COMMENTS(ByVal key As String)
    Try
      gCOMMENTS = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  ''-設定gobjWMS
  'Private Sub set_objWMS(ByVal objWMS As clsWMSObject)
  '  Try
  '    gobjWMS = objWMS
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub



  '非標準的Function
  '=================Public Function=======================
  Public Function Update_UUID(ByRef objUUID As clsUUID) As Boolean
    Try
      Dim key As String = objUUID.get_gid()
      If key <> get_gid() Then
        SendMessageToLog("Key can not Update, old_Key=" & get_gid() & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      set_UUID_NO(objUUID.gUUID_NO)
      set_UUID_SEQ(objUUID.gUUID_SEQ)
      set_IDLENGTH(objUUID.gIDLENGTH)
      set_APPEND(objUUID.gAPPEND)
      set_COMMENTS(objUUID.gCOMMENTS)
      set_RESETABLE(objUUID.gRESETABLE)
      set_UPDATE_DATE(objUUID.gUPDATE_DATE)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  'Public Function Get_NewUUID() As String
  '  Try
  '    Dim rtnSEQ As String = "" '-回覆的UUID
  '    For Each item In gAPPEND.Split(",")
  '      Select Case item.Chars(0)
  '        Case "@"
  '          If item.Contains("UUID") Then
  '            'For i = 0 To gIDLENGTH.ToString.Length - gUUID_SEQ.ToString.Length - 1
  '            '  rtnSEQ += "0" '-補0
  '            'Next
  '            rtnSEQ += gUUID_SEQ.ToString.PadLeft(gIDLENGTH.ToString.Length, "0")
  '          Else
  '            Try '-防止錯誤日期格式
  '              rtnSEQ += GetNewTime_ByDataTimeFormat(item.TrimStart("@"))
  '            Catch ex As Exception
  '              SendMessageToLog("UUID APPEND Error, UUID_NO=" & get_UUID_NO(), eCALogTool.ILogTool.enuTrcLevel.lvError)
  '            End Try
  '          End If
  '        Case Else
  '          rtnSEQ += item
  '      End Select
  '    Next

  '    '-數值+1
  '    set_UUID_SEQ(gUUID_SEQ + 1)

  '    '-檢查是否過限制
  '    If gUUID_SEQ > gIDLENGTH Then
  '      set_UUID_SEQ(1)
  '    End If

  '    '在資料庫Update UUID的資料
  '    If WMS_M_UUIDManagement.UpdateWMS_M_UUIDData(Me) Then
  '      Return rtnSEQ
  '    Else
  '      SendMessageToLog("Update UUID to DB Failed ,TableName = " & WMS_M_UUIDManagement.TableName & " ,UUID_NO=" & get_UUID_NO(), eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    End If
  '    Return rtnSEQ

  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '    Return Nothing
  '  End Try
  'End Function

  Public Function Get_NewUUID() As String
    Try
      SyncLock objNewUUIDLock
        Dim rtnSEQ As String = "" '-回覆的UUID
        For Each item In get_APPEND.Split(",")
          Select Case item.Chars(0)
            Case "@"
              If item.Contains("UUID") Then
                rtnSEQ += get_UUID_SEQ.ToString.PadLeft(get_IDLENGTH.ToString.Length, "0")
              Else  '除了UUID目前只設定日期
                Try '-防止錯誤日期格式
                  Dim NewDate As String = GetNewTime_ByDataTimeFormat(item.TrimStart("@"))
                  If NewDate <> get_UPDATE_DATE Then
                    set_UPDATE_DATE(NewDate)
                    set_UUID_SEQ(1)
                  End If
                  rtnSEQ += GetNewTime_ByDataTimeFormat(item.TrimStart("@"))
                Catch ex As Exception
                  SendMessageToLog("UUID APPEND Error, UUID_NO=" & get_UUID_NO, eCALogTool.ILogTool.enuTrcLevel.lvError)
                End Try
              End If
            Case Else
              rtnSEQ += item
          End Select
        Next
        '-數值+1
        set_UUID_SEQ(get_UUID_SEQ() + 1)
        '-檢查是否過限制
        If get_UUID_SEQ() > get_IDLENGTH() Then
          set_UUID_SEQ(1)
        End If
        '在資料庫Update UUID的資料
        If WMS_M_UUIDManagement.UpdateWMS_M_UUIDData(Me) Then
          Return rtnSEQ
        Else
          SendMessageToLog("Update UUID to DB Failed ,TableName = " & WMS_M_UUIDManagement.TableName & " ,UUID_NO=" & get_UUID_NO(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
        Return rtnSEQ
      End SyncLock
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
