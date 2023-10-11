Public Class clsPO_LINE_Back
  Private ShareName As String = "PO_LINE"
  Private ShareKey As String = ""

  Private gid As String


  Private gPO_ID As String
  Private gPO_LINE_NO As String
  Private gQTY As Long
  Private gQTY_PROCESS As Long
  Private gQTY_FINISH As Long



  Private gobjWMS As clsHandlingObject



  '物件建立時執行的事件
  Public Sub New(ByVal PO_ID As String, ByVal PO_LINE_NO As String, ByVal QTY As Long, ByVal QTY_PROCESS As Long, ByVal QTY_FINISH As Long)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PO_ID, PO_LINE_NO)
      set_gid(key)
      set_PO_ID(PO_ID)
      set_PO_LINE_NO(PO_LINE_NO)
      set_QTY(QTY)
      set_QTY_PROCESS(QTY_PROCESS)
      set_QTY_FINISH(QTY_FINISH)

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
    gobjWMS = Nothing
  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal PO_ID As String, ByVal PO_LINE_NO As String) As String
    Try
      Dim key As String = PO_ID & "_" & PO_LINE_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsPO_LINE_Back
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Sub Add_Relationship(ByRef objWMS As clsHandlingObject)
    Try
      '挷定Customer和WMS的關係
      If objWMS IsNot Nothing Then
        set_objWMS(objWMS)
        objWMS.O_Add_PO_Line(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub Remove_Relationship()
    Try
      '解除Block和WMS的關係
      If gobjWMS IsNot Nothing Then
        gobjWMS.O_Remove_PO_Line(Me)
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_T_PO_LINEManagement_BackUp.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_PO_LINEManagement_BackUp.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_PO_LINEManagement_BackUp.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

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
  '-取得gPO_ID
  Public Function get_PO_ID() As String
    Try
      Return gPO_ID
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gPO_LINE_NO
  Public Function get_PO_LINE_NO() As String
    Try
      Return gPO_LINE_NO
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gQTY
  Public Function get_QTY() As Long
    Try
      Return gQTY
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gQTY_PROCESS
  Public Function get_QTY_PROCESS() As String
    Try
      Return gQTY_PROCESS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '-取得gQTY_FINISH
  Public Function get_QTY_FINISH() As String
    Try
      Return gQTY_FINISH
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '-得到gobjWMS
  Public Function get_objWMS() As clsHandlingObject
    Try
      Return gobjWMS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function



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
  '-設定gPO_ID
  Private Sub set_PO_ID(ByVal key As String)
    Try
      gPO_ID = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gPO_LINE_NO
  Private Sub set_PO_LINE_NO(ByVal key As String)
    Try
      gPO_LINE_NO = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gQTY
  Private Sub set_QTY(ByVal key As Long)
    Try
      gQTY = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gQTY_PROCESS
  Private Sub set_QTY_PROCESS(ByVal key As Long)
    Try
      gQTY_PROCESS = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '-設定gQTY_FINISH
  Private Sub set_QTY_FINISH(ByVal key As Long)
    Try
      gQTY_FINISH = key
      ShareKey = key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '-設定gobjWMS
  Private Sub set_objWMS(ByVal objWMS As clsHandlingObject)
    Try
      gobjWMS = objWMS
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '非標準的Function
  '=================Public Function=======================
  Public Function Update_To_Memory(ByRef objPO_LINE As clsPO_LINE_Back) As Boolean
    Try
      Dim key As String = objPO_LINE.get_gid()
      If key <> get_gid() Then
        SendMessageToLog("Key can not Update, old_Key=" & get_gid() & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      set_PO_ID(objPO_LINE.get_PO_ID)
      set_PO_LINE_NO(objPO_LINE.get_PO_LINE_NO)
      set_QTY(objPO_LINE.get_QTY)
      set_QTY_PROCESS(objPO_LINE.get_QTY_PROCESS)
      set_QTY_FINISH(objPO_LINE.get_QTY_FINISH)

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '新增資料到DB
  Public Function O_Insert_Class_To_DB(ByRef objWMS As clsHandlingObject) As Boolean
    Try
      '一定要寫成功，才更新記憶體的狀態			
      If WMS_T_PO_LINEManagement_BackUp.AddWMS_T_PO_LINEData(Me) = True Then
        '建立梆定
        Add_Relationship(objWMS)
        Return True
      Else
        SendMessageToLog("Insert Class to DB Failed ,TableName = " & WMS_T_PO_LINEManagement_BackUp.TableName & " ,key=" & get_gid(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '刪除
  Public Function O_Delete_Class_To_DB() As Boolean
    Try
      '一定要寫成功，才更新記憶體的狀態			
      If WMS_T_PO_LINEManagement_BackUp.DeleteWMS_T_PO_LINEData(Me) = True Then
        '解除梆定
        Remove_Relationship()
        Return True
      Else
        SendMessageToLog("Delete Class to DB Failed ,TableName = " & WMS_T_PO_LINEManagement_BackUp.TableName & " ,key=" & get_gid(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
