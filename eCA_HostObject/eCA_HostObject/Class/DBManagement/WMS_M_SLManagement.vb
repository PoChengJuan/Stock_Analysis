Partial Class WMS_M_SLManagement
  Public Shared TableName As String = "WMS_M_SL"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    OWNER_NO
    SL_NO
    SL_ID
    SL_ALIS
    SL_DESC
    BND
    QC_STATUS
    REPORT_TO_HOST
    CREATE_TIME
    UPDATE_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsSL) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}')",
      strSQL,
      TableName,
      IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
      IdxColumnName.SL_NO.ToString, Info.SL_NO,
      IdxColumnName.SL_ID.ToString, Info.SL_ID,
      IdxColumnName.SL_ALIS.ToString, Info.SL_ALIS,
      IdxColumnName.SL_DESC.ToString, Info.SL_DESC,
      IdxColumnName.BND.ToString, Info.BND,
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
      IdxColumnName.REPORT_TO_HOST.ToString, CInt(Info.REPORT_TO_HOST),
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME
     )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsSL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}' WHERE {2}='{3}' And {4}='{5}'",
      strSQL,
      TableName,
      IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
      IdxColumnName.SL_NO.ToString, Info.SL_NO,
      IdxColumnName.SL_ID.ToString, Info.SL_ID,
      IdxColumnName.SL_ALIS.ToString, Info.SL_ALIS,
      IdxColumnName.SL_DESC.ToString, Info.SL_DESC,
      IdxColumnName.BND.ToString, Info.BND,
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
      IdxColumnName.REPORT_TO_HOST.ToString, CInt(Info.REPORT_TO_HOST),
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME
      )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsSL) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' ",
      strSQL,
      TableName,
      IdxColumnName.OWNER_NO.ToString, Info.OWNER_NO,
      IdxColumnName.SL_NO.ToString, Info.SL_NO,
      IdxColumnName.SL_ID.ToString, Info.SL_ID,
      IdxColumnName.SL_ALIS.ToString, Info.SL_ALIS,
      IdxColumnName.SL_DESC.ToString, Info.SL_DESC,
      IdxColumnName.BND.ToString, Info.BND,
      IdxColumnName.QC_STATUS.ToString, CInt(Info.QC_STATUS),
      IdxColumnName.REPORT_TO_HOST.ToString, CInt(Info.REPORT_TO_HOST),
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME
      )
      Dim NewSQL As String = ""
      If SQLCorrect(strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Private Shared Function SetInfoFromDB(ByRef Info As clsSL, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim OWNER_NO = "" & RowData.Item(IdxColumnName.OWNER_NO.ToString)
        Dim SL_NO = "" & RowData.Item(IdxColumnName.SL_NO.ToString)
        Dim SL_ID = "" & RowData.Item(IdxColumnName.SL_ID.ToString)
        Dim SL_ALIS = "" & RowData.Item(IdxColumnName.SL_ID.ToString)
        Dim SL_DESC = "" & RowData.Item(IdxColumnName.SL_ID.ToString)
        Dim BND = If(IsNumeric(RowData.Item(IdxColumnName.BND.ToString)), RowData.Item(IdxColumnName.BND.ToString), 0 & RowData.Item(IdxColumnName.BND.ToString))
        Dim QC_STATUS = If(IsNumeric(RowData.Item(IdxColumnName.QC_STATUS.ToString)), RowData.Item(IdxColumnName.QC_STATUS.ToString), 0 & RowData.Item(IdxColumnName.QC_STATUS.ToString))
        Dim REPORT_TO_HOST = If(IsNumeric(RowData.Item(IdxColumnName.REPORT_TO_HOST.ToString)), RowData.Item(IdxColumnName.REPORT_TO_HOST.ToString), 0 & RowData.Item(IdxColumnName.REPORT_TO_HOST.ToString))
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Info = New clsSL(OWNER_NO, SL_NO, SL_ID, SL_ALIS, SL_DESC, BND, QC_STATUS, REPORT_TO_HOST, CREATE_TIME, UPDATE_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetWMS_M_SLListByALL() As Dictionary(Of String, clsSL)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsSL)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} ",
        strSQL,
        TableName
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsSL = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info.gid, Info)
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetdicSLBy_Owner_SL(ByVal OWNER_NO As String, ByVal SL_NO As String) As Dictionary(Of String, clsSL)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsSL)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strSKUNoList As String = ""


        'If i - count_flag > 800 OrElse i = (dicSKUNo.Count - 1) Then
        'count_flag = i
        strWhere = ""
        If OWNER_NO <> "" Then
          If strWhere = "" Then
            strWhere = String.Format("WHERE {0} IN ('{1}') ", IdxColumnName.OWNER_NO.ToString, OWNER_NO)
          Else
            strWhere = String.Format("{0} AND {1} = ('{2}') ", strWhere, IdxColumnName.OWNER_NO.ToString, OWNER_NO)
          End If
        End If
        If SL_NO <> "" Then
          If strWhere = "" Then
            strWhere = String.Format("WHERE {0} IN ('{1}') ", IdxColumnName.SL_NO.ToString, SL_NO)
          Else
            strWhere = String.Format("{0} AND {1} = ('{2}') ", strWhere, IdxColumnName.SL_NO.ToString, SL_NO)
          End If
        End If

        strSQL = String.Format("Select * from {1} {2} ",
            strSQL,
            TableName,
            strWhere
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsSL = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
          Next
        End If
        strSKUNoList = ""
        'End If
        'Next

      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetdicSLBydicSL(ByVal dicSL As Dictionary(Of String, clsSL)) As Dictionary(Of String, clsSL)       'ALAN0315
    Try
      Dim _lstReturn As New Dictionary(Of String, clsSL)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strSL_NO_OWNER_NOList As String = ""
        Dim count_flag = 0

        For Each objSL In dicSL.Values

          strWhere = ""
          Dim OWNER_NO = objSL.OWNER_NO
          Dim SL_NO = objSL.SL_NO
          If OWNER_NO <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} IN ('{1}') ", IdxColumnName.OWNER_NO.ToString, OWNER_NO)
            Else
              strWhere = String.Format("{0} AND {1} = ('{2}') ", strWhere, IdxColumnName.OWNER_NO.ToString, OWNER_NO)
            End If
          End If
          If SL_NO <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} IN ('{1}') ", IdxColumnName.SL_NO.ToString, SL_NO)
            Else
              strWhere = String.Format("{0} AND {1} = ('{2}') ", strWhere, IdxColumnName.SL_NO.ToString, SL_NO)
            End If
          End If

          strSQL = String.Format("Select * from {1} {2} ",
                strSQL,
                TableName,
                strWhere
            )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
          DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsSL = Nothing
              SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
              If _lstReturn.ContainsKey(Info.gid) = False Then
                _lstReturn.Add(Info.gid, Info)
              End If
            Next
          End If

          'End If
          'Next
        Next
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
