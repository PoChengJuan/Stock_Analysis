
Partial Class WMS_T_Carrier_StatusManagement
  Public Shared TableName As String = "WMS_T_Carrier_Status"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    CARRIER_ID
    LOCATION_NO
    SUBLOCATION_X
    SUBLOCATION_Y
    SUBLOCATION_Z
    RESERVED
    LOCKED
    LOCKED_USER
    LOCKED_REASON
    LOCKED_TIME
    STAGE_ID
    RETURN_LOCATION_NO
    STOCKTAKING_TIME
    FIRSTIN_TIME
    LAST_TRANSFER_TIME
    UPDATE_TIME
    UPDATE_USER_ID
    UNPACK_TIME
    TALLY_ENABLE
  End Enum

  '- GetSQL
  '-請將 clsCarrier 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsCarrier) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38}) values ('{3}','{5}',{7},{9},'{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}',{29},{31},{33},'{35}','{37}','{39}')",
            strSQL,
            TableName,
            IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
            IdxColumnName.LOCATION_NO.ToString, Info.Location_No,
            IdxColumnName.RESERVED.ToString, BooleanConvertToInteger(Info.RESERVED),
            IdxColumnName.LOCKED.ToString, BooleanConvertToInteger(Info.LOCKED),
            IdxColumnName.LOCKED_USER.ToString, Info.LOCKED_USER,
            IdxColumnName.LOCKED_REASON.ToString, Info.LOCKED_REASON,
            IdxColumnName.LOCKED_TIME.ToString, Info.LOCKED_TIME,
            IdxColumnName.STAGE_ID.ToString, Info.STAGE_ID,
            IdxColumnName.RETURN_LOCATION_NO.ToString, Info.RETURN_LOCATION_NO,
            IdxColumnName.STOCKTAKING_TIME.ToString, Info.STOCKTAKING_TIME,
            IdxColumnName.FIRSTIN_TIME.ToString, Info.FIRSTIN_TIME,
            IdxColumnName.LAST_TRANSFER_TIME.ToString, Info.LAST_TRANSFER_TIME,
            IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
            IdxColumnName.SUBLOCATION_X, Info.SUBLOCATION_X,
            IdxColumnName.SUBLOCATION_Y, Info.SUBLOCATION_Y,
            IdxColumnName.SUBLOCATION_Z, Info.SUBLOCATION_Z,
            IdxColumnName.UPDATE_USER_ID, Info.Update_User_ID,
            IdxColumnName.UNPACK_TIME, Info.UNPACK_TIME,
            IdxColumnName.TALLY_ENABLE.ToString, BooleanConvertToInteger(Info.TALLY_ENABLE)
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsCarrier) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
            strSQL,
            TableName,
            IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
            IdxColumnName.LOCATION_NO.ToString, Info.Location_No,
            IdxColumnName.RESERVED.ToString, BooleanConvertToInteger(Info.RESERVED),
            IdxColumnName.LOCKED.ToString, BooleanConvertToInteger(Info.LOCKED),
            IdxColumnName.LOCKED_USER.ToString, Info.LOCKED_USER,
            IdxColumnName.LOCKED_REASON.ToString, Info.LOCKED_REASON,
            IdxColumnName.LOCKED_TIME.ToString, Info.LOCKED_TIME,
            IdxColumnName.STAGE_ID.ToString, Info.STAGE_ID,
            IdxColumnName.RETURN_LOCATION_NO.ToString, Info.RETURN_LOCATION_NO,
            IdxColumnName.STOCKTAKING_TIME.ToString, Info.STOCKTAKING_TIME,
            IdxColumnName.FIRSTIN_TIME.ToString, Info.FIRSTIN_TIME,
            IdxColumnName.LAST_TRANSFER_TIME.ToString, Info.LAST_TRANSFER_TIME,
            IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
            IdxColumnName.TALLY_ENABLE.ToString, BooleanConvertToInteger(Info.TALLY_ENABLE)
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsCarrier) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}={7},{8}={9},{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}={29},{30}={31},{32}={33},{34}='{35}',{36}='{37}',{38}={39} WHERE {2}='{3}'",
            strSQL,
            TableName,
            IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
            IdxColumnName.LOCATION_NO.ToString, Info.Location_No,
            IdxColumnName.RESERVED.ToString, BooleanConvertToInteger(Info.RESERVED),
            IdxColumnName.LOCKED.ToString, BooleanConvertToInteger(Info.LOCKED),
            IdxColumnName.LOCKED_USER.ToString, Info.LOCKED_USER,
            IdxColumnName.LOCKED_REASON.ToString, Info.LOCKED_REASON,
            IdxColumnName.LOCKED_TIME.ToString, Info.LOCKED_TIME,
            IdxColumnName.STAGE_ID.ToString, Info.STAGE_ID,
            IdxColumnName.RETURN_LOCATION_NO.ToString, Info.RETURN_LOCATION_NO,
            IdxColumnName.STOCKTAKING_TIME.ToString, Info.STOCKTAKING_TIME,
            IdxColumnName.FIRSTIN_TIME.ToString, Info.FIRSTIN_TIME,
            IdxColumnName.LAST_TRANSFER_TIME.ToString, Info.LAST_TRANSFER_TIME,
            IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
            IdxColumnName.SUBLOCATION_X, Info.SUBLOCATION_X,
            IdxColumnName.SUBLOCATION_Y, Info.SUBLOCATION_Y,
            IdxColumnName.SUBLOCATION_Z, Info.SUBLOCATION_Z,
            IdxColumnName.UPDATE_USER_ID, Info.Update_User_ID,
            IdxColumnName.UNPACK_TIME, Info.UNPACK_TIME,
            IdxColumnName.TALLY_ENABLE.ToString, BooleanConvertToInteger(Info.TALLY_ENABLE)
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

  '- GET
  Public Shared Function GetWMS_T_Carrier_StatusDataListByALL() As Dictionary(Of String, clsCarrier)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsCarrier)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCarrier = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
          Next
        End If
        'End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsCarrier, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim CARRIER_ID = "" & RowData.Item(IdxColumnName.CARRIER_ID.ToString)
        Dim LOCATION_NO = "" & RowData.Item(IdxColumnName.LOCATION_NO.ToString)
        Dim SUBLOCATION_X = IIf(IsNumeric(RowData.Item(IdxColumnName.SUBLOCATION_X.ToString)), RowData.Item(IdxColumnName.SUBLOCATION_X.ToString), 0)
        Dim SUBLOCATION_Y = IIf(IsNumeric(RowData.Item(IdxColumnName.SUBLOCATION_Y.ToString)), RowData.Item(IdxColumnName.SUBLOCATION_Y.ToString), 0)
        Dim SUBLOCATION_Z = IIf(IsNumeric(RowData.Item(IdxColumnName.SUBLOCATION_Z.ToString)), RowData.Item(IdxColumnName.SUBLOCATION_Z.ToString), 0)
        Dim RESERVED = IntegerConvertToBoolean("" & RowData.Item(IdxColumnName.RESERVED.ToString))
        Dim LOCKED = IntegerConvertToBoolean("" & RowData.Item(IdxColumnName.LOCKED.ToString))
        Dim LOCKED_USER = "" & RowData.Item(IdxColumnName.LOCKED_USER.ToString)
        Dim LOCKED_REASON = "" & RowData.Item(IdxColumnName.LOCKED_REASON.ToString)
        Dim LOCKED_TIME = "" & RowData.Item(IdxColumnName.LOCKED_TIME.ToString)
        Dim STAGE_ID = "" & RowData.Item(IdxColumnName.STAGE_ID.ToString)
        Dim RETUEN_LOCATION_NO = "" & RowData.Item(IdxColumnName.RETURN_LOCATION_NO.ToString)
        Dim STOCKTAKING_TIME = "" & RowData.Item(IdxColumnName.STOCKTAKING_TIME.ToString)
        Dim FIRSTIN_TIME = "" & RowData.Item(IdxColumnName.FIRSTIN_TIME.ToString)
        Dim LAST_TRANSFER_TIME = "" & RowData.Item(IdxColumnName.LAST_TRANSFER_TIME.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim UPDATE_USER_ID = "" & RowData.Item(IdxColumnName.UPDATE_USER_ID.ToString)
        Dim UNPACK_TIME = "" & RowData.Item(IdxColumnName.UNPACK_TIME.ToString)
        Dim TALLY_ENABLE = IntegerConvertToBoolean("" & RowData.Item(IdxColumnName.TALLY_ENABLE.ToString))

        Dim Carrier_Alis As String = ""
        Dim Carrier_Desc As String = ""
        Dim Carrier_Type As Long = 0
        Dim Carrier_Mode As enuCarrierMode = 0
        Dim Create_Time As String = ""
        Dim Comments As String = ""
        Dim SubLocation_Index_X As Long = 0
        Dim SubLocation_Index_Y As Long = 0
        Dim SubLocation_Index_Z As Long = 0

        Dim WEIGHT As Long = 0
        Dim LENGTH As Long = 0
        Dim WIDTH As Long = 0
        Dim HEIGHT As Long = 0
        Dim CARRIER_COMMON01 As String = ""
        Dim CARRIER_COMMON02 As String = ""
        Dim CARRIER_COMMON03 As String = ""
        Dim CARRIER_COMMON04 As String = ""
        Dim CARRIER_COMMON05 As String = ""
        Dim CARRIER_COMMON06 As String = ""
        Dim CARRIER_COMMON07 As String = ""
        Dim CARRIER_COMMON08 As String = ""
        Dim CARRIER_COMMON09 As String = ""
        Dim CARRIER_COMMON10 As String = ""
        Dim CREATE_USER_ID As String = ""


        Info = New clsCarrier(CARRIER_ID, Carrier_Alis, Carrier_Desc, Carrier_Type, Carrier_Mode, Create_Time, Comments, SubLocation_Index_X,
                              SubLocation_Index_Y, SubLocation_Index_Z, LOCATION_NO, RESERVED, LOCKED, LOCKED_USER, LOCKED_REASON,
                              LOCKED_TIME, STAGE_ID, RETUEN_LOCATION_NO, STOCKTAKING_TIME, FIRSTIN_TIME, LAST_TRANSFER_TIME, UPDATE_TIME, WEIGHT,
                              LENGTH, WIDTH, HEIGHT, CARRIER_COMMON01, CARRIER_COMMON02, CARRIER_COMMON03, CARRIER_COMMON04, CARRIER_COMMON05, CARRIER_COMMON06, CARRIER_COMMON07,
                              CARRIER_COMMON08, CARRIER_COMMON09, CARRIER_COMMON10, SUBLOCATION_X, SUBLOCATION_Y, SUBLOCATION_Z, UPDATE_USER_ID, UNPACK_TIME, CREATE_USER_ID, TALLY_ENABLE)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
