
Partial Class WMS_M_CarrierManagement
  Public Shared TableName As String = "WMS_M_Carrier"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    CARRIER_ID
    CARRIER_ALIS
    CARRIER_DESC
    CARRIER_TYPE
    CARRIER_MODE
    CREATE_TIME
    COMMENTS
    SUBLOCATION_INDEX_X
    SUBLOCATION_INDEX_Y
    SUBLOCATION_INDEX_Z
    WEIGHT
    LENGTH
    WIDTH
    HEIGHT
    CARRIER_COMMON01
    CARRIER_COMMON02
    CARRIER_COMMON03
    CARRIER_COMMON04
    CARRIER_COMMON05
    CARRIER_COMMON06
    CARRIER_COMMON07
    CARRIER_COMMON08
    CARRIER_COMMON09
    CARRIER_COMMON10
    CREATE_USER_ID
  End Enum

  '- GetSQL
  '-請將 clsCarrier 取代成對應的cls
  '-請將 updateObjData 取代成對應的名稱
  Public Shared Function GetInsertSQL(ByRef Info As clsCarrier) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50}) values ('{3}','{5}','{7}',{9},{11},'{13}','{15}',{17},{19},{21},{23},{25},{27},{29},'{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}')",
      strSQL,
      TableName,
      IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
      IdxColumnName.CARRIER_ALIS.ToString, Info.Carrier_Alis,
      IdxColumnName.CARRIER_DESC.ToString, Info.Carrier_Desc,
      IdxColumnName.CARRIER_TYPE.ToString, CInt(Info.Carrier_Type),
      IdxColumnName.CARRIER_MODE.ToString, CInt(Info.Carrier_Mode),
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.COMMENTS.ToString, Info.Comments,
      IdxColumnName.SUBLOCATION_INDEX_X.ToString, Info.SubLocation_Index_X,
      IdxColumnName.SUBLOCATION_INDEX_Y.ToString, Info.SubLocation_Index_Y,
      IdxColumnName.SUBLOCATION_INDEX_Z.ToString, Info.SubLocation_Index_Z,
      IdxColumnName.WEIGHT.ToString, Info.WEIGHT,
      IdxColumnName.LENGTH.ToString, Info.LENGTH,
      IdxColumnName.WIDTH.ToString, Info.WIDTH,
      IdxColumnName.HEIGHT.ToString, Info.HEIGHT,
      IdxColumnName.CARRIER_COMMON01.ToString, Info.CARRIER_COMMON01,
      IdxColumnName.CARRIER_COMMON02.ToString, Info.CARRIER_COMMON02,
      IdxColumnName.CARRIER_COMMON03.ToString, Info.CARRIER_COMMON03,
      IdxColumnName.CARRIER_COMMON04.ToString, Info.CARRIER_COMMON04,
      IdxColumnName.CARRIER_COMMON05.ToString, Info.CARRIER_COMMON05,
      IdxColumnName.CARRIER_COMMON06.ToString, Info.CARRIER_COMMON06,
      IdxColumnName.CARRIER_COMMON07.ToString, Info.CARRIER_COMMON07,
      IdxColumnName.CARRIER_COMMON08.ToString, Info.CARRIER_COMMON08,
      IdxColumnName.CARRIER_COMMON09.ToString, Info.CARRIER_COMMON09,
      IdxColumnName.CARRIER_COMMON10.ToString, Info.CARRIER_COMMON10,
      IdxColumnName.CREATE_USER_ID.ToString, Info.CREATE_USER_ID
     )
      Return strSQL
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
      IdxColumnName.CARRIER_ALIS.ToString, Info.Carrier_Alis,
      IdxColumnName.CARRIER_DESC.ToString, Info.Carrier_Desc,
      IdxColumnName.CARRIER_TYPE.ToString, Info.Carrier_Type,
      IdxColumnName.CARRIER_MODE.ToString, Info.Carrier_Mode,
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.COMMENTS.ToString, Info.Comments,
      IdxColumnName.SUBLOCATION_INDEX_X.ToString, Info.SubLocation_Index_X,
      IdxColumnName.SUBLOCATION_INDEX_Y.ToString, Info.SubLocation_Index_Y,
      IdxColumnName.SUBLOCATION_INDEX_Z.ToString, Info.SubLocation_Index_Z,
      IdxColumnName.WEIGHT.ToString, Info.WEIGHT,
      IdxColumnName.LENGTH.ToString, Info.LENGTH,
      IdxColumnName.WIDTH.ToString, Info.WIDTH,
      IdxColumnName.HEIGHT.ToString, Info.HEIGHT,
      IdxColumnName.CARRIER_COMMON01.ToString, Info.CARRIER_COMMON01,
      IdxColumnName.CARRIER_COMMON02.ToString, Info.CARRIER_COMMON02,
      IdxColumnName.CARRIER_COMMON03.ToString, Info.CARRIER_COMMON03,
      IdxColumnName.CARRIER_COMMON04.ToString, Info.CARRIER_COMMON04,
      IdxColumnName.CARRIER_COMMON05.ToString, Info.CARRIER_COMMON05,
      IdxColumnName.CARRIER_COMMON06.ToString, Info.CARRIER_COMMON06,
      IdxColumnName.CARRIER_COMMON07.ToString, Info.CARRIER_COMMON07,
      IdxColumnName.CARRIER_COMMON08.ToString, Info.CARRIER_COMMON08,
      IdxColumnName.CARRIER_COMMON09.ToString, Info.CARRIER_COMMON09,
      IdxColumnName.CARRIER_COMMON10.ToString, Info.CARRIER_COMMON10
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsCarrier) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}={9},{10}={11},{12}='{13}',{14}='{15}',{16}={17},{18}={19},{20}={21},{22}={23},{24}={25},{26}={27},{28}={29},{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.CARRIER_ID.ToString, Info.Carrier_ID,
      IdxColumnName.CARRIER_ALIS.ToString, Info.Carrier_Alis,
      IdxColumnName.CARRIER_DESC.ToString, Info.Carrier_Desc,
      IdxColumnName.CARRIER_TYPE.ToString, CInt(Info.Carrier_Type),
      IdxColumnName.CARRIER_MODE.ToString, CInt(Info.Carrier_Mode),
      IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
      IdxColumnName.COMMENTS.ToString, Info.Comments,
      IdxColumnName.SUBLOCATION_INDEX_X.ToString, Info.SubLocation_Index_X,
      IdxColumnName.SUBLOCATION_INDEX_Y.ToString, Info.SubLocation_Index_Y,
      IdxColumnName.SUBLOCATION_INDEX_Z.ToString, Info.SubLocation_Index_Z,
      IdxColumnName.WEIGHT.ToString, Info.WEIGHT,
      IdxColumnName.LENGTH.ToString, Info.LENGTH,
      IdxColumnName.WIDTH.ToString, Info.WIDTH,
      IdxColumnName.HEIGHT.ToString, Info.HEIGHT,
      IdxColumnName.CARRIER_COMMON01.ToString, Info.CARRIER_COMMON01,
      IdxColumnName.CARRIER_COMMON02.ToString, Info.CARRIER_COMMON02,
      IdxColumnName.CARRIER_COMMON03.ToString, Info.CARRIER_COMMON03,
      IdxColumnName.CARRIER_COMMON04.ToString, Info.CARRIER_COMMON04,
      IdxColumnName.CARRIER_COMMON05.ToString, Info.CARRIER_COMMON05,
      IdxColumnName.CARRIER_COMMON06.ToString, Info.CARRIER_COMMON06,
      IdxColumnName.CARRIER_COMMON07.ToString, Info.CARRIER_COMMON07,
      IdxColumnName.CARRIER_COMMON08.ToString, Info.CARRIER_COMMON08,
      IdxColumnName.CARRIER_COMMON09.ToString, Info.CARRIER_COMMON09,
      IdxColumnName.CARRIER_COMMON10.ToString, Info.CARRIER_COMMON10,
      IdxColumnName.CREATE_USER_ID.ToString, Info.CREATE_USER_ID
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function



  '- GET
  Public Shared Function GetWMS_M_CarrierDataListByALL() As List(Of clsCarrier)
    Try
      Dim _lstReturn As New List(Of clsCarrier)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
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
            _lstReturn.Add(Info)
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
  Public Shared Function GetclsCarrierListByCARRIER_ID(ByVal carrier_id As String) As List(Of clsCarrier)
    Try
      Dim _lstReturn As New List(Of clsCarrier)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
          strSQL,
          TableName,
          IdxColumnName.CARRIER_ID.ToString, carrier_id
          )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsCarrier = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info)
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

  '-Function


  '-以下為內部私人用


  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsCarrier, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim CARRIER_ID = "" & RowData.Item(IdxColumnName.CARRIER_ID.ToString)
        Dim CARRIER_ALIS = "" & RowData.Item(IdxColumnName.CARRIER_ALIS.ToString)
        Dim CARRIER_DESC = "" & RowData.Item(IdxColumnName.CARRIER_DESC.ToString)
        Dim CARRIER_TYPE = IIf(IsNumeric(RowData.Item(IdxColumnName.CARRIER_TYPE.ToString)), RowData.Item(IdxColumnName.CARRIER_TYPE.ToString), 0)
        Dim CARRIER_MODE = IIf(IsNumeric(RowData.Item(IdxColumnName.CARRIER_MODE.ToString)), RowData.Item(IdxColumnName.CARRIER_MODE.ToString), 0)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim SUBLOCATION_INDEX_X = IIf(IsNumeric(RowData.Item(IdxColumnName.SUBLOCATION_INDEX_X.ToString)), RowData.Item(IdxColumnName.SUBLOCATION_INDEX_X.ToString), 0)
        Dim SUBLOCATION_INDEX_Y = IIf(IsNumeric(RowData.Item(IdxColumnName.SUBLOCATION_INDEX_Y.ToString)), RowData.Item(IdxColumnName.SUBLOCATION_INDEX_Y.ToString), 0)
        Dim SUBLOCATION_INDEX_Z = IIf(IsNumeric(RowData.Item(IdxColumnName.SUBLOCATION_INDEX_Z.ToString)), RowData.Item(IdxColumnName.SUBLOCATION_INDEX_Z.ToString), 0)
        Dim WEIGHT = IIf(IsNumeric(RowData.Item(IdxColumnName.WEIGHT.ToString)), RowData.Item(IdxColumnName.WEIGHT.ToString), 0)
        Dim LENGTH = IIf(IsNumeric(RowData.Item(IdxColumnName.LENGTH.ToString)), RowData.Item(IdxColumnName.LENGTH.ToString), 0)
        Dim WIDTH = IIf(IsNumeric(RowData.Item(IdxColumnName.WIDTH.ToString)), RowData.Item(IdxColumnName.WIDTH.ToString), 0)
        Dim HEIGHT = IIf(IsNumeric(RowData.Item(IdxColumnName.HEIGHT.ToString)), RowData.Item(IdxColumnName.HEIGHT.ToString), 0)
        Dim CARRIER_COMMON01 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON01.ToString)
        Dim CARRIER_COMMON02 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON02.ToString)
        Dim CARRIER_COMMON03 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON03.ToString)
        Dim CARRIER_COMMON04 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON04.ToString)
        Dim CARRIER_COMMON05 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON05.ToString)
        Dim CARRIER_COMMON06 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON06.ToString)
        Dim CARRIER_COMMON07 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON07.ToString)
        Dim CARRIER_COMMON08 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON08.ToString)
        Dim CARRIER_COMMON09 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON09.ToString)
        Dim CARRIER_COMMON10 = "" & RowData.Item(IdxColumnName.CARRIER_COMMON10.ToString)

        Dim LOCATION_NO = ""
        Dim SUBLOCATION_X = 0
        Dim SUBLOCATION_Y = 0
        Dim SUBLOCATION_Z = 0
        Dim RESERVED = False
        Dim LOCKED = False
        Dim LOCKED_USER = ""
        Dim LOCKED_REASON = ""
        Dim LOCKED_TIME = ""
        Dim STAGE_ID = ""
        Dim RETUEN_LOCATION_NO = ""
        Dim STOCKTAKING_TIME = ""
        Dim FIRSTIN_TIME = ""
        Dim LAST_TRANSFER_TIME = ""
        Dim UPDATE_TIME = ""
        Dim UPDATE_USER_ID = ""
        Dim UNPACK_TIME = ""
        Dim TALLY_ENABLE = False
        Dim CREATE_USER_ID = "" & RowData.Item(IdxColumnName.CREATE_USER_ID.ToString)
        Info = New clsCarrier(CARRIER_ID, CARRIER_ALIS, CARRIER_DESC, CARRIER_TYPE, CARRIER_MODE, CREATE_TIME, COMMENTS, SUBLOCATION_INDEX_X, SUBLOCATION_INDEX_Y,
                              SUBLOCATION_INDEX_Z, LOCATION_NO, RESERVED, LOCKED, LOCKED_USER, LOCKED_REASON, LOCKED_TIME, STAGE_ID, RETUEN_LOCATION_NO,
                              STOCKTAKING_TIME, FIRSTIN_TIME, LAST_TRANSFER_TIME, UPDATE_TIME, WEIGHT, LENGTH, WIDTH, HEIGHT, CARRIER_COMMON01, CARRIER_COMMON02,
                              CARRIER_COMMON03, CARRIER_COMMON04, CARRIER_COMMON05, CARRIER_COMMON06, CARRIER_COMMON07, CARRIER_COMMON08, CARRIER_COMMON09,
                              CARRIER_COMMON10, SUBLOCATION_X, SUBLOCATION_Y, SUBLOCATION_Z, UPDATE_USER_ID, UNPACK_TIME, CREATE_USER_ID, TALLY_ENABLE)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
