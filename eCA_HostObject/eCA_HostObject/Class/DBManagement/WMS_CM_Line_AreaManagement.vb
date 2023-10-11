Public Class WMS_CM_Line_AreaManagement
  Public Shared TableName As String = "WMS_CM_Line_Area"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer

    FACTORY_NO
    AREA_NO
    AREA_ID
    AREA_ALIS
    AREA_DESC
    AREA_TYPE2
    HIGH_WATER
    LOW_WATER
    DEVICE_NO
    ENABLE
    SHOW_INDEX
    SHOW_GROUP
    SHOW_COLOR
    Process_ID
    Process_CODE
    TB004
    TB005
    TB007
    TB008
    TB010
    PREVIOUS_AREA_NO
    AREA_INDEX
    AREA_TYPE1
    Report
  End Enum
  Public Shared Function GetUpdateSQL(ByRef Info As clsLine_Area) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}',{8}='{9}',{10}='{11}',{12}={13},{14}={15},{16}={17},{18}='{19}',{20}={21},{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}={39},{40}={41},{42}={43} WHERE {2}='{3}' And {4}='{5}'",
      strSQL,
      TableName,
      IdxColumnName.FACTORY_NO.ToString, Info.Factory_No,
      IdxColumnName.AREA_NO.ToString, Info.Area_No,
      IdxColumnName.AREA_ID.ToString, Info.Area_ID,
      IdxColumnName.AREA_ALIS.ToString, Info.Area_Alis,
      IdxColumnName.AREA_DESC.ToString, Info.Area_Desc,
      IdxColumnName.AREA_TYPE2.ToString, CInt(Info.Area_Type2),
      IdxColumnName.HIGH_WATER.ToString, Info.High_Water,
      IdxColumnName.LOW_WATER.ToString, Info.Low_Water,
      IdxColumnName.DEVICE_NO.ToString, Info.Device_No,
      IdxColumnName.ENABLE.ToString, BooleanConvertToInteger(Info.Enable),
      IdxColumnName.Process_ID.ToString, Info.Process_ID,
      IdxColumnName.Process_CODE.ToString, Info.Process_CODE,
      IdxColumnName.TB004.ToString, Info.TB004,
      IdxColumnName.TB005.ToString, Info.TB005,
      IdxColumnName.TB007.ToString, Info.TB007,
      IdxColumnName.TB008.ToString, Info.TB008,
      IdxColumnName.TB010.ToString, Info.TB010,
      IdxColumnName.PREVIOUS_AREA_NO.ToString, Info.PREVIOUS_AREA_NO,
      IdxColumnName.AREA_INDEX.ToString, Info.AREA_INDEX,
      IdxColumnName.AREA_TYPE1.ToString, CInt(Info.AREA_TYPE1),
      IdxColumnName.Report.ToString, BooleanConvertToInteger(Info.Report)
      )
      Return strSQL
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- GET
  Public Shared Function GetWMS_CM_Line_AreaDataListByALL() As List(Of clsLine_Area)
    Try
      Dim _lstReturn As New List(Of clsLine_Area)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsLine_Area = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info)
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsLine_Area, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim Factory_No = "" & RowData.Item(IdxColumnName.FACTORY_NO.ToString)
        Dim Area_No = "" & RowData.Item(IdxColumnName.AREA_NO.ToString)
        Dim Area_ID = "" & RowData.Item(IdxColumnName.AREA_ID.ToString)
        Dim Area_Alis = "" & RowData.Item(IdxColumnName.AREA_ALIS.ToString)
        Dim Area_Desc = "" & RowData.Item(IdxColumnName.AREA_DESC.ToString)
        Dim Area_Type = 0 & RowData.Item(IdxColumnName.AREA_TYPE2.ToString)
        Dim High_Water = 0 & RowData.Item(IdxColumnName.HIGH_WATER.ToString)
        Dim Low_Water = 0 & RowData.Item(IdxColumnName.LOW_WATER.ToString)
        Dim Device_No = "" & RowData.Item(IdxColumnName.DEVICE_NO.ToString)
        Dim Enable As Boolean = IntegerConvertToBoolean(RowData.Item(IdxColumnName.ENABLE.ToString))
        Dim Show_Index = 0 & RowData.Item(IdxColumnName.SHOW_INDEX.ToString)
        Dim Show_Group = 0 & RowData.Item(IdxColumnName.SHOW_GROUP.ToString)
        Dim Show_Color = "" & RowData.Item(IdxColumnName.SHOW_COLOR.ToString)

        Dim Process_ID = "" & RowData.Item(IdxColumnName.Process_ID.ToString)
        Dim Process_CODE = "" & RowData.Item(IdxColumnName.Process_CODE.ToString)
        Dim TB004 = "" & RowData.Item(IdxColumnName.TB004.ToString)
        Dim TB005 = "" & RowData.Item(IdxColumnName.TB005.ToString)
        Dim TB007 = "" & RowData.Item(IdxColumnName.TB007.ToString)
        Dim TB008 = "" & RowData.Item(IdxColumnName.TB008.ToString)
        Dim TB010 = "" & RowData.Item(IdxColumnName.TB010.ToString)

        Dim PREVIOUS_AREA_NO = "" & RowData.Item(IdxColumnName.PREVIOUS_AREA_NO.ToString)
        Dim AREA_INDEX = 0 & RowData.Item(IdxColumnName.AREA_INDEX.ToString)
        Dim AREA_TYPE1 = 0 & RowData.Item(IdxColumnName.AREA_TYPE1.ToString)
        Dim Report = IntegerConvertToBoolean(RowData.Item(IdxColumnName.Report.ToString))

        Info = New clsLine_Area(Factory_No, Area_No, Area_ID, Area_Alis, Area_Desc, Area_Type, High_Water,
                                                                 Low_Water, Device_No, Enable, Show_Index, Show_Group, Show_Color, Process_ID, Process_CODE, TB004,
                                                                TB005, TB007, TB008, TB010, PREVIOUS_AREA_NO, AREA_INDEX, AREA_TYPE1, Report)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
