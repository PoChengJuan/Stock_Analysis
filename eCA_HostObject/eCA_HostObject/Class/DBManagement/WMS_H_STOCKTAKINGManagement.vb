Partial Class WMS_H_STOCKTAKINGManagement
Public Shared TableName As String = "WMS_H_STOCKTAKING"
Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
 
Enum IdxColumnName As Integer
STOCKTAKING_ID
STOCKTAKING_TYPE1
STOCKTAKING_TYPE2
STOCKTAKING_TYPE3
CREATE_TIME
START_TIME
FINISH_TIME
CREATE_USER
STATUS
LOCATION_GROUP_NO
PRIORITY
CARRIER_QTY
CARRIER_QTY_CHECKED
MATCH_TYPE
SEND_TO_HOST
CHANGE_INVENTORY
UPLOAD_STATUS
UPLOAD_COMMENTS
HIST_TIME
End Enum
'- GetSQL
 Public Shared Function GetInsertSQL(ByRef Info As  clsHSTOCKTAKING) As String
	Try
 
 Dim strSQL As String = ""
 strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38}) values ('{3}',{5},{7},'{9}','{11}','{13}','{15}','{17}',{19},'{21}',{23},{25},{27},{29},{31},{33},{35},'{37}','{39}')",
 strSQL,
 TableName,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_TYPE1.ToString, CINT(Info.STOCKTAKING_TYPE1),
 IdxColumnName.STOCKTAKING_TYPE2.ToString, CINT(Info.STOCKTAKING_TYPE2),
 IdxColumnName.STOCKTAKING_TYPE3.ToString, CINT(Info.STOCKTAKING_TYPE3),
 IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
 IdxColumnName.START_TIME.ToString, Info.START_TIME,
 IdxColumnName.FINISH_TIME.ToString, Info.FINISH_TIME,
 IdxColumnName.CREATE_USER.ToString, Info.CREATE_USER,
 IdxColumnName.STATUS.ToString, CINT(Info.STATUS),
 IdxColumnName.LOCATION_GROUP_NO.ToString, Info.LOCATION_GROUP_NO,
 IdxColumnName.PRIORITY.ToString, Info.PRIORITY,
 IdxColumnName.CARRIER_QTY.ToString, Info.CARRIER_QTY,
 IdxColumnName.CARRIER_QTY_CHECKED.ToString, Info.CARRIER_QTY_CHECKED,
 IdxColumnName.MATCH_TYPE.ToString, CINT(Info.MATCH_TYPE),
 IdxColumnName.SEND_TO_HOST.ToString, Info.SEND_TO_HOST,
 IdxColumnName.CHANGE_INVENTORY.ToString, Info.CHANGE_INVENTORY,
 IdxColumnName.UPLOAD_STATUS.ToString, Info.UPLOAD_STATUS,
 IdxColumnName.UPLOAD_COMMENTS.ToString, Info.UPLOAD_COMMENTS,
 IdxColumnName.HIST_TIME.ToString, Info.HIST_TIME
)
 Dim NewSQL As String = ""
 If SQLCorrect(strSQL, NewSQL) Then
 Return NewSQL
 End If
 Return Nothing
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return nothing
 End Try
 End Function
 Public Shared Function GetUpdateSQL(ByRef Info As clsHSTOCKTAKING) As String
	Try
 Dim strSQL As String = ""
 strSQL = String.Format("Update {1} SET {4}={5},{6}={7},{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}={19},{20}='{21}',{22}={23},{24}={25},{26}={27},{28}={29},{30}={31},{32}={33},{34}={35},{36}='{37}',{38}='{39}' WHERE {2}='{3}'",
 strSQL,
 TableName,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_TYPE1.ToString, CINT(Info.STOCKTAKING_TYPE1),
 IdxColumnName.STOCKTAKING_TYPE2.ToString, CINT(Info.STOCKTAKING_TYPE2),
 IdxColumnName.STOCKTAKING_TYPE3.ToString, CINT(Info.STOCKTAKING_TYPE3),
 IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
 IdxColumnName.START_TIME.ToString, Info.START_TIME,
 IdxColumnName.FINISH_TIME.ToString, Info.FINISH_TIME,
 IdxColumnName.CREATE_USER.ToString, Info.CREATE_USER,
 IdxColumnName.STATUS.ToString, CINT(Info.STATUS),
 IdxColumnName.LOCATION_GROUP_NO.ToString, Info.LOCATION_GROUP_NO,
 IdxColumnName.PRIORITY.ToString, Info.PRIORITY,
 IdxColumnName.CARRIER_QTY.ToString, Info.CARRIER_QTY,
 IdxColumnName.CARRIER_QTY_CHECKED.ToString, Info.CARRIER_QTY_CHECKED,
 IdxColumnName.MATCH_TYPE.ToString, CINT(Info.MATCH_TYPE),
 IdxColumnName.SEND_TO_HOST.ToString, Info.SEND_TO_HOST,
 IdxColumnName.CHANGE_INVENTORY.ToString, Info.CHANGE_INVENTORY,
 IdxColumnName.UPLOAD_STATUS.ToString, Info.UPLOAD_STATUS,
 IdxColumnName.UPLOAD_COMMENTS.ToString, Info.UPLOAD_COMMENTS,
 IdxColumnName.HIST_TIME.ToString, Info.HIST_TIME
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
 Public Shared Function GetDeleteSQL(ByRef Info As clsHSTOCKTAKING) As String
	Try
 Dim strSQL As String = ""
 strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
 strSQL,
 TableName,
 IdxColumnName.STOCKTAKING_ID.ToString, Info.STOCKTAKING_ID,
 IdxColumnName.STOCKTAKING_TYPE1.ToString, CINT(Info.STOCKTAKING_TYPE1),
 IdxColumnName.STOCKTAKING_TYPE2.ToString, CINT(Info.STOCKTAKING_TYPE2),
 IdxColumnName.STOCKTAKING_TYPE3.ToString, CINT(Info.STOCKTAKING_TYPE3),
 IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
 IdxColumnName.START_TIME.ToString, Info.START_TIME,
 IdxColumnName.FINISH_TIME.ToString, Info.FINISH_TIME,
 IdxColumnName.CREATE_USER.ToString, Info.CREATE_USER,
 IdxColumnName.STATUS.ToString, CINT(Info.STATUS),
 IdxColumnName.LOCATION_GROUP_NO.ToString, Info.LOCATION_GROUP_NO,
 IdxColumnName.PRIORITY.ToString, Info.PRIORITY,
 IdxColumnName.CARRIER_QTY.ToString, Info.CARRIER_QTY,
 IdxColumnName.CARRIER_QTY_CHECKED.ToString, Info.CARRIER_QTY_CHECKED,
 IdxColumnName.MATCH_TYPE.ToString, CINT(Info.MATCH_TYPE),
 IdxColumnName.SEND_TO_HOST.ToString, Info.SEND_TO_HOST,
 IdxColumnName.CHANGE_INVENTORY.ToString, Info.CHANGE_INVENTORY,
 IdxColumnName.UPLOAD_STATUS.ToString, Info.UPLOAD_STATUS,
 IdxColumnName.UPLOAD_COMMENTS.ToString, Info.UPLOAD_COMMENTS,
 IdxColumnName.HIST_TIME.ToString, Info.HIST_TIME
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
Private Shared Function SetInfoFromDB(ByRef Info As clsHSTOCKTAKING , ByRef RowData As DataRow) As Boolean
 Try
If RowData IsNot Nothing Then
Dim STOCKTAKING_ID=""&RowData.Item(IdxColumnName.STOCKTAKING_ID.ToString)
Dim STOCKTAKING_TYPE1=If(IsNumeric(RowData.Item(IdxColumnName.STOCKTAKING_TYPE1.ToString))  ,RowData.Item(IdxColumnName.STOCKTAKING_TYPE1.ToString), 0 & RowData.Item(IdxColumnName.STOCKTAKING_TYPE1.ToString))
Dim STOCKTAKING_TYPE2=If(IsNumeric(RowData.Item(IdxColumnName.STOCKTAKING_TYPE2.ToString))  ,RowData.Item(IdxColumnName.STOCKTAKING_TYPE2.ToString), 0 & RowData.Item(IdxColumnName.STOCKTAKING_TYPE2.ToString))
Dim STOCKTAKING_TYPE3=""&RowData.Item(IdxColumnName.STOCKTAKING_TYPE3.ToString)
Dim CREATE_TIME=""&RowData.Item(IdxColumnName.CREATE_TIME.ToString)
Dim START_TIME=""&RowData.Item(IdxColumnName.START_TIME.ToString)
Dim FINISH_TIME=""&RowData.Item(IdxColumnName.FINISH_TIME.ToString)
Dim CREATE_USER=""&RowData.Item(IdxColumnName.CREATE_USER.ToString)
Dim STATUS=If(IsNumeric(RowData.Item(IdxColumnName.STATUS.ToString))  ,RowData.Item(IdxColumnName.STATUS.ToString), 0 & RowData.Item(IdxColumnName.STATUS.ToString))
Dim LOCATION_GROUP_NO=""&RowData.Item(IdxColumnName.LOCATION_GROUP_NO.ToString)
Dim PRIORITY=If(IsNumeric(RowData.Item(IdxColumnName.PRIORITY.ToString))  ,RowData.Item(IdxColumnName.PRIORITY.ToString), 0 & RowData.Item(IdxColumnName.PRIORITY.ToString))
Dim CARRIER_QTY=If(IsNumeric(RowData.Item(IdxColumnName.CARRIER_QTY.ToString))  ,RowData.Item(IdxColumnName.CARRIER_QTY.ToString), 0 & RowData.Item(IdxColumnName.CARRIER_QTY.ToString))
Dim CARRIER_QTY_CHECKED=If(IsNumeric(RowData.Item(IdxColumnName.CARRIER_QTY_CHECKED.ToString))  ,RowData.Item(IdxColumnName.CARRIER_QTY_CHECKED.ToString), 0 & RowData.Item(IdxColumnName.CARRIER_QTY_CHECKED.ToString))
Dim MATCH_TYPE=If(IsNumeric(RowData.Item(IdxColumnName.MATCH_TYPE.ToString))  ,RowData.Item(IdxColumnName.MATCH_TYPE.ToString), 0 & RowData.Item(IdxColumnName.MATCH_TYPE.ToString))
Dim SEND_TO_HOST=If(IsNumeric(RowData.Item(IdxColumnName.SEND_TO_HOST.ToString))  ,RowData.Item(IdxColumnName.SEND_TO_HOST.ToString), 0 & RowData.Item(IdxColumnName.SEND_TO_HOST.ToString))
Dim CHANGE_INVENTORY=If(IsNumeric(RowData.Item(IdxColumnName.CHANGE_INVENTORY.ToString))  ,RowData.Item(IdxColumnName.CHANGE_INVENTORY.ToString), 0 & RowData.Item(IdxColumnName.CHANGE_INVENTORY.ToString))
Dim UPLOAD_STATUS=If(IsNumeric(RowData.Item(IdxColumnName.UPLOAD_STATUS.ToString))  ,RowData.Item(IdxColumnName.UPLOAD_STATUS.ToString), 0 & RowData.Item(IdxColumnName.UPLOAD_STATUS.ToString))
Dim UPLOAD_COMMENTS=""&RowData.Item(IdxColumnName.UPLOAD_COMMENTS.ToString)
Dim HIST_TIME=""&RowData.Item(IdxColumnName.HIST_TIME.ToString)
 Info = New clsHSTOCKTAKING(STOCKTAKING_ID,STOCKTAKING_TYPE1,STOCKTAKING_TYPE2,STOCKTAKING_TYPE3,CREATE_TIME,START_TIME,FINISH_TIME,CREATE_USER,STATUS,LOCATION_GROUP_NO,PRIORITY,CARRIER_QTY,CARRIER_QTY_CHECKED,MATCH_TYPE,SEND_TO_HOST,CHANGE_INVENTORY,UPLOAD_STATUS,UPLOAD_COMMENTS,HIST_TIME)
 
End If
Return true
Catch ex As Exception
SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
Return false
End Try
End Function
  Public Shared Function GetWMS_H_STOCKTAKINGdicByStocktaking_ID(ByVal STOCKTAKING_ID As String) As Dictionary(Of String, clsHSTOCKTAKING)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsHSTOCKTAKING)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {1} WHERE {2}='{3}'",
 strSQL,
 TableName, IdxColumnName.STOCKTAKING_ID.ToString,
 STOCKTAKING_ID
 )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsHSTOCKTAKING = Nothing
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
End Class
