Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_CM_Split_LabelManagement
    Public Shared TableName As String = "WMS_CM_Split_Label"
    Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsWMS_CM_Split_Label)
    Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
    Public Shared objLock As New Object
    Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
    Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

    Enum IdxColumnName As Integer
        COMPANY_CODE
        MATERIAL_NUMBER
        PLANT
        BATCH
        LAST_CHANGE_DATE
        BIN
        GP
        HF
        EXP_DATE
        LOT
        DC
        SAFETY
        ESD
        MFR
        MFRPN
        LABEL_COMMON1
        LABEL_COMMON2
        LABEL_COMMON3
        LABEL_COMMON4
        LABEL_COMMON5
        UPDATE_TIME
    End Enum

    Public Enum UpdateOption As Integer
        UpdateDic = 0
        UpdateDB = 1
    End Enum
    '- GetSQL
    Public Shared Function GetInsertSQL(ByRef CI As clsWMS_CM_Split_Label) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}')",
            strSQL,
            TableName,
            IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
            IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
            IdxColumnName.PLANT.ToString, CI.PLANT,
            IdxColumnName.BATCH.ToString, CI.BATCH,
            IdxColumnName.LAST_CHANGE_DATE.ToString, CI.LAST_CHANGE_DATE,
            IdxColumnName.BIN.ToString, CI.BIN,
            IdxColumnName.GP.ToString, CI.GP,
            IdxColumnName.HF.ToString, CI.HF,
            IdxColumnName.EXP_DATE.ToString, CI.EXP_DATE,
            IdxColumnName.LOT.ToString, CI.LOT,
            IdxColumnName.DC.ToString, CI.DC,
            IdxColumnName.SAFETY.ToString, CI.SAFETY,
            IdxColumnName.ESD.ToString, CI.ESD,
            IdxColumnName.MFR.ToString, CI.MFR,
            IdxColumnName.MFRPN.ToString, CI.MFRPN,
            IdxColumnName.LABEL_COMMON1.ToString, CI.LABEL_COMMON1,
            IdxColumnName.LABEL_COMMON2.ToString, CI.LABEL_COMMON2,
            IdxColumnName.LABEL_COMMON3.ToString, CI.LABEL_COMMON3,
            IdxColumnName.LABEL_COMMON4.ToString, CI.LABEL_COMMON4,
            IdxColumnName.LABEL_COMMON5.ToString, CI.LABEL_COMMON5,
            IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME
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
    Public Shared Function GetDeleteSQL(ByRef CI As clsWMS_CM_Split_Label) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
            strSQL,
            TableName,
            IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
            IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
            IdxColumnName.PLANT.ToString, CI.PLANT,
            IdxColumnName.BATCH.ToString, CI.BATCH,
            IdxColumnName.LAST_CHANGE_DATE.ToString, CI.LAST_CHANGE_DATE,
            IdxColumnName.BIN.ToString, CI.BIN,
            IdxColumnName.GP.ToString, CI.GP,
            IdxColumnName.HF.ToString, CI.HF,
            IdxColumnName.EXP_DATE.ToString, CI.EXP_DATE,
            IdxColumnName.LOT.ToString, CI.LOT,
            IdxColumnName.DC.ToString, CI.DC,
            IdxColumnName.SAFETY.ToString, CI.SAFETY,
            IdxColumnName.ESD.ToString, CI.ESD,
            IdxColumnName.MFR.ToString, CI.MFR,
            IdxColumnName.MFRPN.ToString, CI.MFRPN,
            IdxColumnName.LABEL_COMMON1.ToString, CI.LABEL_COMMON1,
            IdxColumnName.LABEL_COMMON2.ToString, CI.LABEL_COMMON2,
            IdxColumnName.LABEL_COMMON3.ToString, CI.LABEL_COMMON3,
            IdxColumnName.LABEL_COMMON4.ToString, CI.LABEL_COMMON4,
            IdxColumnName.LABEL_COMMON5.ToString, CI.LABEL_COMMON5,
            IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME
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
    Public Shared Function GetUpdateSQL(ByRef CI As clsWMS_CM_Split_Label) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
            strSQL,
            TableName,
            IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
            IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
            IdxColumnName.PLANT.ToString, CI.PLANT,
            IdxColumnName.BATCH.ToString, CI.BATCH,
            IdxColumnName.LAST_CHANGE_DATE.ToString, CI.LAST_CHANGE_DATE,
            IdxColumnName.BIN.ToString, CI.BIN,
            IdxColumnName.GP.ToString, CI.GP,
            IdxColumnName.HF.ToString, CI.HF,
            IdxColumnName.EXP_DATE.ToString, CI.EXP_DATE,
            IdxColumnName.LOT.ToString, CI.LOT,
            IdxColumnName.DC.ToString, CI.DC,
            IdxColumnName.SAFETY.ToString, CI.SAFETY,
            IdxColumnName.ESD.ToString, CI.ESD,
            IdxColumnName.MFR.ToString, CI.MFR,
            IdxColumnName.MFRPN.ToString, CI.MFRPN,
            IdxColumnName.LABEL_COMMON1.ToString, CI.LABEL_COMMON1,
            IdxColumnName.LABEL_COMMON2.ToString, CI.LABEL_COMMON2,
            IdxColumnName.LABEL_COMMON3.ToString, CI.LABEL_COMMON3,
            IdxColumnName.LABEL_COMMON4.ToString, CI.LABEL_COMMON4,
            IdxColumnName.LABEL_COMMON5.ToString, CI.LABEL_COMMON5,
            IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME
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
    Public Shared Function DeleteWMS_CM_Split_LabelData(ByVal Info As clsWMS_CM_Split_Label, Optional ByVal SendToDB As Boolean = True) As Boolean
        SyncLock objLock
            Try
                If Info Is Nothing Then Return False
                If DeletelstWMS_CM_Split_LabelData(New List(Of clsWMS_CM_Split_Label)({Info}), SendToDB) = True Then
                    Return True
                End If '-載不載入記憶體都是呼叫同一個function
                Return False
            Catch ex As Exception
                SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End Try
        End SyncLock
    End Function
    Public Shared Function DeletelstWMS_CM_Split_LabelData(ByVal Info As List(Of clsWMS_CM_Split_Label), Optional ByVal SendToDB As Boolean = True) As Boolean
        SyncLock objLock
            Try
                If Info Is Nothing Then Return False
                If Info.Count = 0 Then Return True
                If DictionaryNeeded = 1 Then '-載入記憶體
                    For i = 0 To Info.Count - 1
                        Dim key As String = Info(i).gid
                        If dicData.ContainsKey(key) = True Then
                            SendMessageToLog("There is no key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
                            Return False
                        End If
                    Next
                    If SendToDB Then
                        If DeleteWMS_CM_Split_LabelDataToDB(Info) Then
                            If DeleteWMS_CM_Split_LabelDataToDictionary(Info) Then
                                SendMessageToLog("DeleteDic WMS_CM_Split_LabelData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                            Else
                                SendMessageToLog("DeleteDic WMS_CM_Split_LabelData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                                Return False
                            End If
                        Else
                            SendMessageToLog("DeleteDB WMS_CM_Split_LabelData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                            Return False
                        End If
                    Else
                        If DeleteWMS_CM_Split_LabelDataToDictionary(Info) Then
                            SendMessageToLog("DeleteDic WMS_CM_Split_LabelData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                        Else
                            SendMessageToLog("DeleteDic WMS_CM_Split_LabelData Fail", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                            Return False
                        End If
                    End If
                Else
                    If SendToDB Then
                        If DeleteWMS_CM_Split_LabelDataToDB(Info) Then
                            Return True
                        Else
                            SendMessageToLog("DeleteDic WMS_CM_Split_LabelData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                            Return False
                        End If
                    Else
                        SendMessageToLog("Do Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                        Return True
                    End If
                End If
                Return True
            Catch ex As Exception
                SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End Try
        End SyncLock
    End Function

    Private Shared Function InsertWMS_CM_Split_LabelDataToDB(ByRef Info As List(Of clsWMS_CM_Split_Label)) As Integer
        Try
            If Info Is Nothing Then Return -1
            If Info.Count = 0 Then Return 0

            Dim strSQL As String = ""
            Dim rs As ADODB.Recordset = Nothing
            Dim lstSql As New List(Of String)
            For Each CI In Info
                strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}')",
                strSQL,
                TableName,
                IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
                IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
                IdxColumnName.PLANT.ToString, CI.PLANT,
                IdxColumnName.BATCH.ToString, CI.BATCH,
                IdxColumnName.LAST_CHANGE_DATE.ToString, CI.LAST_CHANGE_DATE,
                IdxColumnName.BIN.ToString, CI.BIN,
                IdxColumnName.GP.ToString, CI.GP,
                IdxColumnName.HF.ToString, CI.HF,
                IdxColumnName.EXP_DATE.ToString, CI.EXP_DATE,
                IdxColumnName.LOT.ToString, CI.LOT,
                IdxColumnName.DC.ToString, CI.DC,
                IdxColumnName.SAFETY.ToString, CI.SAFETY,
                IdxColumnName.ESD.ToString, CI.ESD,
                IdxColumnName.MFR.ToString, CI.MFR,
                IdxColumnName.MFRPN.ToString, CI.MFRPN,
                IdxColumnName.LABEL_COMMON1.ToString, CI.LABEL_COMMON1,
                IdxColumnName.LABEL_COMMON2.ToString, CI.LABEL_COMMON2,
                IdxColumnName.LABEL_COMMON3.ToString, CI.LABEL_COMMON3,
                IdxColumnName.LABEL_COMMON4.ToString, CI.LABEL_COMMON4,
                IdxColumnName.LABEL_COMMON5.ToString, CI.LABEL_COMMON5,
                IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME
                )
                lstSql.Add(strSQL)
            Next

            Dim NewSQL As New List(Of String)
            If SQLCorrect(lstSql, NewSQL) = False Then
                Return Nothing
            End If
            If SendSQLToDB(NewSQL) = True Then
                Return True
            Else
                SendMessageToLog("Insert to WMS_CM_Split_Label DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function UpdateWMS_CM_Split_LabelDataToDB(ByRef Info As List(Of clsWMS_CM_Split_Label)) As Integer
        Try
            If Info Is Nothing Then Return -1
            If Info.Count = 0 Then Return 0

            Dim strSQL As String = ""
            Dim rs As ADODB.Recordset = Nothing
            Dim lstSql As New List(Of String)
            For Each CI In Info
                strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
                strSQL,
                TableName,
                IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
                IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
                IdxColumnName.PLANT.ToString, CI.PLANT,
                IdxColumnName.BATCH.ToString, CI.BATCH,
                IdxColumnName.LAST_CHANGE_DATE.ToString, CI.LAST_CHANGE_DATE,
                IdxColumnName.BIN.ToString, CI.BIN,
                IdxColumnName.GP.ToString, CI.GP,
                IdxColumnName.HF.ToString, CI.HF,
                IdxColumnName.EXP_DATE.ToString, CI.EXP_DATE,
                IdxColumnName.LOT.ToString, CI.LOT,
                IdxColumnName.DC.ToString, CI.DC,
                IdxColumnName.SAFETY.ToString, CI.SAFETY,
                IdxColumnName.ESD.ToString, CI.ESD,
                IdxColumnName.MFR.ToString, CI.MFR,
                IdxColumnName.MFRPN.ToString, CI.MFRPN,
                IdxColumnName.LABEL_COMMON1.ToString, CI.LABEL_COMMON1,
                IdxColumnName.LABEL_COMMON2.ToString, CI.LABEL_COMMON2,
                IdxColumnName.LABEL_COMMON3.ToString, CI.LABEL_COMMON3,
                IdxColumnName.LABEL_COMMON4.ToString, CI.LABEL_COMMON4,
                IdxColumnName.LABEL_COMMON5.ToString, CI.LABEL_COMMON5,
                IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME
                )
                lstSql.Add(strSQL)
            Next

            Dim NewSQL As New List(Of String)
            If SQLCorrect(lstSql, NewSQL) = False Then
                Return Nothing
            End If
            If SendSQLToDB(NewSQL) = True Then
                Return True
            Else
                SendMessageToLog("Update to WMS_CM_Split_Label DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function DeleteWMS_CM_Split_LabelDataToDB(ByRef Info As List(Of clsWMS_CM_Split_Label)) As Integer
        Try
            If Info Is Nothing Then Return -1
            If Info.Count = 0 Then Return 0

            Dim strSQL As String = ""
            Dim rs As ADODB.Recordset = Nothing
            Dim lstSql As New List(Of String)
            For Each CI In Info
                strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
                strSQL,
                TableName,
                IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
                IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
                IdxColumnName.PLANT.ToString, CI.PLANT,
                IdxColumnName.BATCH.ToString, CI.BATCH,
                IdxColumnName.LAST_CHANGE_DATE.ToString, CI.LAST_CHANGE_DATE,
                IdxColumnName.BIN.ToString, CI.BIN,
                IdxColumnName.GP.ToString, CI.GP,
                IdxColumnName.HF.ToString, CI.HF,
                IdxColumnName.EXP_DATE.ToString, CI.EXP_DATE,
                IdxColumnName.LOT.ToString, CI.LOT,
                IdxColumnName.DC.ToString, CI.DC,
                IdxColumnName.SAFETY.ToString, CI.SAFETY,
                IdxColumnName.ESD.ToString, CI.ESD,
                IdxColumnName.MFR.ToString, CI.MFR,
                IdxColumnName.MFRPN.ToString, CI.MFRPN,
                IdxColumnName.LABEL_COMMON1.ToString, CI.LABEL_COMMON1,
                IdxColumnName.LABEL_COMMON2.ToString, CI.LABEL_COMMON2,
                IdxColumnName.LABEL_COMMON3.ToString, CI.LABEL_COMMON3,
                IdxColumnName.LABEL_COMMON4.ToString, CI.LABEL_COMMON4,
                IdxColumnName.LABEL_COMMON5.ToString, CI.LABEL_COMMON5,
                IdxColumnName.UPDATE_TIME.ToString, CI.UPDATE_TIME
                )
                lstSql.Add(strSQL)
            Next

            Dim NewSQL As New List(Of String)
            If SQLCorrect(lstSql, NewSQL) = False Then
                Return Nothing
            End If
            If SendSQLToDB(NewSQL) = True Then
                Return True
            Else
                SendMessageToLog("Delete to WMS_CM_Split_Label DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function

    Private Shared Function DeleteWMS_CM_Split_LabelDataToDictionary(ByRef Info As List(Of clsWMS_CM_Split_Label)) As Boolean
        Try
            If Info Is Nothing Then Return False
            If Info.Count = 0 Then Return True
            For i = 0 To Info.Count - 1
                Dim key As String = Info(i).gid
                If dicData.TryRemove(key, Nothing) = False Then
                    SendMessageToLog("dicData.TryRemove Failed -WMS_CM_Split_LabelData", eCALogTool.ILogTool.enuTrcLevel.lvError)
                End If
            Next
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function SetInfoFromDB(ByRef Info As clsWMS_CM_Split_Label, ByRef RowData As DataRow) As Boolean
        Try
            If RowData IsNot Nothing Then
                Dim COMPANY_CODE = "" & RowData.Item(IdxColumnName.COMPANY_CODE.ToString)
                Dim MATERIAL_NUMBER = "" & RowData.Item(IdxColumnName.MATERIAL_NUMBER.ToString)
                Dim PLANT = "" & RowData.Item(IdxColumnName.PLANT.ToString)
                Dim BATCH = "" & RowData.Item(IdxColumnName.BATCH.ToString)
                Dim LAST_CHANGE_DATE = "" & RowData.Item(IdxColumnName.LAST_CHANGE_DATE.ToString)
                Dim BIN = "" & RowData.Item(IdxColumnName.BIN.ToString)
                Dim GP = "" & RowData.Item(IdxColumnName.GP.ToString)
                Dim HF = "" & RowData.Item(IdxColumnName.HF.ToString)
                Dim EXP_DATE = "" & RowData.Item(IdxColumnName.EXP_DATE.ToString)
                Dim LOT = "" & RowData.Item(IdxColumnName.LOT.ToString)
                Dim DC = "" & RowData.Item(IdxColumnName.DC.ToString)
                Dim SAFETY = "" & RowData.Item(IdxColumnName.SAFETY.ToString)
                Dim ESD = "" & RowData.Item(IdxColumnName.ESD.ToString)
                Dim MFR = "" & RowData.Item(IdxColumnName.MFR.ToString)
                Dim MFRPN = "" & RowData.Item(IdxColumnName.MFRPN.ToString)
                Dim LABEL_COMMON1 = "" & RowData.Item(IdxColumnName.LABEL_COMMON1.ToString)
                Dim LABEL_COMMON2 = "" & RowData.Item(IdxColumnName.LABEL_COMMON2.ToString)
                Dim LABEL_COMMON3 = "" & RowData.Item(IdxColumnName.LABEL_COMMON3.ToString)
                Dim LABEL_COMMON4 = "" & RowData.Item(IdxColumnName.LABEL_COMMON4.ToString)
                Dim LABEL_COMMON5 = "" & RowData.Item(IdxColumnName.LABEL_COMMON5.ToString)
                Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
                Info = New clsWMS_CM_Split_Label(COMPANY_CODE, MATERIAL_NUMBER, PLANT, BATCH, LAST_CHANGE_DATE, BIN, GP, HF, EXP_DATE, LOT, DC, SAFETY, ESD, MFR, MFRPN, LABEL_COMMON1, LABEL_COMMON2, LABEL_COMMON3, LABEL_COMMON4, LABEL_COMMON5, UPDATE_TIME)

            End If
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function SendSQLToDB(ByRef lstSQL As List(Of String)) As Boolean
        Try
            If lstSQL Is Nothing Then Return False
            If lstSQL.Count = 0 Then Return True
            For i = 0 To lstSQL.Count - 1
                SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            Next
            If fUseBatchUpdate_DynamicConnection = 0 Then
                For i = 0 To lstSQL.Count - 1
                    DBTool.O_AddSQLQueue(TableName, lstSQL(i))
                Next
            Else
                Dim rtnMsg As String = DBTool.BatchUpdate_DynamicConnection(lstSQL)
                If rtnMsg.StartsWith("OK") Then
                    SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                Else
                    SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    '- GET
    Public Shared Function GetdicSplit_LabelByALL() As Dictionary(Of String, clsWMS_CM_Split_Label)
        Try
            Dim dicReturn As New Dictionary(Of String, clsWMS_CM_Split_Label)
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
                        Dim Info As clsWMS_CM_Split_Label = Nothing
                        SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                        If dicReturn.ContainsKey(Info.gid) = False Then
                            dicReturn.Add(Info.gid, Info)
                        End If
                    Next
                End If
                'End If
            End If
            Return dicReturn
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return Nothing
        End Try
    End Function
End Class
