Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_CT_VCManagement
    Public Shared TableName As String = "WMS_CT_VC"
    Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsWMS_CT_VC)
    Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
    Public Shared objLock As New Object
    Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
    Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

    Enum IdxColumnName As Integer
        COMPANY_CODE
        MATERIAL_DOCUMENT_YEAR
        MATERIAL_DOCUMENT
        MATERIAL_DOCUMENT_ITEM
        MATERIAL_NUMBER
        PLANT
        STORAGE_LOCATION
        BATCH
        QUANTITY
        CREATED_DATE
        CREATE_TIME
        VC_COMMON1
        VC_COMMON2
        VC_COMMON3
        VC_COMMON4
        VC_COMMON5
    End Enum

    Public Enum UpdateOption As Integer
        UpdateDic = 0
        UpdateDB = 1
    End Enum
    '- GetSQL
    Public Shared Function GetInsertSQL(ByRef CI As clsWMS_CT_VC) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}')",
            strSQL,
            TableName,
            IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
            IdxColumnName.MATERIAL_DOCUMENT_YEAR.ToString, CI.MATERIAL_DOCUMENT_YEAR,
            IdxColumnName.MATERIAL_DOCUMENT.ToString, CI.MATERIAL_DOCUMENT,
            IdxColumnName.MATERIAL_DOCUMENT_ITEM.ToString, CI.MATERIAL_DOCUMENT_ITEM,
            IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
            IdxColumnName.PLANT.ToString, CI.PLANT,
            IdxColumnName.STORAGE_LOCATION.ToString, CI.STORAGE_LOCATION,
            IdxColumnName.BATCH.ToString, CI.BATCH,
            IdxColumnName.QUANTITY.ToString, CI.QUANTITY,
            IdxColumnName.CREATED_DATE.ToString, CI.CREATED_DATE,
            IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
            IdxColumnName.VC_COMMON1.ToString, CI.VC_COMMON1,
            IdxColumnName.VC_COMMON2.ToString, CI.VC_COMMON2,
            IdxColumnName.VC_COMMON3.ToString, CI.VC_COMMON3,
            IdxColumnName.VC_COMMON4.ToString, CI.VC_COMMON4,
            IdxColumnName.VC_COMMON5.ToString, CI.VC_COMMON5
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
    Public Shared Function GetDeleteSQL(ByRef CI As clsWMS_CT_VC) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' AND {8}='{9}' ",
            strSQL,
            TableName,
            IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
            IdxColumnName.MATERIAL_DOCUMENT_YEAR.ToString, CI.MATERIAL_DOCUMENT_YEAR,
            IdxColumnName.MATERIAL_DOCUMENT.ToString, CI.MATERIAL_DOCUMENT,
            IdxColumnName.MATERIAL_DOCUMENT_ITEM.ToString, CI.MATERIAL_DOCUMENT_ITEM,
            IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
            IdxColumnName.PLANT.ToString, CI.PLANT,
            IdxColumnName.STORAGE_LOCATION.ToString, CI.STORAGE_LOCATION,
            IdxColumnName.BATCH.ToString, CI.BATCH,
            IdxColumnName.QUANTITY.ToString, CI.QUANTITY,
            IdxColumnName.CREATED_DATE.ToString, CI.CREATED_DATE,
            IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
            IdxColumnName.VC_COMMON1.ToString, CI.VC_COMMON1,
            IdxColumnName.VC_COMMON2.ToString, CI.VC_COMMON2,
            IdxColumnName.VC_COMMON3.ToString, CI.VC_COMMON3,
            IdxColumnName.VC_COMMON4.ToString, CI.VC_COMMON4,
            IdxColumnName.VC_COMMON5.ToString, CI.VC_COMMON5
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
    Public Shared Function GetUpdateSQL(ByRef CI As clsWMS_CT_VC) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
            strSQL,
            TableName,
            IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
            IdxColumnName.MATERIAL_DOCUMENT_YEAR.ToString, CI.MATERIAL_DOCUMENT_YEAR,
            IdxColumnName.MATERIAL_DOCUMENT.ToString, CI.MATERIAL_DOCUMENT,
            IdxColumnName.MATERIAL_DOCUMENT_ITEM.ToString, CI.MATERIAL_DOCUMENT_ITEM,
            IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
            IdxColumnName.PLANT.ToString, CI.PLANT,
            IdxColumnName.STORAGE_LOCATION.ToString, CI.STORAGE_LOCATION,
            IdxColumnName.BATCH.ToString, CI.BATCH,
            IdxColumnName.QUANTITY.ToString, CI.QUANTITY,
            IdxColumnName.CREATED_DATE.ToString, CI.CREATED_DATE,
            IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
            IdxColumnName.VC_COMMON1.ToString, CI.VC_COMMON1,
            IdxColumnName.VC_COMMON2.ToString, CI.VC_COMMON2,
            IdxColumnName.VC_COMMON3.ToString, CI.VC_COMMON3,
            IdxColumnName.VC_COMMON4.ToString, CI.VC_COMMON4,
            IdxColumnName.VC_COMMON5.ToString, CI.VC_COMMON5
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
    Private Shared Function InsertWMS_CT_VCDataToDB(ByRef Info As List(Of clsWMS_CT_VC)) As Integer
        Try
            If Info Is Nothing Then Return -1
            If Info.Count = 0 Then Return 0

            Dim strSQL As String = ""
            Dim rs As ADODB.Recordset = Nothing
            Dim lstSql As New List(Of String)
            For Each CI In Info
                strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}')",
                strSQL,
                TableName,
                IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
                IdxColumnName.MATERIAL_DOCUMENT_YEAR.ToString, CI.MATERIAL_DOCUMENT_YEAR,
                IdxColumnName.MATERIAL_DOCUMENT.ToString, CI.MATERIAL_DOCUMENT,
                IdxColumnName.MATERIAL_DOCUMENT_ITEM.ToString, CI.MATERIAL_DOCUMENT_ITEM,
                IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
                IdxColumnName.PLANT.ToString, CI.PLANT,
                IdxColumnName.STORAGE_LOCATION.ToString, CI.STORAGE_LOCATION,
                IdxColumnName.BATCH.ToString, CI.BATCH,
                IdxColumnName.QUANTITY.ToString, CI.QUANTITY,
                IdxColumnName.CREATED_DATE.ToString, CI.CREATED_DATE,
                IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
                IdxColumnName.VC_COMMON1.ToString, CI.VC_COMMON1,
                IdxColumnName.VC_COMMON2.ToString, CI.VC_COMMON2,
                IdxColumnName.VC_COMMON3.ToString, CI.VC_COMMON3,
                IdxColumnName.VC_COMMON4.ToString, CI.VC_COMMON4,
                IdxColumnName.VC_COMMON5.ToString, CI.VC_COMMON5
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
                SendMessageToLog("Insert to WMS_CT_VC DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function UpdateWMS_CT_VCDataToDB(ByRef Info As List(Of clsWMS_CT_VC)) As Integer
        Try
            If Info Is Nothing Then Return -1
            If Info.Count = 0 Then Return 0

            Dim strSQL As String = ""
            Dim rs As ADODB.Recordset = Nothing
            Dim lstSql As New List(Of String)
            For Each CI In Info
                strSQL = String.Format("Update {1} SET {10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}' WHERE {2}='{3}' And {4}='{5}' And {6}='{7}' And {8}='{9}'",
                strSQL,
                TableName,
                IdxColumnName.COMPANY_CODE.ToString, CI.COMPANY_CODE,
                IdxColumnName.MATERIAL_DOCUMENT_YEAR.ToString, CI.MATERIAL_DOCUMENT_YEAR,
                IdxColumnName.MATERIAL_DOCUMENT.ToString, CI.MATERIAL_DOCUMENT,
                IdxColumnName.MATERIAL_DOCUMENT_ITEM.ToString, CI.MATERIAL_DOCUMENT_ITEM,
                IdxColumnName.MATERIAL_NUMBER.ToString, CI.MATERIAL_NUMBER,
                IdxColumnName.PLANT.ToString, CI.PLANT,
                IdxColumnName.STORAGE_LOCATION.ToString, CI.STORAGE_LOCATION,
                IdxColumnName.BATCH.ToString, CI.BATCH,
                IdxColumnName.QUANTITY.ToString, CI.QUANTITY,
                IdxColumnName.CREATED_DATE.ToString, CI.CREATED_DATE,
                IdxColumnName.CREATE_TIME.ToString, CI.CREATE_TIME,
                IdxColumnName.VC_COMMON1.ToString, CI.VC_COMMON1,
                IdxColumnName.VC_COMMON2.ToString, CI.VC_COMMON2,
                IdxColumnName.VC_COMMON3.ToString, CI.VC_COMMON3,
                IdxColumnName.VC_COMMON4.ToString, CI.VC_COMMON4,
                IdxColumnName.VC_COMMON5.ToString, CI.VC_COMMON5
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
                SendMessageToLog("Update to WMS_CT_VC DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function

    Private Shared Function DeleteWMS_CT_VCDataToDictionary(ByRef Info As List(Of clsWMS_CT_VC)) As Boolean
        Try
            If Info Is Nothing Then Return False
            If Info.Count = 0 Then Return True
            For i = 0 To Info.Count - 1
                Dim key As String = Info(i).gid
                If dicData.TryRemove(key, Nothing) = False Then
                    SendMessageToLog("dicData.TryRemove Failed -WMS_CT_VCData", eCALogTool.ILogTool.enuTrcLevel.lvError)
                End If
            Next
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function SetInfoFromDB(ByRef Info As clsWMS_CT_VC, ByRef RowData As DataRow) As Boolean
        Try
            If RowData IsNot Nothing Then
                Dim COMPANY_CODE = "" & RowData.Item(IdxColumnName.COMPANY_CODE.ToString)
                Dim MATERIAL_DOCUMENT_YEAR = "" & RowData.Item(IdxColumnName.MATERIAL_DOCUMENT_YEAR.ToString)
                Dim MATERIAL_DOCUMENT = "" & RowData.Item(IdxColumnName.MATERIAL_DOCUMENT.ToString)
                Dim MATERIAL_DOCUMENT_ITEM = "" & RowData.Item(IdxColumnName.MATERIAL_DOCUMENT_ITEM.ToString)
                Dim MATERIAL_NUMBER = "" & RowData.Item(IdxColumnName.MATERIAL_NUMBER.ToString)
                Dim PLANT = "" & RowData.Item(IdxColumnName.PLANT.ToString)
                Dim STORAGE_LOCATION = "" & RowData.Item(IdxColumnName.STORAGE_LOCATION.ToString)
                Dim BATCH = "" & RowData.Item(IdxColumnName.BATCH.ToString)
                Dim QUANTITY = "" & RowData.Item(IdxColumnName.QUANTITY.ToString)
                Dim CREATED_DATE = "" & RowData.Item(IdxColumnName.CREATED_DATE.ToString)
                Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
                Dim VC_COMMON1 = "" & RowData.Item(IdxColumnName.VC_COMMON1.ToString)
                Dim VC_COMMON2 = "" & RowData.Item(IdxColumnName.VC_COMMON2.ToString)
                Dim VC_COMMON3 = "" & RowData.Item(IdxColumnName.VC_COMMON3.ToString)
                Dim VC_COMMON4 = "" & RowData.Item(IdxColumnName.VC_COMMON4.ToString)
                Dim VC_COMMON5 = "" & RowData.Item(IdxColumnName.VC_COMMON5.ToString)
                Info = New clsWMS_CT_VC(COMPANY_CODE, MATERIAL_DOCUMENT_YEAR, MATERIAL_DOCUMENT, MATERIAL_DOCUMENT_ITEM, MATERIAL_NUMBER, PLANT, STORAGE_LOCATION, BATCH, QUANTITY, CREATED_DATE, CREATE_TIME, VC_COMMON1, VC_COMMON2, VC_COMMON3, VC_COMMON4, VC_COMMON5)

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
    Public Shared Function GetdicVCDataByALL() As Dictionary(Of String, clsWMS_CT_VC)
        Try
            Dim dicReturn As New Dictionary(Of String, clsWMS_CT_VC)
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
                        Dim Info As clsWMS_CT_VC = Nothing
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
