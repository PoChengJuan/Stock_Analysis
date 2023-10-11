Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class WMS_M_OwnerManagement
  Public Shared TableName As String = "WMS_M_OWNER"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsOwner)
    Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
    Public Shared objLock As New Object
    Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
    Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

    Enum IdxColumnName As Integer
        OWNER_NO
        OWNER_ID1
        OWNER_ID2
        OWNER_ID3
        OWNER_ALIS1
        OWNER_ALIS2
        OWNER_DESC
        OWNER_TYPE
        OWNER_COMMON1
        OWNER_COMMON2
        OWNER_COMMON3
        OWNER_COMMON4
        OWNER_COMMON5
        OWNER_COMMON6
        OWNER_COMMON7
        OWNER_COMMON8
        OWNER_COMMON9
        OWNER_COMMON10
        COMMENTS
        ENABLE
    End Enum

    Public Enum UpdateOption As Integer
        UpdateDic = 0
        UpdateDB = 1
    End Enum
    '- GetSQL
    Public Shared Function GetInsertSQL(ByRef CI As clsOwner) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}',{17},'{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}',{41})",
            strSQL,
            TableName,
            IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
            IdxColumnName.OWNER_ID1.ToString, CI.OWNER_ID1,
            IdxColumnName.OWNER_ID2.ToString, CI.OWNER_ID2,
            IdxColumnName.OWNER_ID3.ToString, CI.OWNER_ID3,
            IdxColumnName.OWNER_ALIS1.ToString, CI.OWNER_ALIS1,
            IdxColumnName.OWNER_ALIS2.ToString, CI.OWNER_ALIS2,
            IdxColumnName.OWNER_DESC.ToString, CI.OWNER_DESC,
            IdxColumnName.OWNER_TYPE.ToString, CI.OWNER_TYPE,
            IdxColumnName.OWNER_COMMON1.ToString, CI.OWNER_COMMON1,
            IdxColumnName.OWNER_COMMON2.ToString, CI.OWNER_COMMON2,
            IdxColumnName.OWNER_COMMON3.ToString, CI.OWNER_COMMON3,
            IdxColumnName.OWNER_COMMON4.ToString, CI.OWNER_COMMON4,
            IdxColumnName.OWNER_COMMON5.ToString, CI.OWNER_COMMON5,
            IdxColumnName.OWNER_COMMON6.ToString, CI.OWNER_COMMON6,
            IdxColumnName.OWNER_COMMON7.ToString, CI.OWNER_COMMON7,
            IdxColumnName.OWNER_COMMON8.ToString, CI.OWNER_COMMON8,
            IdxColumnName.OWNER_COMMON9.ToString, CI.OWNER_COMMON9,
            IdxColumnName.OWNER_COMMON10.ToString, CI.OWNER_COMMON10,
            IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
            IdxColumnName.ENABLE.ToString, CI.ENABLE
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
    Public Shared Function GetDeleteSQL(ByRef CI As clsOwner) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
            strSQL,
            TableName,
            IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
            IdxColumnName.OWNER_ID1.ToString, CI.OWNER_ID1,
            IdxColumnName.OWNER_ID2.ToString, CI.OWNER_ID2,
            IdxColumnName.OWNER_ID3.ToString, CI.OWNER_ID3,
            IdxColumnName.OWNER_ALIS1.ToString, CI.OWNER_ALIS1,
            IdxColumnName.OWNER_ALIS2.ToString, CI.OWNER_ALIS2,
            IdxColumnName.OWNER_DESC.ToString, CI.OWNER_DESC,
            IdxColumnName.OWNER_TYPE.ToString, CI.OWNER_TYPE,
            IdxColumnName.OWNER_COMMON1.ToString, CI.OWNER_COMMON1,
            IdxColumnName.OWNER_COMMON2.ToString, CI.OWNER_COMMON2,
            IdxColumnName.OWNER_COMMON3.ToString, CI.OWNER_COMMON3,
            IdxColumnName.OWNER_COMMON4.ToString, CI.OWNER_COMMON4,
            IdxColumnName.OWNER_COMMON5.ToString, CI.OWNER_COMMON5,
            IdxColumnName.OWNER_COMMON6.ToString, CI.OWNER_COMMON6,
            IdxColumnName.OWNER_COMMON7.ToString, CI.OWNER_COMMON7,
            IdxColumnName.OWNER_COMMON8.ToString, CI.OWNER_COMMON8,
            IdxColumnName.OWNER_COMMON9.ToString, CI.OWNER_COMMON9,
            IdxColumnName.OWNER_COMMON10.ToString, CI.OWNER_COMMON10,
            IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
            IdxColumnName.ENABLE.ToString, CI.ENABLE
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
    Public Shared Function GetUpdateSQL(ByRef CI As clsOwner) As String
        Try
            Dim strSQL As String = ""
            strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}={17},{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}={41} WHERE {2}='{3}'",
            strSQL,
            TableName,
            IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
            IdxColumnName.OWNER_ID1.ToString, CI.OWNER_ID1,
            IdxColumnName.OWNER_ID2.ToString, CI.OWNER_ID2,
            IdxColumnName.OWNER_ID3.ToString, CI.OWNER_ID3,
            IdxColumnName.OWNER_ALIS1.ToString, CI.OWNER_ALIS1,
            IdxColumnName.OWNER_ALIS2.ToString, CI.OWNER_ALIS2,
            IdxColumnName.OWNER_DESC.ToString, CI.OWNER_DESC,
            IdxColumnName.OWNER_TYPE.ToString, CI.OWNER_TYPE,
            IdxColumnName.OWNER_COMMON1.ToString, CI.OWNER_COMMON1,
            IdxColumnName.OWNER_COMMON2.ToString, CI.OWNER_COMMON2,
            IdxColumnName.OWNER_COMMON3.ToString, CI.OWNER_COMMON3,
            IdxColumnName.OWNER_COMMON4.ToString, CI.OWNER_COMMON4,
            IdxColumnName.OWNER_COMMON5.ToString, CI.OWNER_COMMON5,
            IdxColumnName.OWNER_COMMON6.ToString, CI.OWNER_COMMON6,
            IdxColumnName.OWNER_COMMON7.ToString, CI.OWNER_COMMON7,
            IdxColumnName.OWNER_COMMON8.ToString, CI.OWNER_COMMON8,
            IdxColumnName.OWNER_COMMON9.ToString, CI.OWNER_COMMON9,
            IdxColumnName.OWNER_COMMON10.ToString, CI.OWNER_COMMON10,
            IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
            IdxColumnName.ENABLE.ToString, CI.ENABLE
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
    Public Shared Function AddWMS_M_OwnerData(ByVal Info As clsOwner, Optional ByVal SendToDB As Boolean = True) As Boolean
        SyncLock objLock
            Try
                If Info Is Nothing Then Return False
                If AddlstWMS_M_OwnerData(New List(Of clsOwner)({Info}), SendToDB) = True Then
                    Return True
                End If '-載不載入記憶體都是呼叫同一個function
                Return False
            Catch ex As Exception
                SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End Try
        End SyncLock
    End Function
    Public Shared Function AddlstWMS_M_OwnerData(ByVal Info As List(Of clsOwner), Optional ByVal SendToDB As Boolean = True) As Boolean
        SyncLock objLock
            Try
                If Info Is Nothing Then Return False
                If Info.Count = 0 Then Return True
                If DictionaryNeeded = 1 Then '-載入記憶體
                    For i = 0 To Info.Count - 1
                        Dim key As String = Info(i).gid
                        If dicData.ContainsKey(key) = True Then
                            SendMessageToLog("Add the same key: " & key, eCALogTool.ILogTool.enuTrcLevel.lvError)
                            Return False
                        End If
                    Next
                    If SendToDB Then
                        If InsertWMS_M_OwnerDataToDB(Info) Then
                            If AddOrUpdateWMS_M_OwnerDataToDictionary(Info) Then
                                SendMessageToLog("InsertDic WMS_M_OwnerData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                            Else
                                SendMessageToLog("InsertDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                                Return False
                            End If
                        Else
                            SendMessageToLog("InsertDB WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                            Return False
                        End If
                    Else
                        If AddOrUpdateWMS_M_OwnerDataToDictionary(Info) Then
                            SendMessageToLog("InsertDic WMS_M_OwnerData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                        Else
                            SendMessageToLog("InsertDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                            Return False
                        End If
                    End If
                Else
                    If SendToDB Then
                        If InsertWMS_M_OwnerDataToDB(Info) Then
                            Return True
                        Else
                            SendMessageToLog("InsertDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
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
    Public Shared Function UpdateWMS_M_OwnerData(ByVal Info As clsOwner, Optional ByVal SendToDB As Boolean = True) As Boolean
        SyncLock objLock
            Try
                If Info Is Nothing Then Return False
                If UpdatelstWMS_M_OwnerData(New List(Of clsOwner)({Info}), SendToDB) = True Then
                    Return True
                End If
                Return False
            Catch ex As Exception
                SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End Try
        End SyncLock
    End Function
    Public Shared Function UpdatelstWMS_M_OwnerData(ByVal Info As List(Of clsOwner), Optional ByVal SendToDB As Boolean = True) As Boolean
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
                        If UpdateWMS_M_OwnerDataToDB(Info) Then
                            If AddOrUpdateWMS_M_OwnerDataToDictionary(Info) Then
                                SendMessageToLog("UpdateDic WMS_M_OwnerData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                            Else
                                SendMessageToLog("UpdateDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                                Return False
                            End If
                        Else
                            SendMessageToLog("UpdateDB WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                            Return False
                        End If
                    Else
                        If AddOrUpdateWMS_M_OwnerDataToDictionary(Info) Then
                            SendMessageToLog("UpdateDic WMS_M_OwnerData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                        Else
                            SendMessageToLog("UpdateDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                            Return False
                        End If
                    End If
                Else
                    If SendToDB Then
                        If UpdateWMS_M_OwnerDataToDB(Info) Then
                            Return True
                        Else
                            SendMessageToLog("UpdateDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
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
    Public Shared Function DeleteWMS_M_OwnerData(ByVal Info As clsOwner, Optional ByVal SendToDB As Boolean = True) As Boolean
        SyncLock objLock
            Try
                If Info Is Nothing Then Return False
                If DeletelstWMS_M_OwnerData(New List(Of clsOwner)({Info}), SendToDB) = True Then
                    Return True
                End If '-載不載入記憶體都是呼叫同一個function
                Return False
            Catch ex As Exception
                SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End Try
        End SyncLock
    End Function
    Public Shared Function DeletelstWMS_M_OwnerData(ByVal Info As List(Of clsOwner), Optional ByVal SendToDB As Boolean = True) As Boolean
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
                        If DeleteWMS_M_OwnerDataToDB(Info) Then
                            If DeleteWMS_M_OwnerDataToDictionary(Info) Then
                                SendMessageToLog("DeleteDic WMS_M_OwnerData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                            Else
                                SendMessageToLog("DeleteDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                                Return False
                            End If
                        Else
                            SendMessageToLog("DeleteDB WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
                            Return False
                        End If
                    Else
                        If DeleteWMS_M_OwnerDataToDictionary(Info) Then
                            SendMessageToLog("DeleteDic WMS_M_OwnerData Success", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                        Else
                            SendMessageToLog("DeleteDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
                            Return False
                        End If
                    End If
                Else
                    If SendToDB Then
                        If DeleteWMS_M_OwnerDataToDB(Info) Then
                            Return True
                        Else
                            SendMessageToLog("DeleteDic WMS_M_OwnerData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
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
    Private Shared Function InsertWMS_M_OwnerDataToDB(ByRef Info As List(Of clsOwner)) As Integer
        Try
            If Info Is Nothing Then Return -1
            If Info.Count = 0 Then Return 0

            Dim strSQL As String = ""
            Dim rs As ADODB.Recordset = Nothing
            Dim lstSql As New List(Of String)
            For Each CI In Info
                strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}',{17},'{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}',{41})",
                strSQL,
                TableName,
                IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
                IdxColumnName.OWNER_ID1.ToString, CI.OWNER_ID1,
                IdxColumnName.OWNER_ID2.ToString, CI.OWNER_ID2,
                IdxColumnName.OWNER_ID3.ToString, CI.OWNER_ID3,
                IdxColumnName.OWNER_ALIS1.ToString, CI.OWNER_ALIS1,
                IdxColumnName.OWNER_ALIS2.ToString, CI.OWNER_ALIS2,
                IdxColumnName.OWNER_DESC.ToString, CI.OWNER_DESC,
                IdxColumnName.OWNER_TYPE.ToString, CI.OWNER_TYPE,
                IdxColumnName.OWNER_COMMON1.ToString, CI.OWNER_COMMON1,
                IdxColumnName.OWNER_COMMON2.ToString, CI.OWNER_COMMON2,
                IdxColumnName.OWNER_COMMON3.ToString, CI.OWNER_COMMON3,
                IdxColumnName.OWNER_COMMON4.ToString, CI.OWNER_COMMON4,
                IdxColumnName.OWNER_COMMON5.ToString, CI.OWNER_COMMON5,
                IdxColumnName.OWNER_COMMON6.ToString, CI.OWNER_COMMON6,
                IdxColumnName.OWNER_COMMON7.ToString, CI.OWNER_COMMON7,
                IdxColumnName.OWNER_COMMON8.ToString, CI.OWNER_COMMON8,
                IdxColumnName.OWNER_COMMON9.ToString, CI.OWNER_COMMON9,
                IdxColumnName.OWNER_COMMON10.ToString, CI.OWNER_COMMON10,
                IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
                IdxColumnName.ENABLE.ToString, CI.ENABLE
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
                SendMessageToLog("Insert to WMS_M_Owner DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function UpdateWMS_M_OwnerDataToDB(ByRef Info As List(Of clsOwner)) As Integer
        Try
            If Info Is Nothing Then Return -1
            If Info.Count = 0 Then Return 0

            Dim strSQL As String = ""
            Dim rs As ADODB.Recordset = Nothing
            Dim lstSql As New List(Of String)
            For Each CI In Info
                strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}={17},{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}={41} WHERE {2}='{3}'",
                strSQL,
                TableName,
                IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
                IdxColumnName.OWNER_ID1.ToString, CI.OWNER_ID1,
                IdxColumnName.OWNER_ID2.ToString, CI.OWNER_ID2,
                IdxColumnName.OWNER_ID3.ToString, CI.OWNER_ID3,
                IdxColumnName.OWNER_ALIS1.ToString, CI.OWNER_ALIS1,
                IdxColumnName.OWNER_ALIS2.ToString, CI.OWNER_ALIS2,
                IdxColumnName.OWNER_DESC.ToString, CI.OWNER_DESC,
                IdxColumnName.OWNER_TYPE.ToString, CI.OWNER_TYPE,
                IdxColumnName.OWNER_COMMON1.ToString, CI.OWNER_COMMON1,
                IdxColumnName.OWNER_COMMON2.ToString, CI.OWNER_COMMON2,
                IdxColumnName.OWNER_COMMON3.ToString, CI.OWNER_COMMON3,
                IdxColumnName.OWNER_COMMON4.ToString, CI.OWNER_COMMON4,
                IdxColumnName.OWNER_COMMON5.ToString, CI.OWNER_COMMON5,
                IdxColumnName.OWNER_COMMON6.ToString, CI.OWNER_COMMON6,
                IdxColumnName.OWNER_COMMON7.ToString, CI.OWNER_COMMON7,
                IdxColumnName.OWNER_COMMON8.ToString, CI.OWNER_COMMON8,
                IdxColumnName.OWNER_COMMON9.ToString, CI.OWNER_COMMON9,
                IdxColumnName.OWNER_COMMON10.ToString, CI.OWNER_COMMON10,
                IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
                IdxColumnName.ENABLE.ToString, CI.ENABLE
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
                SendMessageToLog("Update to WMS_M_Owner DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function DeleteWMS_M_OwnerDataToDB(ByRef Info As List(Of clsOwner)) As Integer
        Try
            If Info Is Nothing Then Return -1
            If Info.Count = 0 Then Return 0

            Dim strSQL As String = ""
            Dim rs As ADODB.Recordset = Nothing
            Dim lstSql As New List(Of String)
            For Each CI In Info
                strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
                strSQL,
                TableName,
                IdxColumnName.OWNER_NO.ToString, CI.OWNER_NO,
                IdxColumnName.OWNER_ID1.ToString, CI.OWNER_ID1,
                IdxColumnName.OWNER_ID2.ToString, CI.OWNER_ID2,
                IdxColumnName.OWNER_ID3.ToString, CI.OWNER_ID3,
                IdxColumnName.OWNER_ALIS1.ToString, CI.OWNER_ALIS1,
                IdxColumnName.OWNER_ALIS2.ToString, CI.OWNER_ALIS2,
                IdxColumnName.OWNER_DESC.ToString, CI.OWNER_DESC,
                IdxColumnName.OWNER_TYPE.ToString, CI.OWNER_TYPE,
                IdxColumnName.OWNER_COMMON1.ToString, CI.OWNER_COMMON1,
                IdxColumnName.OWNER_COMMON2.ToString, CI.OWNER_COMMON2,
                IdxColumnName.OWNER_COMMON3.ToString, CI.OWNER_COMMON3,
                IdxColumnName.OWNER_COMMON4.ToString, CI.OWNER_COMMON4,
                IdxColumnName.OWNER_COMMON5.ToString, CI.OWNER_COMMON5,
                IdxColumnName.OWNER_COMMON6.ToString, CI.OWNER_COMMON6,
                IdxColumnName.OWNER_COMMON7.ToString, CI.OWNER_COMMON7,
                IdxColumnName.OWNER_COMMON8.ToString, CI.OWNER_COMMON8,
                IdxColumnName.OWNER_COMMON9.ToString, CI.OWNER_COMMON9,
                IdxColumnName.OWNER_COMMON10.ToString, CI.OWNER_COMMON10,
                IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
                IdxColumnName.ENABLE.ToString, CI.ENABLE
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
                SendMessageToLog("Delete to WMS_M_Owner DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    '-內部記憶體增刪修
    Private Shared Function AddOrUpdateWMS_M_OwnerDataToDictionary(ByRef Info As List(Of clsOwner)) As Boolean
        Try
            If Info Is Nothing Then Return False
            If Info.Count = 0 Then Return True
            For Each CI In Info
                Dim _Data As clsOwner = CI
                Dim key As String = _Data.gid
                dicData.AddOrUpdate(key,
                _Data,
                Function(dicKey, ExistVal)
                    UpdateInfo(dicKey, ExistVal, _Data)
                    Return ExistVal
                End Function)
            Next
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function DeleteWMS_M_OwnerDataToDictionary(ByRef Info As List(Of clsOwner)) As Boolean
        Try
            If Info Is Nothing Then Return False
            If Info.Count = 0 Then Return True
            For i = 0 To Info.Count - 1
                Dim key As String = Info(i).gid
                If dicData.TryRemove(key, Nothing) = False Then
                    SendMessageToLog("dicData.TryRemove Failed -WMS_M_OwnerData", eCALogTool.ILogTool.enuTrcLevel.lvError)
                End If
            Next
            Return True
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
        End Try
    End Function
    Private Shared Function UpdateInfo(ByRef Key As String, ByRef Info As clsOwner, ByRef objNewTC As clsOwner) As clsOwner
        Try
            If Key = Info.gid Then
                Info.Update_To_Memory(objNewTC)
            Else
                SendMessageToLog("Dictionary has the different key", eCALogTool.ILogTool.enuTrcLevel.lvError)
            End If
        Catch ex As Exception
            SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
        End Try
        Return Info
    End Function
    Private Shared Function SetInfoFromDB(ByRef Info As clsOwner, ByRef RowData As DataRow) As Boolean
        Try
            If RowData IsNot Nothing Then
                Dim OWNER_NO = "" & RowData.Item(IdxColumnName.OWNER_NO.ToString)
                Dim OWNER_ID1 = "" & RowData.Item(IdxColumnName.OWNER_ID1.ToString)
                Dim OWNER_ID2 = "" & RowData.Item(IdxColumnName.OWNER_ID2.ToString)
                Dim OWNER_ID3 = "" & RowData.Item(IdxColumnName.OWNER_ID3.ToString)
                Dim OWNER_ALIS1 = "" & RowData.Item(IdxColumnName.OWNER_ALIS1.ToString)
                Dim OWNER_ALIS2 = "" & RowData.Item(IdxColumnName.OWNER_ALIS2.ToString)
                Dim OWNER_DESC = "" & RowData.Item(IdxColumnName.OWNER_DESC.ToString)
                Dim OWNER_TYPE = 0 & RowData.Item(IdxColumnName.OWNER_TYPE.ToString)
                Dim OWNER_COMMON1 = "" & RowData.Item(IdxColumnName.OWNER_COMMON1.ToString)
                Dim OWNER_COMMON2 = "" & RowData.Item(IdxColumnName.OWNER_COMMON2.ToString)
                Dim OWNER_COMMON3 = "" & RowData.Item(IdxColumnName.OWNER_COMMON3.ToString)
                Dim OWNER_COMMON4 = "" & RowData.Item(IdxColumnName.OWNER_COMMON4.ToString)
                Dim OWNER_COMMON5 = "" & RowData.Item(IdxColumnName.OWNER_COMMON5.ToString)
                Dim OWNER_COMMON6 = "" & RowData.Item(IdxColumnName.OWNER_COMMON6.ToString)
                Dim OWNER_COMMON7 = "" & RowData.Item(IdxColumnName.OWNER_COMMON7.ToString)
                Dim OWNER_COMMON8 = "" & RowData.Item(IdxColumnName.OWNER_COMMON8.ToString)
                Dim OWNER_COMMON9 = "" & RowData.Item(IdxColumnName.OWNER_COMMON9.ToString)
                Dim OWNER_COMMON10 = "" & RowData.Item(IdxColumnName.OWNER_COMMON10.ToString)
                Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
                Dim ENABLE = 0 & RowData.Item(IdxColumnName.ENABLE.ToString)
                Info = New clsOwner(OWNER_NO, OWNER_ID1, OWNER_ID2, OWNER_ID3, OWNER_ALIS1, OWNER_ALIS2, OWNER_DESC, OWNER_TYPE, OWNER_COMMON1, OWNER_COMMON2, OWNER_COMMON3, OWNER_COMMON4, OWNER_COMMON5, OWNER_COMMON6, OWNER_COMMON7, OWNER_COMMON8, OWNER_COMMON9, OWNER_COMMON10, COMMENTS, ENABLE)

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
    Public Shared Function GetdicOwnerByALL() As Dictionary(Of String, clsOwner)
        Try
            Dim dicReturn As New Dictionary(Of String, clsOwner)
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
                        Dim Info As clsOwner = Nothing
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
