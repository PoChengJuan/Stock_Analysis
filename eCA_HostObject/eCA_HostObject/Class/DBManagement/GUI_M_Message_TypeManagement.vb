﻿Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Partial Class GUI_M_Message_TypeManagement
  Public Shared TableName As String = "GUI_M_Message_Type"
  Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsGUI_M_Message_Type)
  Public Shared Property DictionaryNeeded As Integer = 0  '-需不需要載入記憶體
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    MESSAGE_TYPE
    MESSAGE_TYPE_ALIS
    ENABLE
  End Enum

  Public Enum UpdateOption As Integer
    UpdateDic = 0
    UpdateDB = 1
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef CI As clsGUI_M_Message_Type) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6}) values ({3},'{5}',{7})",
      strSQL,
      TableName,
      IdxColumnName.MESSAGE_TYPE.ToString, CI.MESSAGE_TYPE,
      IdxColumnName.MESSAGE_TYPE_ALIS.ToString, CI.MESSAGE_TYPE_ALIS,
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
  Public Shared Function GetDeleteSQL(ByRef CI As clsGUI_M_Message_Type) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}={3} ",
      strSQL,
      TableName,
      IdxColumnName.MESSAGE_TYPE.ToString, CI.MESSAGE_TYPE,
      IdxColumnName.MESSAGE_TYPE_ALIS.ToString, CI.MESSAGE_TYPE_ALIS,
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
  Public Shared Function GetUpdateSQL(ByRef CI As clsGUI_M_Message_Type) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}={7} WHERE {2}={3}",
      strSQL,
      TableName,
      IdxColumnName.MESSAGE_TYPE.ToString, CI.MESSAGE_TYPE,
      IdxColumnName.MESSAGE_TYPE_ALIS.ToString, CI.MESSAGE_TYPE_ALIS,
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

  Public Shared Function GetGUI_M_Message_TypeDataListByKey_MESSAGE_TYPE(ByVal message_type As Double) As Dictionary(Of String, clsGUI_M_Message_Type)
    SyncLock objLock
      Try
        Dim _lstReturn As New Dictionary(Of String, clsGUI_M_Message_Type)
        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty


            strSQL = String.Format("Select * from {1} WHERE  {2} = {3} ",
            strSQL,
            TableName,
            IdxColumnName.MESSAGE_TYPE.ToString, message_type
            )
            Dim DatasetMessage As New DataSet
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsGUI_M_Message_Type = Nothing
                SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
                If _lstReturn.ContainsKey(Info.gid) = False Then
                  _lstReturn.Add(Info.gid, Info)
                End If
              Next
            End If
          End If
        End If
        Return _lstReturn
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function
  Private Shared Function UpdateInfo(ByRef Key As String, ByRef Info As clsGUI_M_Message_Type, ByRef objNewTC As clsGUI_M_Message_Type) As clsGUI_M_Message_Type
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsGUI_M_Message_Type, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim MESSAGE_TYPE = 0 & RowData.Item(IdxColumnName.MESSAGE_TYPE.ToString)
        Dim MESSAGE_TYPE_DESC = "" & RowData.Item(IdxColumnName.MESSAGE_TYPE_ALIS.ToString)
        Dim ENABLE = 0 & RowData.Item(IdxColumnName.ENABLE.ToString)
        Info = New clsGUI_M_Message_Type(MESSAGE_TYPE, MESSAGE_TYPE_DESC, ENABLE)

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
End Class
