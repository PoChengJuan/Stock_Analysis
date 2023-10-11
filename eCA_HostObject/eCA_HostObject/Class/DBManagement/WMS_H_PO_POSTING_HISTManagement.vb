Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Public Class WMS_H_PO_POSTING_HISTManagement
    Public Shared TableName As String = "WMS_H_PO_POSTING_HIST"
    Public Shared dicData As New Concurrent.ConcurrentDictionary(Of String, clsPO_POSTING)
    Public Shared Property DictionaryNeeded As Integer = 1  '-需不需要載入記憶體
    Public Shared objLock As New Object
    Private Shared fUseBatchUpdate_DynamicConnection As Integer = 1
    Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
    Public Shared LogTool As eCALogTool._ILogTool = Nothing

    Enum IdxColumnName As Integer
        PO_ID
        PO_LINE_NO
        WO_ID
        SORT_ITEM_COMMON1
        SORT_ITEM_COMMON2
        SORT_ITEM_COMMON3
        SORT_ITEM_COMMON4
        SORT_ITEM_COMMON5
        QTY
        UUID
        CREATE_TIME
        UPDATE_TIME
        RESULT
        RESULT_MESSAGE
        H_POP1
        H_POP2
        H_POP3
        H_POP4
        H_POP5
        HIST_TIME
        SKU_NO
        CLOSE_USER_ID
        START_TRANSFER_TIME
        FINISH_TRANSFER_TIME
        ORDER_TYPE
        PO_SERIAL_NO
        TKNUM
        LOT_NO
        OWNER
        SUBOWNER
    End Enum

    '- GetSQL
    Public Shared Function GetInsertSQL(ByRef Info As clsPO_POSTING_HIST) As String
        Try

            Dim strSQL As String = ""
            strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60}) values ('{3}','{5}','{7}','{9}','{11}',{13},'{15}','{17}','{19}',{21},'{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}')",
            strSQL,
            TableName,
            IdxColumnName.PO_ID.ToString, Info.PO_ID,
            IdxColumnName.PO_LINE_NO.ToString, Info.PO_LINE_NO,
            IdxColumnName.WO_ID.ToString, Info.WO_ID,
            IdxColumnName.SORT_ITEM_COMMON1.ToString, Info.SORT_ITEM_COMMON1,
            IdxColumnName.SORT_ITEM_COMMON2.ToString, Info.SORT_ITEM_COMMON2,
            IdxColumnName.QTY.ToString, Info.QTY,
            IdxColumnName.UUID.ToString, Info.UUID,
            IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME,
            IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
            IdxColumnName.RESULT.ToString, Info.RESULT,
            IdxColumnName.RESULT_MESSAGE.ToString, Info.RESULT_MESSAGE,
            IdxColumnName.H_POP1.ToString, Info.H_POP1,
            IdxColumnName.H_POP2.ToString, Info.H_POP2,
            IdxColumnName.H_POP3.ToString, Info.H_POP3,
            IdxColumnName.H_POP4.ToString, Info.H_POP4,
            IdxColumnName.H_POP5.ToString, Info.H_POP5,
            IdxColumnName.HIST_TIME.ToString, Info.HIST_TIME,
            IdxColumnName.SORT_ITEM_COMMON3.ToString, Info.SORT_ITEM_COMMON3,
            IdxColumnName.SORT_ITEM_COMMON4.ToString, Info.SORT_ITEM_COMMON4,
            IdxColumnName.SORT_ITEM_COMMON5.ToString, Info.SORT_ITEM_COMMON5,
            IdxColumnName.SKU_NO.ToString, Info.SKU_NO,
            IdxColumnName.CLOSE_USER_ID.ToString, Info.CLOSE_USER_ID,
            IdxColumnName.START_TRANSFER_TIME.ToString, Info.START_TRANSFER_TIME,
            IdxColumnName.FINISH_TRANSFER_TIME.ToString, Info.FINISH_TRANSFER_TIME,
            IdxColumnName.ORDER_TYPE.ToString, CInt(Info.ORDER_TYPE),
            IdxColumnName.PO_SERIAL_NO.ToString, Info.PO_SERIAL_NO,
            IdxColumnName.TKNUM.ToString, Info.TKNUM,
            IdxColumnName.LOT_NO.ToString, Info.LOT_NO,
            IdxColumnName.OWNER.ToString, Info.OWNER,
            IdxColumnName.SUBOWNER.ToString, Info.SUBOWNER
      )
            Return strSQL
        Catch ex As Exception
            SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return Nothing
        End Try
    End Function

    Public Shared Function CheckPO_POSTING_HISTByPO_ID(ByVal PO_ID As String) As Boolean
        SyncLock objLock
            Try
                If DBTool IsNot Nothing Then
                    If DBTool.isConnection(DBTool.m_CN) = True Then
                        Dim strSQL As String = String.Empty
                        Dim DatasetMessage As New DataSet

                        strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' AND {4}={5} ",
                        strSQL,
                        TableName,
                        IdxColumnName.PO_ID.ToString, PO_ID,
                        IdxColumnName.RESULT.ToString, 0
                        )
                        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
                        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)


                        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
                            Return True
                        Else
                            Return False
                        End If
                    End If
                End If
                Return True
            Catch ex As Exception
                SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
                Return False
            End Try
        End SyncLock
    End Function
End Class
