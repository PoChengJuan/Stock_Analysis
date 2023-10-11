Partial Class WMS_M_Item_LabelManagement
  Public Shared TableName As String = "WMS_M_Item_Label"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    ITEM_LABEL_ID
    ITEM_LABEL_TYPE
    PO_ID
    TAG1
    TAG2
    TAG3
    TAG4
    TAG5
    TAG6
    TAG7
    TAG8
    TAG9
    TAG10
    TAG11
    TAG12
    TAG13
    TAG14
    TAG15
    TAG16
    TAG17
    TAG18
    TAG19
    TAG20
    TAG21
    TAG22
    TAG23
    TAG24
    TAG25
    TAG26
    TAG27
    TAG28
    TAG29
    TAG30
    TAG31
    TAG32
    TAG33
    TAG34
    TAG35
    PRINTED
    CREATE_USER
    FIRST_PRINT_TIME
    LAST_PRINT_TIME
    UPDATE_TIME
    CREATE_TIME
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsItemLabel) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26},{28},{30},{32},{34},{36},{38},{40},{42},{44},{46},{48},{50},{52},{54},{56},{58},{60},{62},{64},{66},{68},{70},{72},{74},{76},{78},{80},{82},{84},{86},{88}) values ('{3}',{5},'{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}','{29}','{31}','{33}','{35}','{37}','{39}','{41}','{43}','{45}','{47}','{49}','{51}','{53}','{55}','{57}','{59}','{61}','{63}','{65}','{67}','{69}','{71}','{73}','{75}','{77}'),'{79}','{81}','{83}','{85}','{87}','{89}'",
      strSQL,
      TableName,
      IdxColumnName.ITEM_LABEL_ID.ToString, Info.ITEM_LABEL_ID,
      IdxColumnName.ITEM_LABEL_TYPE.ToString, Info.ITEM_LABEL_TYPE,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.TAG1.ToString, Info.TAG1,
      IdxColumnName.TAG2.ToString, Info.TAG2,
      IdxColumnName.TAG3.ToString, Info.TAG3,
      IdxColumnName.TAG4.ToString, Info.TAG4,
      IdxColumnName.TAG5.ToString, Info.TAG5,
      IdxColumnName.TAG6.ToString, Info.TAG6,
      IdxColumnName.TAG7.ToString, Info.TAG7,
      IdxColumnName.TAG8.ToString, Info.TAG8,
      IdxColumnName.TAG9.ToString, Info.TAG9,
      IdxColumnName.TAG10.ToString, Info.TAG10,
      IdxColumnName.TAG11.ToString, Info.TAG11,
      IdxColumnName.TAG12.ToString, Info.TAG12,
      IdxColumnName.TAG13.ToString, Info.TAG13,
      IdxColumnName.TAG14.ToString, Info.TAG14,
      IdxColumnName.TAG15.ToString, Info.TAG15,
      IdxColumnName.TAG16.ToString, Info.TAG16,
      IdxColumnName.TAG17.ToString, Info.TAG17,
      IdxColumnName.TAG18.ToString, Info.TAG18,
      IdxColumnName.TAG19.ToString, Info.TAG19,
      IdxColumnName.TAG20.ToString, Info.TAG20,
      IdxColumnName.TAG21.ToString, Info.TAG21,
      IdxColumnName.TAG22.ToString, Info.TAG22,
      IdxColumnName.TAG23.ToString, Info.TAG23,
      IdxColumnName.TAG24.ToString, Info.TAG24,
      IdxColumnName.TAG25.ToString, Info.TAG25,
      IdxColumnName.TAG26.ToString, Info.TAG26,
      IdxColumnName.TAG27.ToString, Info.TAG27,
      IdxColumnName.TAG28.ToString, Info.TAG28,
      IdxColumnName.TAG29.ToString, Info.TAG29,
      IdxColumnName.TAG30.ToString, Info.TAG30,
      IdxColumnName.TAG31.ToString, Info.TAG31,
      IdxColumnName.TAG32.ToString, Info.TAG32,
      IdxColumnName.TAG33.ToString, Info.TAG33,
      IdxColumnName.TAG34.ToString, Info.TAG34,
      IdxColumnName.TAG35.ToString, Info.TAG35,
      IdxColumnName.PRINTED.ToString, Info.PRINTED,
      IdxColumnName.CREATE_USER.ToString, Info.CREATE_USER,
      IdxColumnName.FIRST_PRINT_TIME.ToString, Info.FIRST_PRINT_TIME,
      IdxColumnName.LAST_PRINT_TIME.ToString, Info.LAST_PRINT_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsItemLabel) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}={5},{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}',{16}='{17}',{18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}',{28}='{29}',{30}='{31}',{32}='{33}',{34}='{35}',{36}='{37}',{38}='{39}',{40}='{41}',{42}='{43}',{44}='{45}',{46}='{47}',{48}='{49}',{50}='{51}',{52}='{53}',{54}='{55}',{56}='{57}',{58}='{59}',{60}='{61}',{62}='{63}',{64}='{65}',{66}='{67}',{68}='{69}',{70}='{71}',{72}='{73}',{74}='{75}',{76}='{77}',{78}='{79}',{80}='{81}',{82}='{83}',{84}='{85}',{86}='{87}',{88}='{89}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.ITEM_LABEL_ID.ToString, Info.ITEM_LABEL_ID,
      IdxColumnName.ITEM_LABEL_TYPE.ToString, Info.ITEM_LABEL_TYPE,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.TAG1.ToString, Info.TAG1,
      IdxColumnName.TAG2.ToString, Info.TAG2,
      IdxColumnName.TAG3.ToString, Info.TAG3,
      IdxColumnName.TAG4.ToString, Info.TAG4,
      IdxColumnName.TAG5.ToString, Info.TAG5,
      IdxColumnName.TAG6.ToString, Info.TAG6,
      IdxColumnName.TAG7.ToString, Info.TAG7,
      IdxColumnName.TAG8.ToString, Info.TAG8,
      IdxColumnName.TAG9.ToString, Info.TAG9,
      IdxColumnName.TAG10.ToString, Info.TAG10,
      IdxColumnName.TAG11.ToString, Info.TAG11,
      IdxColumnName.TAG12.ToString, Info.TAG12,
      IdxColumnName.TAG13.ToString, Info.TAG13,
      IdxColumnName.TAG14.ToString, Info.TAG14,
      IdxColumnName.TAG15.ToString, Info.TAG15,
      IdxColumnName.TAG16.ToString, Info.TAG16,
      IdxColumnName.TAG17.ToString, Info.TAG17,
      IdxColumnName.TAG18.ToString, Info.TAG18,
      IdxColumnName.TAG19.ToString, Info.TAG19,
      IdxColumnName.TAG20.ToString, Info.TAG20,
      IdxColumnName.TAG21.ToString, Info.TAG21,
      IdxColumnName.TAG22.ToString, Info.TAG22,
      IdxColumnName.TAG23.ToString, Info.TAG23,
      IdxColumnName.TAG24.ToString, Info.TAG24,
      IdxColumnName.TAG25.ToString, Info.TAG25,
      IdxColumnName.TAG26.ToString, Info.TAG26,
      IdxColumnName.TAG27.ToString, Info.TAG27,
      IdxColumnName.TAG28.ToString, Info.TAG28,
      IdxColumnName.TAG29.ToString, Info.TAG29,
      IdxColumnName.TAG30.ToString, Info.TAG30,
      IdxColumnName.TAG31.ToString, Info.TAG31,
      IdxColumnName.TAG32.ToString, Info.TAG32,
      IdxColumnName.TAG33.ToString, Info.TAG33,
      IdxColumnName.TAG34.ToString, Info.TAG34,
      IdxColumnName.TAG35.ToString, Info.TAG35,
      IdxColumnName.PRINTED.ToString, Info.PRINTED,
      IdxColumnName.CREATE_USER.ToString, Info.CREATE_USER,
      IdxColumnName.FIRST_PRINT_TIME.ToString, Info.FIRST_PRINT_TIME,
      IdxColumnName.LAST_PRINT_TIME.ToString, Info.LAST_PRINT_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsItemLabel) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.ITEM_LABEL_ID.ToString, Info.ITEM_LABEL_ID,
      IdxColumnName.ITEM_LABEL_TYPE.ToString, Info.ITEM_LABEL_TYPE,
      IdxColumnName.PO_ID.ToString, Info.PO_ID,
      IdxColumnName.TAG1.ToString, Info.TAG1,
      IdxColumnName.TAG2.ToString, Info.TAG2,
      IdxColumnName.TAG3.ToString, Info.TAG3,
      IdxColumnName.TAG4.ToString, Info.TAG4,
      IdxColumnName.TAG5.ToString, Info.TAG5,
      IdxColumnName.TAG6.ToString, Info.TAG6,
      IdxColumnName.TAG7.ToString, Info.TAG7,
      IdxColumnName.TAG8.ToString, Info.TAG8,
      IdxColumnName.TAG9.ToString, Info.TAG9,
      IdxColumnName.TAG10.ToString, Info.TAG10,
      IdxColumnName.TAG11.ToString, Info.TAG11,
      IdxColumnName.TAG12.ToString, Info.TAG12,
      IdxColumnName.TAG13.ToString, Info.TAG13,
      IdxColumnName.TAG14.ToString, Info.TAG14,
      IdxColumnName.TAG15.ToString, Info.TAG15,
      IdxColumnName.TAG16.ToString, Info.TAG16,
      IdxColumnName.TAG17.ToString, Info.TAG17,
      IdxColumnName.TAG18.ToString, Info.TAG18,
      IdxColumnName.TAG19.ToString, Info.TAG19,
      IdxColumnName.TAG20.ToString, Info.TAG20,
      IdxColumnName.TAG21.ToString, Info.TAG21,
      IdxColumnName.TAG22.ToString, Info.TAG22,
      IdxColumnName.TAG23.ToString, Info.TAG23,
      IdxColumnName.TAG24.ToString, Info.TAG24,
      IdxColumnName.TAG25.ToString, Info.TAG25,
      IdxColumnName.TAG26.ToString, Info.TAG26,
      IdxColumnName.TAG27.ToString, Info.TAG27,
      IdxColumnName.TAG28.ToString, Info.TAG28,
      IdxColumnName.TAG29.ToString, Info.TAG29,
      IdxColumnName.TAG30.ToString, Info.TAG30,
      IdxColumnName.TAG31.ToString, Info.TAG31,
      IdxColumnName.TAG32.ToString, Info.TAG32,
      IdxColumnName.TAG33.ToString, Info.TAG33,
      IdxColumnName.TAG34.ToString, Info.TAG34,
      IdxColumnName.TAG35.ToString, Info.TAG35,
      IdxColumnName.PRINTED.ToString, Info.PRINTED,
      IdxColumnName.CREATE_USER.ToString, Info.CREATE_USER,
      IdxColumnName.FIRST_PRINT_TIME.ToString, Info.FIRST_PRINT_TIME,
      IdxColumnName.LAST_PRINT_TIME.ToString, Info.LAST_PRINT_TIME,
      IdxColumnName.UPDATE_TIME.ToString, Info.UPDATE_TIME,
      IdxColumnName.CREATE_TIME.ToString, Info.CREATE_TIME
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsItemLabel, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim ITEM_LABEL_ID = "" & RowData.Item(IdxColumnName.ITEM_LABEL_ID.ToString)
        Dim ITEM_LABEL_TYPE = If(IsNumeric(RowData.Item(IdxColumnName.ITEM_LABEL_TYPE.ToString)), RowData.Item(IdxColumnName.ITEM_LABEL_TYPE.ToString), 0 & RowData.Item(IdxColumnName.ITEM_LABEL_TYPE.ToString))
        Dim PO_ID = "" & RowData.Item(IdxColumnName.PO_ID.ToString)
        Dim TAG1 = "" & RowData.Item(IdxColumnName.TAG1.ToString)
        Dim TAG2 = "" & RowData.Item(IdxColumnName.TAG2.ToString)
        Dim TAG3 = "" & RowData.Item(IdxColumnName.TAG3.ToString)
        Dim TAG4 = "" & RowData.Item(IdxColumnName.TAG4.ToString)
        Dim TAG5 = "" & RowData.Item(IdxColumnName.TAG5.ToString)
        Dim TAG6 = "" & RowData.Item(IdxColumnName.TAG6.ToString)
        Dim TAG7 = "" & RowData.Item(IdxColumnName.TAG7.ToString)
        Dim TAG8 = "" & RowData.Item(IdxColumnName.TAG8.ToString)
        Dim TAG9 = "" & RowData.Item(IdxColumnName.TAG9.ToString)
        Dim TAG10 = "" & RowData.Item(IdxColumnName.TAG10.ToString)
        Dim TAG11 = "" & RowData.Item(IdxColumnName.TAG11.ToString)
        Dim TAG12 = "" & RowData.Item(IdxColumnName.TAG12.ToString)
        Dim TAG13 = "" & RowData.Item(IdxColumnName.TAG13.ToString)
        Dim TAG14 = "" & RowData.Item(IdxColumnName.TAG14.ToString)
        Dim TAG15 = "" & RowData.Item(IdxColumnName.TAG15.ToString)
        Dim TAG16 = "" & RowData.Item(IdxColumnName.TAG16.ToString)
        Dim TAG17 = "" & RowData.Item(IdxColumnName.TAG17.ToString)
        Dim TAG18 = "" & RowData.Item(IdxColumnName.TAG18.ToString)
        Dim TAG19 = "" & RowData.Item(IdxColumnName.TAG19.ToString)
        Dim TAG20 = "" & RowData.Item(IdxColumnName.TAG20.ToString)
        Dim TAG21 = "" & RowData.Item(IdxColumnName.TAG21.ToString)
        Dim TAG22 = "" & RowData.Item(IdxColumnName.TAG22.ToString)
        Dim TAG23 = "" & RowData.Item(IdxColumnName.TAG23.ToString)
        Dim TAG24 = "" & RowData.Item(IdxColumnName.TAG24.ToString)
        Dim TAG25 = "" & RowData.Item(IdxColumnName.TAG25.ToString)
        Dim TAG26 = "" & RowData.Item(IdxColumnName.TAG26.ToString)
        Dim TAG27 = "" & RowData.Item(IdxColumnName.TAG27.ToString)
        Dim TAG28 = "" & RowData.Item(IdxColumnName.TAG28.ToString)
        Dim TAG29 = "" & RowData.Item(IdxColumnName.TAG29.ToString)
        Dim TAG30 = "" & RowData.Item(IdxColumnName.TAG30.ToString)
        Dim TAG31 = "" & RowData.Item(IdxColumnName.TAG31.ToString)
        Dim TAG32 = "" & RowData.Item(IdxColumnName.TAG32.ToString)
        Dim TAG33 = "" & RowData.Item(IdxColumnName.TAG33.ToString)
        Dim TAG34 = "" & RowData.Item(IdxColumnName.TAG34.ToString)
        Dim TAG35 = "" & RowData.Item(IdxColumnName.TAG35.ToString)
        Dim PRINTED = If(IsNumeric(RowData.Item(IdxColumnName.PRINTED.ToString)), RowData.Item(IdxColumnName.PRINTED.ToString), 0 & RowData.Item(IdxColumnName.PRINTED.ToString))
        Dim CREATE_USER = "" & RowData.Item(IdxColumnName.CREATE_USER.ToString)

        Dim FIRST_PRINT_TIME = "" & RowData.Item(IdxColumnName.FIRST_PRINT_TIME.ToString)
        Dim LAST_PRINT_TIME = "" & RowData.Item(IdxColumnName.LAST_PRINT_TIME.ToString)
        Dim UPDATE_TIME = "" & RowData.Item(IdxColumnName.UPDATE_TIME.ToString)
        Dim CREATE_TIME = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
        Info = New clsItemLabel(ITEM_LABEL_ID, ITEM_LABEL_TYPE, PO_ID, TAG1, TAG2, TAG3, TAG4, TAG5, TAG6, TAG7, TAG8, TAG9, TAG10, TAG11, TAG12, TAG13, TAG14, TAG15, TAG16, TAG17, TAG18, TAG19, TAG20, TAG21, TAG22, TAG23, TAG24, TAG25, TAG26, TAG27, TAG28, TAG29, TAG30, TAG31, TAG32, TAG33, TAG34, TAG35, PRINTED, CREATE_USER, FIRST_PRINT_TIME, LAST_PRINT_TIME, UPDATE_TIME, CREATE_TIME)

      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Shared Function GetWMS_M_Item_LabelListByALL() As List(Of clsItemLabel)
    Try
      Dim _lstReturn As New List(Of clsItemLabel)
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
            Dim Info As clsItemLabel = Nothing
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
  Public Shared Function GetdicItemLabelByGuid(ByVal Guid As String) As Dictionary(Of String, clsItemLabel)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsItemLabel)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strUniqueIDList As String = ""

        If strUniqueIDList = "" Then
          strUniqueIDList = "'" & Guid & "'"
        Else
          strUniqueIDList = strUniqueIDList & ",'" & Guid & "'"
        End If
        If strWhere = "" Then
          strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.ITEM_LABEL_ID.ToString, strUniqueIDList)
        Else
          strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.ITEM_LABEL_ID.ToString, strUniqueIDList)
        End If
        strSQL = String.Format("Select * from {1} {2} ",
            strSQL,
            TableName,
            strWhere
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsItemLabel = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetdicItemLabelByPO_ID(ByVal PO_ID As String) As Dictionary(Of String, clsItemLabel)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsItemLabel)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strPO_IDList As String = ""

        If strPO_IDList = "" Then
          strPO_IDList = "'" & PO_ID & "'"
        Else
          strPO_IDList = strPO_IDList & ",'" & PO_ID & "'"
        End If
        If strWhere = "" Then
          strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.PO_ID.ToString, strPO_IDList)
        Else
          strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.PO_ID.ToString, strPO_IDList)
        End If
        strSQL = String.Format("Select * from {1} {2} ",
            strSQL,
            TableName,
            strWhere
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsItemLabel = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
          Next
        End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetdicItemLabelByPackage_ID(ByVal Package_ID As String) As Dictionary(Of String, clsItemLabel)
    Try
      Dim _lstReturn As New Dictionary(Of String, clsItemLabel)
      If DBTool IsNot Nothing Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        Dim strWhere As String = ""
        Dim strPO_IDList As String = ""

        If strPO_IDList = "" Then
          strPO_IDList = "'" & Package_ID & "'"
        Else
          strPO_IDList = strPO_IDList & ",'" & Package_ID & "'"
        End If
        If strWhere = "" Then
          strWhere = String.Format("WHERE {0} IN ({1}) ", IdxColumnName.ITEM_LABEL_ID.ToString, strPO_IDList)
        Else
          strWhere = String.Format("{0} AND {1} = ({2}) ", strWhere, IdxColumnName.ITEM_LABEL_ID.ToString, strPO_IDList)
        End If
        strSQL = String.Format("Select * from {1} {2} ",
            strSQL,
            TableName,
            strWhere
        )
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute_DynamicConnection(strSQL, DatasetMessage)
        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsItemLabel = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            If _lstReturn.ContainsKey(Info.gid) = False Then
              _lstReturn.Add(Info.gid, Info)
            End If
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
