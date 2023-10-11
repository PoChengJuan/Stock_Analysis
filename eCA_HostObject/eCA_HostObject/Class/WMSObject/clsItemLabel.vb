Public Class clsItemLabel
  Private ShareName As String = "ItemLabel"
  Private ShareKey As String = ""
  Private _gid As String
  Private _ITEM_LABEL_ID As String '物料條碼

  Private _ITEM_LABEL_TYPE As Long '物料標籤類型

  Private _PO_ID As String '標籤對應的PO_ID

  Private _TAG1 As String '快速輸入碼，會印在標籤上，二維碼損壞時，可用此碼快速輸入

  Private _TAG2 As String '內部訂單號前6碼

  Private _TAG3 As String 'BOM表PK

  Private _TAG4 As String '面料長度

  Private _TAG5 As String '淨重

  Private _TAG6 As String '毛重

  Private _TAG7 As String '包裝尺寸：深

  Private _TAG8 As String '包裝尺寸：寬

  Private _TAG9 As String '包裝尺寸：高

  Private _TAG10 As String '出口批次號

  Private _TAG11 As String '捲號

  Private _TAG12 As String '缸號

  Private _TAG13 As String '供應商代碼

  Private _TAG14 As String '正常狀態為 N，Y:表示被供應商標示刪除，參考用，避免供應商印出標籤貼上後，誤刪資料。

  Private _TAG15 As String '料號

  Private _TAG16 As String '面料品項規格

  Private _TAG17 As String '

  Private _TAG18 As String '

  Private _TAG19 As String '

  Private _TAG20 As String '

  Private _TAG21 As String '

  Private _TAG22 As String '

  Private _TAG23 As String '

  Private _TAG24 As String '

  Private _TAG25 As String '

  Private _TAG26 As String '

  Private _TAG27 As String '

  Private _TAG28 As String '

  Private _TAG29 As String '

  Private _TAG30 As String '

  Private _TAG31 As String  '

  Private _TAG32 As String

  Private _TAG33 As String

  Private _TAG34 As String

  Private _TAG35 As String


  Private _PRINTED As Long '

  Private _CREATE_USER As String '

  Private _FISRT_PRINT_TIME As String

  Private _LAST_PRINT_TIME As String

  Private _UPDATE_TIME As String

  Private _CREATE_TIME As String '建立時間

  Private _objWMS As clsHandlingObject
  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property ITEM_LABEL_ID() As String
    Get
      Return _ITEM_LABEL_ID
    End Get
    Set(ByVal value As String)
      _ITEM_LABEL_ID = value
    End Set
  End Property
  Public Property ITEM_LABEL_TYPE() As Long
    Get
      Return _ITEM_LABEL_TYPE
    End Get
    Set(ByVal value As Long)
      _ITEM_LABEL_TYPE = value
    End Set
  End Property
  Public Property PO_ID() As String
    Get
      Return _PO_ID
    End Get
    Set(ByVal value As String)
      _PO_ID = value
    End Set
  End Property
  Public Property TAG1() As String
    Get
      Return _TAG1
    End Get
    Set(ByVal value As String)
      _TAG1 = value
    End Set
  End Property
  Public Property TAG2() As String
    Get
      Return _TAG2
    End Get
    Set(ByVal value As String)
      _TAG2 = value
    End Set
  End Property
  Public Property TAG3() As String
    Get
      Return _TAG3
    End Get
    Set(ByVal value As String)
      _TAG3 = value
    End Set
  End Property
  Public Property TAG4() As String
    Get
      Return _TAG4
    End Get
    Set(ByVal value As String)
      _TAG4 = value
    End Set
  End Property
  Public Property TAG5() As String
    Get
      Return _TAG5
    End Get
    Set(ByVal value As String)
      _TAG5 = value
    End Set
  End Property
  Public Property TAG6() As String
    Get
      Return _TAG6
    End Get
    Set(ByVal value As String)
      _TAG6 = value
    End Set
  End Property
  Public Property TAG7() As String
    Get
      Return _TAG7
    End Get
    Set(ByVal value As String)
      _TAG7 = value
    End Set
  End Property
  Public Property TAG8() As String
    Get
      Return _TAG8
    End Get
    Set(ByVal value As String)
      _TAG8 = value
    End Set
  End Property
  Public Property TAG9() As String
    Get
      Return _TAG9
    End Get
    Set(ByVal value As String)
      _TAG9 = value
    End Set
  End Property
  Public Property TAG10() As String
    Get
      Return _TAG10
    End Get
    Set(ByVal value As String)
      _TAG10 = value
    End Set
  End Property
  Public Property TAG11() As String
    Get
      Return _TAG11
    End Get
    Set(ByVal value As String)
      _TAG11 = value
    End Set
  End Property
  Public Property TAG12() As String
    Get
      Return _TAG12
    End Get
    Set(ByVal value As String)
      _TAG12 = value
    End Set
  End Property
  Public Property TAG13() As String
    Get
      Return _TAG13
    End Get
    Set(ByVal value As String)
      _TAG13 = value
    End Set
  End Property
  Public Property TAG14() As String
    Get
      Return _TAG14
    End Get
    Set(ByVal value As String)
      _TAG14 = value
    End Set
  End Property
  Public Property TAG15() As String
    Get
      Return _TAG15
    End Get
    Set(ByVal value As String)
      _TAG15 = value
    End Set
  End Property
  Public Property TAG16() As String
    Get
      Return _TAG16
    End Get
    Set(ByVal value As String)
      _TAG16 = value
    End Set
  End Property
  Public Property TAG17() As String
    Get
      Return _TAG17
    End Get
    Set(ByVal value As String)
      _TAG17 = value
    End Set
  End Property
  Public Property TAG18() As String
    Get
      Return _TAG18
    End Get
    Set(ByVal value As String)
      _TAG18 = value
    End Set
  End Property
  Public Property TAG19() As String
    Get
      Return _TAG19
    End Get
    Set(ByVal value As String)
      _TAG19 = value
    End Set
  End Property
  Public Property TAG20() As String
    Get
      Return _TAG20
    End Get
    Set(ByVal value As String)
      _TAG20 = value
    End Set
  End Property
  Public Property TAG21() As String
    Get
      Return _TAG21
    End Get
    Set(ByVal value As String)
      _TAG21 = value
    End Set
  End Property
  Public Property TAG22() As String
    Get
      Return _TAG22
    End Get
    Set(ByVal value As String)
      _TAG22 = value
    End Set
  End Property
  Public Property TAG23() As String
    Get
      Return _TAG23
    End Get
    Set(ByVal value As String)
      _TAG23 = value
    End Set
  End Property
  Public Property TAG24() As String
    Get
      Return _TAG24
    End Get
    Set(ByVal value As String)
      _TAG24 = value
    End Set
  End Property
  Public Property TAG25() As String
    Get
      Return _TAG25
    End Get
    Set(ByVal value As String)
      _TAG25 = value
    End Set
  End Property
  Public Property TAG26() As String
    Get
      Return _TAG26
    End Get
    Set(ByVal value As String)
      _TAG26 = value
    End Set
  End Property
  Public Property TAG27() As String
    Get
      Return _TAG27
    End Get
    Set(ByVal value As String)
      _TAG27 = value
    End Set
  End Property
  Public Property TAG28() As String
    Get
      Return _TAG28
    End Get
    Set(ByVal value As String)
      _TAG28 = value
    End Set
  End Property
  Public Property TAG29() As String
    Get
      Return _TAG29
    End Get
    Set(ByVal value As String)
      _TAG29 = value
    End Set
  End Property
  Public Property TAG30() As String
    Get
      Return _TAG30
    End Get
    Set(ByVal value As String)
      _TAG30 = value
    End Set
  End Property
  Public Property TAG31() As String
    Get
      Return _TAG31
    End Get
    Set(ByVal value As String)
      _TAG31 = value
    End Set
  End Property
  Public Property TAG32() As String
    Get
      Return _TAG32
    End Get
    Set(ByVal value As String)
      _TAG32 = value
    End Set
  End Property
  Public Property TAG33() As String
    Get
      Return _TAG33
    End Get
    Set(ByVal value As String)
      _TAG33 = value
    End Set
  End Property
  Public Property TAG34() As String
    Get
      Return _TAG34
    End Get
    Set(ByVal value As String)
      _TAG34 = value
    End Set
  End Property
  Public Property TAG35() As String
    Get
      Return _TAG35
    End Get
    Set(ByVal value As String)
      _TAG35 = value
    End Set
  End Property
  Public Property PRINTED() As Long
    Get
      Return _PRINTED
    End Get
    Set(ByVal value As Long)
      _PRINTED = value
    End Set
  End Property
  Public Property CREATE_USER() As String
    Get
      Return _CREATE_USER
    End Get
    Set(ByVal value As String)
      _CREATE_USER = value
    End Set
  End Property
  Public Property FIRST_PRINT_TIME() As String
    Get
      Return _FISRT_PRINT_TIME
    End Get
    Set(ByVal value As String)
      _FISRT_PRINT_TIME = value
    End Set
  End Property
  Public Property LAST_PRINT_TIME() As String
    Get
      Return _LAST_PRINT_TIME
    End Get
    Set(ByVal value As String)
      _LAST_PRINT_TIME = value
    End Set
  End Property
  Public Property UPDATE_TIME() As String
    Get
      Return _UPDATE_TIME
    End Get
    Set(ByVal value As String)
      _UPDATE_TIME = value
    End Set
  End Property
  Public Property CREATE_TIME() As String
    Get
      Return _CREATE_TIME
    End Get
    Set(ByVal value As String)
      _CREATE_TIME = value
    End Set
  End Property
  Public Property objWMS() As clsHandlingObject
    Get
      Return _objWMS
    End Get
    Set(ByVal value As clsHandlingObject)
      _objWMS = value
    End Set
  End Property

  Public Sub New(ByVal ITEM_LABEL_ID As String, ByVal ITEM_LABEL_TYPE As Long, ByVal PO_ID As String, ByVal TAG1 As String, ByVal TAG2 As String, ByVal TAG3 As String, ByVal TAG4 As String, ByVal TAG5 As String, ByVal TAG6 As String, ByVal TAG7 As String, ByVal TAG8 As String, ByVal TAG9 As String, ByVal TAG10 As String, ByVal TAG11 As String, ByVal TAG12 As String, ByVal TAG13 As String, ByVal TAG14 As String, ByVal TAG15 As String, ByVal TAG16 As String, ByVal TAG17 As String, ByVal TAG18 As String, ByVal TAG19 As String, ByVal TAG20 As String, ByVal TAG21 As String, ByVal TAG22 As String, ByVal TAG23 As String, ByVal TAG24 As String, ByVal TAG25 As String, ByVal TAG26 As String, ByVal TAG27 As String, ByVal TAG28 As String, ByVal TAG29 As String, ByVal TAG30 As String, ByVal TAG31 As String, ByVal TAG32 As String, ByVal TAG33 As String, ByVal TAG34 As String, ByVal TAG35 As String, ByVal PRINTED As Long, ByVal CREATE_USER As String, ByVal FIRST_PRINT_TIME As String, ByVal LAST_PRINT_TIME As String, ByVal UPDATE_TIME As String, ByVal CREATE_TIME As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(ITEM_LABEL_ID)
      _gid = key
      _ITEM_LABEL_ID = ITEM_LABEL_ID
      _ITEM_LABEL_TYPE = ITEM_LABEL_TYPE
      _PO_ID = PO_ID
      _TAG1 = TAG1
      _TAG2 = TAG2
      _TAG3 = TAG3
      _TAG4 = TAG4
      _TAG5 = TAG5
      _TAG6 = TAG6
      _TAG7 = TAG7
      _TAG8 = TAG8
      _TAG9 = TAG9
      _TAG10 = TAG10
      _TAG11 = TAG11
      _TAG12 = TAG12
      _TAG13 = TAG13
      _TAG14 = TAG14
      _TAG15 = TAG15
      _TAG16 = TAG16
      _TAG17 = TAG17
      _TAG18 = TAG18
      _TAG19 = TAG19
      _TAG20 = TAG20
      _TAG21 = TAG21
      _TAG22 = TAG22
      _TAG23 = TAG23
      _TAG24 = TAG24
      _TAG25 = TAG25
      _TAG26 = TAG26
      _TAG27 = TAG27
      _TAG28 = TAG28
      _TAG29 = TAG29
      _TAG30 = TAG30
      _TAG31 = TAG31
      _TAG32 = TAG32
      _TAG33 = TAG33
      _TAG34 = TAG34
      _TAG35 = TAG35
      _PRINTED = PRINTED
      _CREATE_USER = CREATE_USER
      _FISRT_PRINT_TIME = FIRST_PRINT_TIME
      _LAST_PRINT_TIME = LAST_PRINT_TIME
      _UPDATE_TIME = UPDATE_TIME
      _CREATE_TIME = CREATE_TIME
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '物件結束時觸發的事件，用來清除物件的內容
  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
  Private Sub Class_Terminate_Renamed()
    '目的:結束物件
  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal ITEM_LABEL_ID As String) As String
    Try
      Dim key As String = ITEM_LABEL_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsItemLabel
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  'Public Sub Add_Relationship(ByRef objWMS As clsWMSObject)
  '  Try
  '    '挷定WMS的關係                                                                        
  '    If objWMS IsNot Nothing Then
  '      _objWMS = objWMS
  '      objWMS.O_Add_!!!!!這邊就是你要改的東西啦(Me)
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub
  'Public Sub Remove_Relationship()
  '  Try
  '    If _objWMS IsNot Nothing Then
  '      _objWMS.O_Remove_!!!!!這也是你要改的東西
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub
  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_M_Item_LabelManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_M_Item_LabelManagement.GetUpdateSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Delete的SQL
  Public Function O_Add_Delete_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = WMS_M_Item_LabelManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Update_To_Memory(ByRef objWMS_M_Item_Label As clsItemLabel) As Boolean
    Try
      Dim key As String = objWMS_M_Item_Label._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & ",new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _ITEM_LABEL_ID = objWMS_M_Item_Label.ITEM_LABEL_ID
      _ITEM_LABEL_TYPE = objWMS_M_Item_Label.ITEM_LABEL_TYPE
      _PO_ID = objWMS_M_Item_Label.PO_ID
      _TAG1 = objWMS_M_Item_Label.TAG1
      _TAG2 = objWMS_M_Item_Label.TAG2
      _TAG3 = objWMS_M_Item_Label.TAG3
      _TAG4 = objWMS_M_Item_Label.TAG4
      _TAG5 = objWMS_M_Item_Label.TAG5
      _TAG6 = objWMS_M_Item_Label.TAG6
      _TAG7 = objWMS_M_Item_Label.TAG7
      _TAG8 = objWMS_M_Item_Label.TAG8
      _TAG9 = objWMS_M_Item_Label.TAG9
      _TAG10 = objWMS_M_Item_Label.TAG10
      _TAG11 = objWMS_M_Item_Label.TAG11
      _TAG12 = objWMS_M_Item_Label.TAG12
      _TAG13 = objWMS_M_Item_Label.TAG13
      _TAG14 = objWMS_M_Item_Label.TAG14
      _TAG15 = objWMS_M_Item_Label.TAG15
      _TAG16 = objWMS_M_Item_Label.TAG16
      _TAG17 = objWMS_M_Item_Label.TAG17
      _TAG18 = objWMS_M_Item_Label.TAG18
      _TAG19 = objWMS_M_Item_Label.TAG19
      _TAG20 = objWMS_M_Item_Label.TAG20
      _TAG21 = objWMS_M_Item_Label.TAG21
      _TAG22 = objWMS_M_Item_Label.TAG22
      _TAG23 = objWMS_M_Item_Label.TAG23
      _TAG24 = objWMS_M_Item_Label.TAG24
      _TAG25 = objWMS_M_Item_Label.TAG25
      _TAG26 = objWMS_M_Item_Label.TAG26
      _TAG27 = objWMS_M_Item_Label.TAG27
      _TAG28 = objWMS_M_Item_Label.TAG28
      _TAG29 = objWMS_M_Item_Label.TAG29
      _TAG30 = objWMS_M_Item_Label.TAG30
      _TAG31 = objWMS_M_Item_Label.TAG31
      _TAG32 = objWMS_M_Item_Label.TAG32
      _TAG33 = objWMS_M_Item_Label.TAG33
      _TAG34 = objWMS_M_Item_Label.TAG34
      _TAG35 = objWMS_M_Item_Label.TAG35
      _CREATE_USER = objWMS_M_Item_Label.CREATE_USER
      _FISRT_PRINT_TIME = objWMS_M_Item_Label.FIRST_PRINT_TIME
      _LAST_PRINT_TIME = objWMS_M_Item_Label.LAST_PRINT_TIME
      _UPDATE_TIME = objWMS_M_Item_Label.UPDATE_TIME
      _CREATE_TIME = objWMS_M_Item_Label.CREATE_TIME
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
