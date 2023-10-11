Imports eCA_HostObject

Module Module_DataFilter
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Filter~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  '使用Dictionary進行Filter請使用此格式為範本
  ''' <summary>
  ''' 傳入Factory_No、Area_No檢查是否有ProductionInfo
  ''' </summary>
  ''' <param name="dic"></param>
  ''' <param name="Factory_No"></param>
  ''' <param name="Area_No"></param>
  ''' <param name="ret_dic"></param>
  ''' <returns></returns>
  Public Function O_Get_dicProductionInfoByFactoryNo_AreaNo(ByVal dic As Dictionary(Of String, clsProduce_Info),
                                                            ByVal Factory_No As String,
                                                            ByVal Area_No As String,
                                                            Optional ByRef ret_dic As Dictionary(Of String, clsProduce_Info) = Nothing) As Boolean
    Try
      ret_dic = dic.Where(Function(_obj)
                            If Factory_No.Length > 0 AndAlso _obj.Value.Factory_No.Equals(Factory_No) = False Then
                              Return False
                            End If
                            If Area_No.Length > 0 AndAlso _obj.Value.Area_No.Equals(Area_No) = False Then
                              Return False
                            End If
                            Return True
                          End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 取得狀態不是完成的ProductionInfo並進行排序
  ''' </summary>
  ''' <param name="dic"></param>
  ''' <param name="ret_dic"></param>
  ''' <returns></returns>
  Public Function O_Get_dicProductionInfoByStutusNotEnd_Sort(ByVal dic As Dictionary(Of String, clsProduce_Info),
                                                                                                                       Optional ByRef ret_dic As Dictionary(Of String, clsProduce_Info) = Nothing) As Boolean
    Try
      ret_dic = dic.Where(Function(_obj)
                            If _obj.Value.Status.Equals(enuProduceStatus.Queued) = False AndAlso _obj.Value.Status.Equals(enuProduceStatus.Process) = False Then
                              Return False
                            End If
                            Return True
                          End Function).OrderBy(Function(obj) obj.Value.Factory_No & obj.Value.Area_No & obj.Value.Create_Time).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 取得目前生產數量>前次上報的生產數量 OR NG的數量>前次上報NG的數量
  ''' </summary>
  ''' <param name="dic"></param>
  ''' <param name="ret_dic"></param>
  ''' <returns></returns>
  Public Function O_Get_dicProductionInfoByReportQTY_Sort(ByVal dic As Dictionary(Of String, clsProduce_Info),
                                                          Optional ByRef ret_dic As Dictionary(Of String, clsProduce_Info) = Nothing) As Boolean
    Try
      ret_dic = dic.Where(Function(_obj)
                            If (_obj.Value.Qty_Process > _obj.Value.Previous_Qty_Process Or _obj.Value.Qty_NG > _obj.Value.Previous_Qty_NG) And _obj.Value.Start_Time <> "" Then
                              Return True
                            End If
                            Return False
                          End Function).OrderBy(Function(obj) obj.Value.Factory_No & obj.Value.Area_No & obj.Value.Create_Time).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 取得LineProductionInfo
  ''' </summary>
  ''' <param name="dic"></param>
  ''' <param name="Factory_No"></param>
  ''' <param name="Area_No"></param>
  ''' <param name="ret_dic"></param>
  ''' <returns></returns>
  Public Function O_Get_dicCLineProductionInfoByFacotryNo_AreaNo(ByVal dic As Dictionary(Of String, clsLineProduction_Info),
                                                                 ByVal Factory_No As String,
                                                                 ByVal Area_No As String,
                                                                 Optional ByRef ret_dic As Dictionary(Of String, clsLineProduction_Info) = Nothing) As Boolean
    Try
      ret_dic = dic.Where(Function(_obj)
                            If Factory_No.Length > 0 AndAlso _obj.Value.Factory_No.Equals(Factory_No) = False Then
                              Return False
                            End If
                            If Area_No.Length > 0 AndAlso _obj.Value.Area_No.Equals(Area_No) = False Then
                              Return False
                            End If
                            Return True
                          End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 取得的LineProductionInfo
  ''' </summary>
  ''' <param name="dic"></param>
  ''' <param name="Factory_No"></param>
  ''' <param name="Area_NO"></param>
  ''' <param name="Device_NO"></param>
  ''' <param name="Unit_ID"></param>
  ''' <param name="ret_dic"></param>
  ''' <returns></returns>
  Public Function O_Get_dicCLineProductionInfoByFacotryNo_AreaNo_DeviceNo_UnitID(ByVal dic As Dictionary(Of String, clsLineProduction_Info),
                                                                                 ByVal Factory_No As String,
                                                                                 ByVal Area_No As String,
                                                                                 ByVal Device_No As String,
                                                                                 ByVal Unit_ID As String,
                                                                                 Optional ByRef ret_dic As Dictionary(Of String, clsLineProduction_Info) = Nothing) As Boolean
    Try
      ret_dic = dic.Where(Function(_obj)
                            If Factory_No.Length > 0 AndAlso _obj.Value.Factory_No.Equals(Factory_No) = False Then
                              Return False
                            End If
                            If Area_No.Length > 0 AndAlso _obj.Value.Area_No.Equals(Area_No) = False Then
                              Return False
                            End If
                            If Device_No.Length > 0 AndAlso _obj.Value.Device_No.Equals(Device_No) = False Then
                              Return False
                            End If
                            If Unit_ID.Length > 0 AndAlso _obj.Value.Unit_ID.Equals(Unit_ID) = False Then
                              Return False
                            End If
                            Return True
                          End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)
      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 根據Factory and AreaNo 取得上一個生產線區塊
  ''' </summary>
  ''' <param name="dic"></param>
  ''' <param name="Factory_No"></param>
  ''' <param name="Area_No"></param>
  ''' <param name="ret_dic"></param>
  ''' <returns></returns>
  Public Function O_Get_dicLastCLineAreaByFacotryNo_AreaNo(ByVal dic As Dictionary(Of String, clsLine_Area),
                                                           ByVal Factory_No As String,
                                                           ByVal Area_No As String,
                                                           Optional ByRef ret_dic As Dictionary(Of String, clsLine_Area) = Nothing) As Boolean
    Try
      Dim _ret_dic As New Dictionary(Of String, clsLine_Area)
      _ret_dic = dic.Where(Function(_obj)
                             If Factory_No.Length > 0 AndAlso _obj.Value.Factory_No.Equals(Factory_No) = False Then
                               SendMessageToLog("Factory:" & Factory_No, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                               Return False
                             End If
                             If Area_No.Length > 0 AndAlso _obj.Value.Area_No.Equals(Area_No) = False Then
                               SendMessageToLog("Area_No:" & Area_No, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                               Return False
                             End If
                             Return True
                           End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)

      If _ret_dic.Any Then '-取的資料
        If O_Get_dicCLineAreaByFacotryNo_AreaNo_AreaType_AreaIndex(dic, Factory_No, Area_No, _ret_dic.First.Value.AREA_TYPE1, _ret_dic.First.Value.AREA_INDEX - 1, ret_dic) = False Then
          Return False
        Else
          Return True

        End If
      Else
        Return False
      End If

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 根據Factory and AreaNo and Area_type 取得生產區塊資料
  ''' </summary>
  ''' <param name="dic"></param>
  ''' <param name="Factory_No"></param>
  ''' <param name="Area_No"></param>
  ''' <param name="Area_type"></param>
  ''' <param name="ret_dic"></param>
  ''' <returns></returns>
  Public Function O_Get_dicCLineAreaByFacotryNo_AreaNo_AreaType_AreaIndex(ByVal dic As Dictionary(Of String, clsLine_Area),
                                                                ByVal Factory_No As String,
                                                                ByVal Area_No As String,
                                                                ByVal Area_type As Double,
                                                                ByVal Area_index As Double,
                                                                Optional ByRef ret_dic As Dictionary(Of String, clsLine_Area) = Nothing) As Boolean
    Try
      ret_dic = dic.Where(Function(_obj)
                            If Factory_No.Length > 0 AndAlso _obj.Value.Factory_No.Equals(Factory_No) = False Then
                              SendMessageToLog("Factory:" & Factory_No, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                              Return False
                            End If
                            If Area_No.Length > 0 AndAlso _obj.Value.Area_No.Equals(Area_No) = False Then
                              SendMessageToLog("Area_No:" & Area_No, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                              Return False
                            End If
                            If _obj.Value.Area_Type2.Equals(Area_type) = False Then
                              SendMessageToLog("Area_type:" & Area_type, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                              Return False
                            End If
                            If _obj.Value.AREA_INDEX.Equals(Area_index) = False Then
                              SendMessageToLog("Area_index:" & Area_index, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                              Return False
                            End If
                            Return True
                          End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)

      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 根據POID and SKU 取得PO 明細
  ''' </summary>
  ''' <param name="dic"></param>
  ''' <param name="PO_ID"></param>
  ''' <param name="SKU"></param>	
  ''' <param name="ret_dic"></param>
  ''' <returns></returns>
  Public Function O_Get_dicPODTLByPOID_SKU(ByVal dic As Dictionary(Of String, clsPO_DTL),
                                          ByVal Po_ID As String,
                                          ByVal SKU As String,
                                          Optional ByRef ret_dic As Dictionary(Of String, clsPO_DTL) = Nothing) As Boolean
    Try
      ret_dic = dic.Where(Function(_obj)
                            If Po_ID.Length > 0 AndAlso _obj.Value.PO_ID.Equals(Po_ID) = False Then
                              SendMessageToLog("Po_ID:" & Po_ID, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                              Return False
                            End If
                            If SKU.Length > 0 AndAlso _obj.Value.SKU_NO.Equals(SKU) = False Then
                              SendMessageToLog("SKU_No:" & SKU, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                              Return False
                            End If
                            Return True
                          End Function).ToDictionary(Function(_obj) _obj.Key, Function(_obj) _obj.Value)

      If ret_dic.Any Then
        Return True
      End If
      Return False
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module
