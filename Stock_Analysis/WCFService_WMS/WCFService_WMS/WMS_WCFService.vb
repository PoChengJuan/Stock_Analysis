' 注意: 您可以使用操作功能表上的 [重新命名] 命令同時變更程式碼和組態檔中的類別名稱 "Service1"。
Imports WCFService_WMS

'<ServiceBehavior([Namespace]:="http://www.ecatch.com.tw/")>
<ServiceBehavior([IncludeExceptionDetailInFaults]:=True)>
Public Class WMS_WCFService
  Implements IWMS_WCFService
  Public Shared Event GetWcfCommandEvent_SendInventoryData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendPickingData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendPurchaseData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendSellData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendSKUChangeData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendSKUData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendTransactionData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendWorkData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendWorkIdChangeData(send As Object, e As BaseReply)
  Public Shared Event GetWcfCommandEvent_SendTransferDataToERP(send As Object, e As BaseReply)

  Public Function SendInventoryData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendInventoryData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendInventoryData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendPickingData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendPickingData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendPickingData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendPurchaseData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendPurchaseData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendPurchaseData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendSellData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendSellData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendSellData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendSKUChangeData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendSKUChangeData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendSKUChangeData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendSKUData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendSKUData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendSKUData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendTransactionData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendTransactionData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendTransactionData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendWorkData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendWorkData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendWorkData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendWorkIdChangeData(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendWorkIdChangeData
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendWorkIdChangeData(_XmlString, returnData)
    Return returnData
  End Function

  Public Function SendTransferDataToERP(_XmlString As String) As BaseReply Implements IWMS_WCFService.SendTransferDataToERP
    Dim returnData As New BaseReply
    RaiseEvent GetWcfCommandEvent_SendTransferDataToERP(_XmlString, returnData)
    Return returnData
  End Function
End Class