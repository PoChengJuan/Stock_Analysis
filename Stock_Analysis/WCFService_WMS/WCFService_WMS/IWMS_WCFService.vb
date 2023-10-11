Imports System.Xml.Serialization
Imports System.Xml
' 注意: 您可以使用操作功能表上的 [重新命名] 命令同時變更程式碼和組態檔中的介面名稱 "IService1"。
'<ServiceContract(Namespace:="uri:/brooks.com/classMCS/XMLSoap/chimeiParticleCount")>
<ServiceContract()>
Public Interface IWMS_WCFService

  ''' <summary>盤點</summary>
  <OperationContract()>
  Function SendInventoryData(ByVal _XmlString As String) As BaseReply

  ''' <summary>領料單/退料單</summary>
  <OperationContract()>
  Function SendPickingData(ByVal _XmlString As String) As BaseReply

  ''' <summary>進貨單</summary>
  <OperationContract()>
  Function SendPurchaseData(ByVal _XmlString As String) As BaseReply

  ''' <summary>銷貨單據</summary>
  <OperationContract()>
  Function SendSellData(ByVal _XmlString As String) As BaseReply

  ''' <summary>機台狀態(WMS提供)</summary>
  <OperationContract()>
  Function SendSKUChangeData(ByVal _XmlString As String) As BaseReply

  ''' <summary>品號資料</summary>
  <OperationContract()>
  Function SendSKUData(ByVal _XmlString As String) As BaseReply

  ''' <summary>庫存異動單據</summary>
  <OperationContract()>
  Function SendTransactionData(ByVal _XmlString As String) As BaseReply

  ''' <summary>移轉單</summary>
  <OperationContract()>
  Function SendTransferDataToERP(ByVal _XmlString As String) As BaseReply

  ''' <summary>製令單</summary>
  <OperationContract()>
  Function SendWorkData(ByVal _XmlString As String) As BaseReply

  ''' <summary>製令單變更</summary>
  <OperationContract()>
  Function SendWorkIdChangeData(ByVal _XmlString As String) As BaseReply
End Interface

' 使用下列範例中所示的資料合約，新增複合型別至服務作業。
' 您可以將 XSD 檔案加入專案。建置專案後，您可以直接以命名空間 "WcfServiceLibrary.ContractType" 使用該處定義的資料型別。

<DataContract()>
<XmlRoot(ElementName:="eWMSMessage")>
Public Class BaseReply
  <DataMember()>
  Public Property ResultCode As Integer = 0

  <DataMember()>
  Public Property ResultMessage As String = "OK"
End Class