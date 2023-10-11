

'CosmoWMS給過來的拉料資訊
Class clsCosmoWMS_Outbound

  Public Message As String '主键
  Public Message_DTL As List(Of MSG_Outbound_Info)
End Class
Public Class MSG_Outbound_Info

  Public id As String '主键
  Public ll_no As String '拉料单号
  Public factory_code As String 'Cosmo工厂号
  Public sap_factory_code As String 'SAP工厂号
  Public material_code As String '物料号
  Public amount As Double '拉料数量
  Public send_spot As String '发货地点
  Public wkpos_code As String '接收地点
  Public supplier As String '供应商
End Class