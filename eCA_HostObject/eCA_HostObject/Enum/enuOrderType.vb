Public Enum enuOrderType
  None = 0

  Unknow = 9999 '碰到不明項目時使用

  '訂單類型(中類)(暫定等於H_PO_ORDER_TYPE以後可能再進行調整)
#Region "100 區 入庫"
  Inbound_Data = 101 '101.原料入庫(進貨單)(採購入庫單)
  Pick_Material_back = 102  '102.退料單(領料後的退料)

  Product_In = 121 '121.成品入庫(生產入庫單)

  Sell_Back = 131 '131.退貨入庫

  transfer_in = 144 '144 調撥入庫

  Other_In = 151 '151.雜收單

  'picking_SKU_in = 133 ' O 133.領料成品入庫(退料單不含製令單)
  'semiSKU_in = 111 ' X 111.半成品入庫
  'other_in = 131 ' X 131.其他入庫(根據各廠定義)
  normal_in = 145 '庫存異動 入 145
  'temp_in = 146 '暫入 -> 145
  'temp_out_return = 147 '暫出歸還
  'picking_material_in = 132 ' O 132.領料原料入庫(退料單含製令單)
#End Region

#Region "300 區 出庫"
  Material_Out = 301 '301.原料出庫
  Picking_Material_Out = 302 '302.領料出庫

  Sell_Out = 321 '321.成品出庫(銷貨單)

  transfer_out = 344 '344 調撥出庫

  Oher_Out = 351 '351.雜發單

  'produce_in_before = 201 ' O 201.產線入庫前製程(製令單)
  'produce_in_after = 202 ' O 202.產線入庫後製程(領料單中的製令單)
  'material_out = 301 ' X 301.原料出庫
  'semiSKU_out = 311 ' X 311.半成品出庫
  'SKU_out = 321 ' O 321.成品出庫(銷貨單)
  'picking_SKU_out = 331 ' O 331.領料成品出庫(領料單不含製令單)

  'temp_out = 346 '暫出 -> 345
  'asfi514 = 345     ' 工單領料維護作業             ->345
  normal_out = 345  ' 庫存異動 出  -> 345
  'temp_in_return = 347  '暫入歸還
  'produce_out = 401 ' X 401.產線出庫
  'picking_material_out = 431 ' O 431.領料原料出庫(領料單含製令單)
#End Region

#Region "其他類型"
  Stocktaking = 501 ' O 501.盤點單(盤點單)
  Transaction_in = 601 ' O 601.轉撥入庫(庫存異動一般)
  Transaction_out = 621 ' O 621.轉撥出庫(庫存異動一般)
  Transaction_account = 631 ' O 631.轉撥(庫存異動帳轉)
  Change_Stock = 701  'O 701.調撥單(庫內庫存異動)
  Change_Out = 702  'O 702.調撥出庫(透過UI操作的調撥出庫，真的有出庫動作)
#End Region



#Region "原本使用的"
  '未確認
  庫存異動單據 = 999

  '訂單類型(中類)(暫定等於H_PO_ORDER_TYPE以後可能再進行調整)
  'material_in = 101 ' O 101.原料入庫(進貨單)
  'semiSKU_in = 111 ' X 111.半成品入庫
  'SKU_in = 121 ' X 121.成品入庫
  'other_in = 131 ' X 131.其他入庫(根據各廠定義)
  'picking_material_in = 132 ' O 132.領料原料入庫(退料單含製令單)
  'picking_SKU_in = 133 ' O 133.領料成品入庫(退料單不含製令單)
  'back_in = 141 '退貨入庫
  'transaction_in = 144 '工單欠料補料單維護作業
  'temp_in = 145 '暫入
  'asfi526 = 145 '工單成套退料單維護作業
  'asfi527 = 146  '工單超領退料單維護作業 ->146 
  'normal_in = 146  ' 一般單據 入  -> 146
  'temp_out_return = 147 '暫出歸還 147
  'asfi529 = 147    '工單領退料維護作業         ->147
  'produce_in_before = 201 ' O 201.產線入庫前製程(製令單)(領料單中的製令單)
  'produce_in_after = 202 ' O 202.產線入庫後製程(製令單)(領料單中的製令單)
  'material_out = 301 ' X 301.原料出庫
  'semiSKU_out = 311 ' X 311.半成品出庫
  'SKU_out = 321 ' O 321.成品出庫(銷貨單)
  'picking_SKU_out = 331 ' O 331.領料成品出庫(領料單不含製令單)
  'transaction_out = 344   ' 工單超領發料單維護作業 ->344
  'temp_out = 345 '暫入
  'asfi514 = 345     ' 工單領料維護作業             ->345
  'normal_out = 346  ' 一般單據 出  -> 346
  'temp_in_return = 347 '暫入歸還 347
  'produce_out = 401 ' X 401.產線出庫
  'picking_material_out = 431 ' O 431.領料原料出庫(領料單含製令單)
  'stocktaking = 501 ' O 501.盤點單(盤點單)
  ''transaction_in = 601 ' O 601.轉撥入庫(庫存異動一般)
  ''transaction_out = 621 ' O 621.轉撥出庫(庫存異動一般)
  'transaction_account = 631 ' O 631.轉撥(庫存異動帳轉)

  ''台盈
  'Inbound_Data = 102
  'ProduceInData = 103
  'OtherInData = 104
  'SellReturn = 105
  'Replenishment = 701
  'm_Replenishment = 702
  'OtherOutData = 303
  'SellData = 304
  'InboundReturn_Data = 305
  'm_Inbound_Data = 181
  'm_Outbound_Data = 381
  ''手動單據(WMS內部單據)
  m_material_in = 151 ' X 151.原料入庫(手工單)
  m_semiSKU_in = 161 ' X 161.半成品入庫(手工單)
  m_SKU_in = 171 ' X 171.成品入庫(手工單)
  m_general_in = 181 ' O 181.通用入庫(手工單)(不分原料成品)
  m_material_out = 351 ' X 351.原料出庫(手工單)
  m_semiSKU_out = 361 ' X 361.半成品出庫(手工單)
  m_SKU_out = 371 ' X 371.成品出庫(手工單)
  m_grneral_out = 381 ' O 381.通用出庫(手工單)(不分原料成品)"
#End Region
End Enum
