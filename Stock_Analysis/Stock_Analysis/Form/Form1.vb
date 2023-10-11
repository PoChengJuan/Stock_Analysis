Public Class Form1
  Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
    STDIN.ProdID = "WMS"
    STDIN.Companyid = "MOSA_TEST_LINDA"
    STDIN.Userid = "DS"
    STDIN.DoAction = "1"
    STDIN.Docase = "1"

    'Data
    Dim Data As New eCA_TransactionMessage.STD_INData
    Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
    FormHead.TableName = "SFCTB"

    '組成Header
    Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
    Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
    Rocord_Head_Info.TB003 = TextBox1.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 
    Rocord_Head_Info.TB004 = TextBox2.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別
    Rocord_Head_Info.TB005 = TextBox_TB005.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別
    Rocord_Head_Info.TB007 = TextBox3.Text.ToString 'eCA_WMSObject.enuTransferInType.Processing_Manufacturer '"TransferInType" 移入類別
    Rocord_Head_Info.TB008 = TextBox_TB008.Text.ToString 'eCA_WMSObject.enuTransferInType.Processing_Manufacturer '"TransferInType" 移入類別
    Rocord_Head_Info.TB010 = TextBox4.Text.ToString ' "F01" ' "FactoryId"
    Rocord_Head(0) = Rocord_Head_Info

    FormHead.RecordList() = Rocord_Head
    Data.FormHead = FormHead



    '組成Boby
    Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
    FormBody.TableName = "SFCTC"

    Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
    Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
    Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
    Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
    Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
    Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
    Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
    Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
    Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
    Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
    Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
    Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
    Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"
    Rocord_Body_Info.TC200 = TextBox_TC200.Text.ToString
    Rocord_Body_Info.TC201 = TextBox_TC201.Text.ToString

    Rocord_Body(0) = Rocord_Body_Info

    FormBody.RecordList = Rocord_Body
    Data.FormBody = FormBody


    STDIN.Data = Data


    '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

    Dim Result_Message = ""
    'O_SendTransferDataToERP(STDIN, Result_Message)

    'STD_IN(STDIN)
  End Sub

  Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    Try
      If RadioButton_COPTG.Checked Then COPTG()
      If RadioButton_INVMB.Checked Then INVMB()
      If RadioButton_INVTA.Checked Then INVTA()
      If RadioButton_INVTE.Checked Then INVTE()
      If RadioButton_INVTL.Checked Then INVTL()
      If RadioButton_MOCTA.Checked Then MOCTA()
      If RadioButton_MOCTC.Checked Then MOCTC()
      If RadioButton_MOCTO.Checked Then MOCTO()
      If RadioButton_PURTG.Checked Then PURTG()
      If RadioButton_SFCTB.Checked Then SFCTB()
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub COPTG()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "COPTG"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>COPTG</TableName>
      '  <RecordList>
      '    <TG001>銷貨單別</TG001>        
      '    <TG002>銷貨單號</TG002>
      '    <TG200>WMS接收成功</TG200> 0:已接收未放行.-1:已接收已放行0.1: 有問題
      '  </RecordList>
      Rocord_Head_Info.TG001 = TextBox_TG001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 
      Rocord_Head_Info.TG002 = TextBox_TG002.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別
      Rocord_Head_Info.TG200 = TextBox_TG200.Text.ToString 'eCA_WMSObject.enuTransferInType.Processing_Manufacturer '"TransferInType" 移入類別			
      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub INVMB()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "INVMB"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>INVMB</TableName>
      '   <RecordList>
      '     <MB001>品號</MB001>    
      '     <MB201>WMS接收成功</MB201>       
      '   </RecordList>
      Rocord_Head_Info.MB001 = TextBox_MB001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 
      Rocord_Head_Info.MB201 = TextBox_MB201.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別

      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub INVTA()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "INVTA"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>INVTA</TableName>
      '    <RecordList>
      '      <TA001>庫存異動單別</TA001>    
      '      <TA002>庫存異動單號</TA002>    
      '      <TA200>WMS接收成功</TA003>               
      '    </RecordList>
      Rocord_Head_Info.TA001 = TextBox_TA001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 
      Rocord_Head_Info.TA002 = TextBox_TA002.Text.ToString
      Rocord_Head_Info.TA200 = TextBox_TA200.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別

      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub MOCTO()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "MOCTO"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>MOCTO</TableName>
      '   <RecordList>
      '     <T0001>製令變更單別</TO001>    
      '     <T0002>製令變更單號</TO002>
      '     <TO003>變更版次</TO003>  
      '     <T0200>WMS接收成功</T0200>                 
      '   </RecordList>			
      Rocord_Head_Info.TO001 = TextBox_T0001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TO002 = TextBox_T0002.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TO003 = TextBox_T0003.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TO200 = TextBox_T0200.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			

      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub MOCTC()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "MOCTC"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>MOCTC</TableName>
      '  <RecordList>
      '    <TC001>領退料單別</TC001>    
      '    <TC002>領退料單號</TC002>  
      '    <TC200>WMS接收成功</TC200>               
      '  </RecordList>
      Rocord_Head_Info.TC001 = TextBox_TC001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TC002 = TextBox_TC002.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TC200 = TextBox_TC200.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別

      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub PURTG()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "PURTG"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>PURTG</TableName>
      '   <RecordList>
      '     <TG001>進貨單單別</TG001>    
      '     <TG002>進貨單單號</TG002>
      '     <TG200>WMS接收成功</TG200>                 
      '   </RecordList>      
      Rocord_Head_Info.TG001 = TextBox_TG001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TG002 = TextBox_TG002.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 						
      Rocord_Head_Info.TG200 = TextBox_TG200.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			

      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub SFCTB()
    Try
      Button1_Click(Nothing, Nothing)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub MOCTA()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "MOCTA"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>MOCTA</TableName>
      '   <RecordList>
      '     <TA001>製令單單別</TA001>    
      '     <TA002>製令單單號</TA002>  
      '     <TA201>WMS接收成功</TA201>               
      '   </RecordList>
      Rocord_Head_Info.TA001 = TextBox_TA001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TA002 = TextBox_TA002.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TA201 = TextBox_TA201.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別

      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub INVTE()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "INVTE"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>INVTE</TableName>
      '  <RecordList>
      '    <TE001>盤點底稿編號</TE001>
      '    <TE200>WMS接收成功</TE200>                   
      '  </RecordList> 
      Rocord_Head_Info.TE001 = TextBox_TE001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TE200 = TextBox_TE200.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別

      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Sub INVTL()
    Try
      Dim STDIN As New eCA_TransactionMessage.MSG_SendTransferDataToERP
      STDIN.ProdID = "WMS"
      STDIN.Companyid = "MOSA_TEST_LINDA"
      STDIN.Userid = "DS"
      STDIN.DoAction = "2"
      STDIN.Docase = "1"

      'Data
      Dim Data As New eCA_TransactionMessage.STD_INData
      Dim FormHead As New eCA_TransactionMessage.STD_INDataFormHead
      FormHead.TableName = "INVTL"

      '組成Header
      Dim Rocord_Head(0) As eCA_TransactionMessage.STD_INDataFormHeadRecordList
      Dim Rocord_Head_Info As New eCA_TransactionMessage.STD_INDataFormHeadRecordList
      '<TableName>INVTL</TableName>
      '  <RecordList>
      '    <TL001>品號</TL001> 
      '    <TL004>變更版次</TL004>
      '    <TL200>WMS接收成功</TL200>                   
      '  </RecordList>  
      Rocord_Head_Info.TL001 = TextBox_TL001.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TL004 = TextBox_TL004.Text.ToString 'ModuleHelpFunc.GetNewTime_ByDataTimeFormat("yyyymmdd") '"TransferDateTime" 			
      Rocord_Head_Info.TL200 = TextBox_TL200.Text.ToString ' eCA_WMSObject.enuTransferOutType.Production_Line '"TransferOutType" 移出類別

      Rocord_Head(0) = Rocord_Head_Info

      FormHead.RecordList() = Rocord_Head
      Data.FormHead = FormHead



      '組成Boby
      Dim FormBody As New eCA_TransactionMessage.STD_INDataFormBody
      FormBody.TableName = "" '"SFCTC"

      Dim Rocord_Body(0) As eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Dim Rocord_Body_Info As New eCA_TransactionMessage.STD_INDataFormBodyRecordList
      Rocord_Body_Info.TC004 = TextBox5.Text.ToString ' "WorkType"
      Rocord_Body_Info.TC005 = TextBox6.Text.ToString '"WorkId" & ModuleHelpFunc.GetNewTime_DBFormat
      Rocord_Body_Info.TC006 = TextBox7.Text.ToString '"TransferOutStage"
      Rocord_Body_Info.TC007 = TextBox8.Text.ToString '"TransferOutProcess"
      Rocord_Body_Info.TC008 = TextBox9.Text.ToString '"TransferInStage"
      Rocord_Body_Info.TC009 = TextBox14.Text.ToString '"TransferInProcess"
      Rocord_Body_Info.TC010 = TextBox13.Text.ToString '"Unit"
      Rocord_Body_Info.TC014 = TextBox12.Text.ToString '"CheckQty"
      Rocord_Body_Info.TC016 = TextBox11.Text.ToString '"ScrapQty"
      Rocord_Body_Info.TC020 = TextBox10.Text.ToString  '"UnderManTime"
      Rocord_Body_Info.TC021 = TextBox15.Text.ToString '"UnderMachineTime"

      Rocord_Body(0) = Rocord_Body_Info

      FormBody.RecordList = Rocord_Body
      Data.FormBody = FormBody


      STDIN.Data = Data


      '模擬WMS送訊息過來 'Host轉傳後根據結果作DB的填寫

      Dim Result_Message = ""
      'O_SendTransferDataToERP(STDIN, Result_Message)

      'STD_IN(STDIN)
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

  End Sub

End Class