Imports System.Xml.Serialization

'<XmlRoot(ElementName:="STD_IN")>
'Public Class clsSendTransferDataToERP
'    Public Property ProdID As String
'    Public Property Companyid As String
'    Public Property Userid As String
'    Public Property DoAction As String
'    Public Property Docase As String

'    <XmlArray(ElementName:="Data")>
'    Public Property Data As New List(Of DataInfo)

'    Public Class DataInfo
'        Public Property DOC_ID As String
'        Public Property DOC_NO_SN As String
'        Public Property WMS_TRANS_QTY As String
'    End Class
'End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False, ElementName:="STD_IN")>
Partial Public Class MSG_SendTransferDataToERP

  Private prodIDField As String

  Private companyidField As String

  Private useridField As String

  Private resultField As String

  Private exceptionField As String = ""

  Private MessageField As String = ""

  Private doActionField As String = ""

  Private docaseField As Byte

  Private dataField As STD_INData


  Public Property Exception() As String
    Get
      Return Me.exceptionField
    End Get
    Set
      Me.exceptionField = Value
    End Set
  End Property

  Public Property Message() As String
    Get
      Return Me.MessageField
    End Get
    Set
      Me.MessageField = Value
    End Set
  End Property

  Public Property Result() As String
    Get
      Return Me.resultField
    End Get
    Set
      Me.resultField = Value
    End Set
  End Property


  '''<remarks/>
  Public Property ProdID() As String
    Get
      Return Me.prodIDField
    End Get
    Set
      Me.prodIDField = Value
    End Set
  End Property

  '''<remarks/>
  Public Property Companyid() As String
    Get
      Return Me.companyidField
    End Get
    Set
      Me.companyidField = Value
    End Set
  End Property

  '''<remarks/>
  Public Property Userid() As String
    Get
      Return Me.useridField
    End Get
    Set
      Me.useridField = Value
    End Set
  End Property

  '''<remarks/>
  Public Property DoAction() As String
    Get
      Return Me.doActionField
    End Get
    Set
      Me.doActionField = Value
    End Set
  End Property

  '''<remarks/>
  Public Property Docase() As Byte
    Get
      Return Me.docaseField
    End Get
    Set
      Me.docaseField = Value
    End Set
  End Property

  '''<remarks/>
  Public Property Data() As STD_INData
    Get
      Return Me.dataField
    End Get
    Set
      Me.dataField = Value
    End Set
  End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class STD_INData

    Private formHeadField As New STD_INDataFormHead

    Private formBodyField As New STD_INDataFormBody

    '''<remarks/>
    Public Property FormHead() As STD_INDataFormHead
        Get
            Return Me.formHeadField
        End Get
        Set
            Me.formHeadField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property FormBody() As STD_INDataFormBody
        Get
            Return Me.formBodyField
        End Get
        Set
            Me.formBodyField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class STD_INDataFormHead

    Private tableNameField As String

    Private recordListField() As STD_INDataFormHeadRecordList

    '''<remarks/>
    Public Property TableName() As String
        Get
            Return Me.tableNameField
        End Get
        Set
            Me.tableNameField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlElementAttribute("RecordList")>
    Public Property RecordList() As STD_INDataFormHeadRecordList()
        Get
            Return Me.recordListField
        End Get
        Set
            Me.recordListField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class STD_INDataFormHeadRecordList

  Private tF001Field As String
  Private tF002Field As String
  Private tF200Field As String

  Private tG001Field As String
  Private tG002Field As String
  Private tG200Field As String

  Private tH001Field As String
  Private tH002Field As String
  Private tH200Field As String

  Private mB001Field As String
  Private mB201Field As String

  Private tA001Field As String
  Private tA002Field As String
  Private tA200Field As String
  Private tA201Field As String

  Private tO001Field As String
  Private tO002Field As String
  Private tO003Field As String
  Private tO200Field As String

  Private tC001Field As String
  Private tC002Field As String
  Private tC200Field As String

  Private tE001Field As String
  Private tE200Field As String

  Private tL001Field As String
  Private tL004Field As String
  Private tL200Field As String

  Private tB003Field As String
  Private tB004Field As String
  Private tB005Field As String
  Private tB007Field As String
  Private tB008Field As String
  Private tB010Field As String

  Private mO001Field As String
  Private mO200Field As String

  Private NWHDTField As String
  Private NWHRSField As String

  '採購單
  Private TPODTField As String
  Private TPONOField As String
  Private TPORSField As String
  '領料單
  Private TWIDTField As String
  Private TWINOField As String
  Private TWIRSField As String
  '退料單
  Private TWRDTField As String
  Private TWRNOField As String
  Private TWRRSField As String
  '生產入庫單
  Private TWSDTField As String
  Private TWSNOField As String
  Private TWSRSField As String
  '調撥單
  Private TWTDTField As String
  Private TWTNOField As String
  Private TWTRSField As String
  '雜發單
  Private TXIDTField As String
  Private TXINOField As String
  Private TXIRSField As String
  '雜收單
  Private TXSDTField As String
  Private TXSNOField As String
  Private TXSRSField As String
  '銷售單
  Private TDNDTField As String
  Private TDNNOField As String
  Private TDNRSField As String
  '銷退單
  Private ST001Field As String
  Private ST002Field As String
  Private ST003Field As String
  '退供商單
  Private PT001Field As String
  Private PT002Field As String
  Private PT003Field As String
  '貨主調撥單
  Private INT01Field As String
  Private INT02Field As String
  Private INT03Field As String



  '''<remarks/>
  Public Property MO001() As String
    Get
      Return Me.mO001Field
    End Get
    Set
      Me.mO001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property MO200() As String
    Get
      Return Me.mO200Field
    End Get
    Set
      Me.mO200Field = Value
    End Set
  End Property
  Public Property NWHDT() As String
    Get
      Return NWHDTField
    End Get
    Set(ByVal value As String)
      NWHDTField = value
    End Set
  End Property
  Public Property NWHRS() As String
    Get
      Return NWHRSField
    End Get
    Set(ByVal value As String)
      NWHRSField = value
    End Set
  End Property
  '''<remarks/>
  Public Property TL001() As String
    Get
      Return Me.tL001Field
    End Get
    Set
      Me.tL001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TL004() As String
    Get
      Return Me.tL004Field
    End Get
    Set
      Me.tL004Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TL200() As String
    Get
      Return Me.tL200Field
    End Get
    Set
      Me.tL200Field = Value
    End Set
  End Property


  '''<remarks/>
  Public Property TE001() As String
    Get
      Return Me.tE001Field
    End Get
    Set
      Me.tE001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TE200() As String
    Get
      Return Me.tE200Field
    End Get
    Set
      Me.tE200Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TC001() As String
    Get
      Return Me.tC001Field
    End Get
    Set
      Me.tC001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TC002() As String
    Get
      Return Me.tC002Field
    End Get
    Set
      Me.tC002Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TC200() As String
    Get
      Return Me.tC200Field
    End Get
    Set
      Me.tC200Field = Value
    End Set
  End Property


  '''<remarks/>
  Public Property TO001() As String
    Get
      Return Me.tO001Field
    End Get
    Set
      Me.tO001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TO002() As String
    Get
      Return Me.tO002Field
    End Get
    Set
      Me.tO002Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TO003() As String
    Get
      Return Me.tO003Field
    End Get
    Set
      Me.tO003Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TO200() As String
    Get
      Return Me.tO200Field
    End Get
    Set
      Me.tO200Field = Value
    End Set
  End Property


  '''<remarks/>
  Public Property TA001() As String
    Get
      Return Me.tA001Field
    End Get
    Set
      Me.tA001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TA002() As String
    Get
      Return Me.tA002Field
    End Get
    Set
      Me.tA002Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TA200() As String
    Get
      Return Me.tA200Field
    End Get
    Set
      Me.tA200Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TA201() As String
    Get
      Return Me.tA201Field
    End Get
    Set
      Me.tA201Field = Value
    End Set
  End Property



  '''<remarks/>
  Public Property MB001() As String
    Get
      Return Me.mB001Field
    End Get
    Set
      Me.mB001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property MB201() As String
    Get
      Return Me.mB201Field
    End Get
    Set
      Me.mB201Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TF001() As String
    Get
      Return Me.tF001Field
    End Get
    Set
      Me.tF001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TF002() As String
    Get
      Return Me.tF002Field
    End Get
    Set
      Me.tF002Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TF200() As String
    Get
      Return Me.tF200Field
    End Get
    Set
      Me.tF200Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TG001() As String
    Get
      Return Me.tG001Field
    End Get
    Set
      Me.tG001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TG002() As String
    Get
      Return Me.tG002Field
    End Get
    Set
      Me.tG002Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TG200() As String
    Get
      Return Me.tG200Field
    End Get
    Set
      Me.tG200Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TH001() As String
    Get
      Return Me.tH001Field
    End Get
    Set
      Me.tH001Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TH002() As String
    Get
      Return Me.tH002Field
    End Get
    Set
      Me.tH002Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TH200() As String
    Get
      Return Me.tH200Field
    End Get
    Set
      Me.tH200Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TB003() As String
    Get
      Return Me.tB003Field
    End Get
    Set
      Me.tB003Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TB004() As String
    Get
      Return Me.tB004Field
    End Get
    Set
      Me.tB004Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TB005() As String
    Get
      Return Me.tB005Field
    End Get
    Set
      Me.tB005Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TB007() As String
    Get
      Return Me.tB007Field
    End Get
    Set
      Me.tB007Field = Value
    End Set
  End Property
  '''<remarks/>
  Public Property TB008() As String
    Get
      Return Me.tB008Field
    End Get
    Set
      Me.tB008Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TB010() As String
    Get
      Return Me.tB010Field
    End Get
    Set
      Me.tB010Field = Value
    End Set
  End Property
  Public Property TPODT() As String
    Get
      Return Me.TPODTField
    End Get
    Set(value As String)
      Me.TPODTField = value
    End Set
  End Property
  Public Property TPONO() As String
    Get
      Return Me.TPONOField
    End Get
    Set(value As String)
      Me.TPONOField = value
    End Set
  End Property
  Public Property TPORS() As String
    Get
      Return Me.TPORSField
    End Get
    Set(value As String)
      Me.TPORSField = value
    End Set
  End Property
  Public Property TWIDT() As String
    Get
      Return Me.TWIDTField
    End Get
    Set(value As String)
      Me.TWIDTField = value
    End Set
  End Property
  Public Property TWINO() As String
    Get
      Return Me.TWINOField
    End Get
    Set(value As String)
      Me.TWINOField = value
    End Set
  End Property
  Public Property TWIRS() As String
    Get
      Return Me.TWIRSField
    End Get
    Set(value As String)
      Me.TWIRSField = value
    End Set
  End Property
  Public Property TWRDT() As String
    Get
      Return Me.TWRDTField
    End Get
    Set(value As String)
      Me.TWRDTField = value
    End Set
  End Property
  Public Property TWRNO() As String
    Get
      Return Me.TWRNOField
    End Get
    Set(value As String)
      Me.TWRNOField = value
    End Set
  End Property
  Public Property TWRRS() As String
    Get
      Return Me.TWRRSField
    End Get
    Set(value As String)
      Me.TWRRSField = value
    End Set
  End Property
  Public Property TWSDT() As String
    Get
      Return Me.TWSDTField
    End Get
    Set(value As String)
      Me.TWSDTField = value
    End Set
  End Property
  Public Property TWSNO() As String
    Get
      Return Me.TWSNOField
    End Get
    Set(value As String)
      Me.TWSNOField = value
    End Set
  End Property
  Public Property TWSRS() As String
    Get
      Return Me.TWSRSField
    End Get
    Set(value As String)
      Me.TWSRSField = value
    End Set
  End Property
  Public Property TWTDT() As String
    Get
      Return Me.TWTDTField
    End Get
    Set(value As String)
      Me.TWTDTField = value
    End Set
  End Property
  Public Property TWTNO() As String
    Get
      Return Me.TWTNOField
    End Get
    Set(value As String)
      Me.TWTNOField = value
    End Set
  End Property
  Public Property TWTRS() As String
    Get
      Return Me.TWTRSField
    End Get
    Set(value As String)
      Me.TWTRSField = value
    End Set
  End Property
  Public Property TXIDT() As String
    Get
      Return Me.TXIDTField
    End Get
    Set(value As String)
      Me.TXIDTField = value
    End Set
  End Property
  Public Property TXINO() As String
    Get
      Return Me.TXINOField
    End Get
    Set(value As String)
      Me.TXINOField = value
    End Set
  End Property
  Public Property TXIRS() As String
    Get
      Return Me.TXIRSField
    End Get
    Set(value As String)
      Me.TXIRSField = value
    End Set
  End Property
  Public Property TXSDT() As String
    Get
      Return Me.TXSDTField
    End Get
    Set(value As String)
      Me.TXSDTField = value
    End Set
  End Property
  Public Property TXSNO() As String
    Get
      Return Me.TXSNOField
    End Get
    Set(value As String)
      Me.TXSNOField = value
    End Set
  End Property
  Public Property TXSRS() As String
    Get
      Return Me.TXSRSField
    End Get
    Set(value As String)
      Me.TXSRSField = value
    End Set
  End Property
  Public Property TDNDT() As String
    Get
      Return Me.TDNDTField
    End Get
    Set(value As String)
      Me.TDNDTField = value
    End Set
  End Property
  Public Property TDNNO() As String
    Get
      Return Me.TDNNOField
    End Get
    Set(value As String)
      Me.TDNNOField = value
    End Set
  End Property
  Public Property TDNRS() As String
    Get
      Return Me.TDNRSField
    End Get
    Set(value As String)
      Me.TDNRSField = value
    End Set
  End Property
  Public Property ST001() As String
    Get
      Return Me.ST001Field
    End Get
    Set(ByVal value As String)
      Me.ST001Field = value
    End Set
  End Property
  Public Property ST002() As String
    Get
      Return Me.ST002Field
    End Get
    Set(ByVal value As String)
      Me.ST002Field = value
    End Set
  End Property
  Public Property ST003() As String
    Get
      Return Me.ST003Field
    End Get
    Set(ByVal value As String)
      Me.ST003Field = value
    End Set
  End Property
  Public Property PT001() As String
    Get
      Return Me.PT001Field
    End Get
    Set(ByVal value As String)
      Me.PT001Field = value
    End Set
  End Property
  Public Property PT002() As String
    Get
      Return Me.PT002Field
    End Get
    Set(ByVal value As String)
      Me.PT002Field = value
    End Set
  End Property
  Public Property PT003() As String
    Get
      Return Me.PT003Field
    End Get
    Set(ByVal value As String)
      Me.PT003Field = value
    End Set
  End Property
  Public Property INT01() As String
    Get
      Return Me.INT01Field
    End Get
    Set(ByVal value As String)
      Me.INT01Field = value
    End Set
  End Property
  Public Property INT02() As String
    Get
      Return Me.INT02Field
    End Get
    Set(ByVal value As String)
      Me.INT02Field = value
    End Set
  End Property
  Public Property INT03() As String
    Get
      Return Me.INT03Field
    End Get
    Set(ByVal value As String)
      Me.INT03Field = value
    End Set
  End Property
End Class



'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class STD_INDataFormBody

    Private tableNameField As String

    Private recordListField() As STD_INDataFormBodyRecordList

    '''<remarks/>
    Public Property TableName() As String
        Get
            Return Me.tableNameField
        End Get
        Set
            Me.tableNameField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlElementAttribute("RecordList")>
    Public Property RecordList() As STD_INDataFormBodyRecordList()
        Get
            Return Me.recordListField
        End Get
        Set
            Me.recordListField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class STD_INDataFormBodyRecordList

    Private tC004Field As String

    Private tC005Field As String

    Private tC006Field As String

    Private tC007Field As String

    Private tC008Field As String

    Private tC009Field As String

    Private tC010Field As String

    Private tC014Field As String

    Private tC016Field As String

    Private tC020Field As String

	Private tC021Field As String

	Private tC200Field As String

	Private tC201Field As String

	'''<remarks/>
	Public Property TC004() As String
        Get
            Return Me.tC004Field
        End Get
        Set
            Me.tC004Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC005() As String
        Get
            Return Me.tC005Field
        End Get
        Set
            Me.tC005Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC006() As String
        Get
            Return Me.tC006Field
        End Get
        Set
            Me.tC006Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC007() As String
        Get
            Return Me.tC007Field
        End Get
        Set
            Me.tC007Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC008() As String
        Get
            Return Me.tC008Field
        End Get
        Set
            Me.tC008Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC009() As String
        Get
            Return Me.tC009Field
        End Get
        Set
            Me.tC009Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC010() As String
        Get
            Return Me.tC010Field
        End Get
        Set
            Me.tC010Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC014() As String
        Get
            Return Me.tC014Field
        End Get
        Set
            Me.tC014Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC016() As String
        Get
            Return Me.tC016Field
        End Get
        Set
            Me.tC016Field = Value
        End Set
    End Property

    '''<remarks/>
    Public Property TC020() As String
        Get
            Return Me.tC020Field
        End Get
        Set
            Me.tC020Field = Value
        End Set
    End Property

	'''<remarks/>
	Public Property TC021() As String
		Get
			Return Me.tC021Field
		End Get
		Set
			Me.tC021Field = Value
		End Set
	End Property
	'''<remarks/>
	Public Property TC200() As String
		Get
			Return Me.tC200Field
		End Get
		Set
			Me.tC200Field = Value
		End Set
	End Property
	'''<remarks/>
	Public Property TC201() As String
		Get
			Return Me.tC201Field
		End Get
		Set
			Me.tC201Field = Value
		End Set
	End Property
End Class




