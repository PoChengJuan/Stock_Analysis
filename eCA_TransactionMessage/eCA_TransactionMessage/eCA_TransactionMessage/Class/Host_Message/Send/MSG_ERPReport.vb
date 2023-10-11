'<STD_IN>
'<ProdID> WMS</ProdID>
'  <Companyid> 元翎精密</Companyid>
'  <Userid> DS</Userid>
'  <DoAction>1</DoAction>
'  <Docase>1</Docase>
'  <Result> success</Result>
'  <Data>
'    <RecordList>
'    <TD001></TD001> '單別
'	 <TD002></TD002>	'單號
'    </RecordList>
'  </Data>
'</STD_IN>


'''<remarks/>
<System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False, ElementName:="STD_IN")>
Partial Public Class MSG_ERPReport

  Private prodIDField As String

  Private companyidField As String

  Private useridField As String

  Private doActionField As String

  Private docaseField As String

  Private dataField As New STD_INData

  Private resultField As String

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
  Public Property Docase() As String
    Get
      Return Me.docaseField
    End Get
    Set
      Me.docaseField = Value
    End Set
  End Property

  '''<remarks/>
  Public Property Result() As String
    Get
      Return Me.resultField
    End Get
    Set
      Me.resultField = Value
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
<System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class MSG_ERPReportData

  Private recordListField As STD_INDataRecordList

  '''<remarks/>
  Public Property RecordList() As STD_INDataRecordList
    Get
      Return Me.recordListField
    End Get
    Set
      Me.recordListField = Value
    End Set
  End Property
End Class

'''<remarks/>
<System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True)>
Partial Public Class STD_INDataRecordList

  Private tD001Field As String

  Private tD002Field As String

  Private textField() As String

  '''<remarks/>
  Public Property TD001() As String
    Get
      Return Me.tD001Field
    End Get
    Set
      Me.tD001Field = Value
    End Set
  End Property

  '''<remarks/>
  Public Property TD002() As String
    Get
      Return Me.tD002Field
    End Get
    Set
      Me.tD002Field = Value
    End Set
  End Property

  '''<remarks/>
  <System.Xml.Serialization.XmlTextAttribute()>
  Public Property Text() As String()
    Get
      Return Me.textField
    End Get
    Set
      Me.textField = Value
    End Set
  End Property
End Class

