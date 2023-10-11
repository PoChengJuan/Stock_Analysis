'<INPUT>
'<MATNR>50-201083-01</MATNR>
'<WERKS>1710</WERKS>
'<CHARG>0103032668</CHARG>
'<LGORT>8500</LGORT>
'</INPUT>

'''<remarks/>
'''環鴻與ERP對帳輸入的XML格式
<System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="", IsNullable:=False, ElementName:="INPUT")>
Partial Public Class ASRS_updateBin

    Private mATNRField As String

    Private wERKSField As String

    Private cHARGField As String

    Private lGORTField As String

    '''<remarks/>
    Public Property MATNR() As String
        Get
            Return Me.mATNRField
        End Get
        Set
            Me.mATNRField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property WERKS() As String
        Get
            Return Me.wERKSField
        End Get
        Set
            Me.wERKSField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property CHARG() As String
        Get
            Return Me.cHARGField
        End Get
        Set
            Me.cHARGField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property LGORT() As String
        Get
            Return Me.lGORTField
        End Get
        Set
            Me.lGORTField = Value
        End Set
    End Property
End Class

