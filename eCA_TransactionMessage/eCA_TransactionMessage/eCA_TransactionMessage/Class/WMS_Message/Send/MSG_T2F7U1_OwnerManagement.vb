Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T2F7U1_OwnerManagement
    Public Property Header As New clsHeader
    Public Property Body As New SKUDataList
    Public Property KeepData As String

    Public Class SKUDataList

        Public Property Action As String

        Public Property OwnerList As New clsOwnerList

        Public Class clsOwnerList
            <XmlElement(ElementName:="OwnerInfo")>
            Public Property OwnerInfo As New List(Of clsOwnerInfo)

            Public Class clsOwnerInfo
                Public Property OWNER_NO As String
                Public Property OWNER_ID1 As String
                Public Property OWNER_ID2 As String
                Public Property OWNER_ID3 As String
                Public Property OWNER_ALIS1 As String
                Public Property OWNER_ALIS2 As String
                Public Property OWNER_DESC As String
                Public Property OWNER_TYPE As String
                Public Property OWNER_COMMON1 As String
                Public Property OWNER_COMMON2 As String
                Public Property OWNER_COMMON3 As String
                Public Property OWNER_COMMON4 As String
                Public Property OWNER_COMMON5 As String
                Public Property OWNER_COMMON6 As String
                Public Property OWNER_COMMON7 As String
                Public Property OWNER_COMMON8 As String
                Public Property OWNER_COMMON9 As String
                Public Property OWNER_COMMON10 As String
                Public Property ENABLE As String
                Public Property COMMENTS As String
            End Class
        End Class
    End Class
End Class