Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T10F2U2_StocktakingExecute
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody

        <XmlElement(ElementName:="StocktakingList")>
        Public Property StocktakingList As New clsStocktakingList

        Public Class clsStocktakingList
            <XmlElement(ElementName:="StocktakingInfo")>
            Public Property StocktakingInfo As New List(Of clsStocktakingInfo)
            Public Class clsStocktakingInfo
        Public Property STOCKTAKING_ID As String
      End Class
    End Class
  End Class
End Class