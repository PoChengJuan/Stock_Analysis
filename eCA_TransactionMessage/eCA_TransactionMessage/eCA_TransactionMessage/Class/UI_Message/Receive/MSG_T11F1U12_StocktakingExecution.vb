'執行上位系統盤點單
Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_T11F1U12_StocktakingExecution
  Public Property Header As clsHeader
  Public Property Body As clsBody
  Public Property KeepData As String

  Public Class clsBody
    Public Property Action As String
    <XmlElement(ElementName:="StocktakingList")>
    Public Property StocktakingList As New clsStocktakingList

    Public Class clsStocktakingList
      <XmlElement(ElementName:="StocktakingInfo")>
      Public Property StocktakingInfo As New List(Of clsStocktakingInfo)

      Public Class clsStocktakingInfo
        Public Property STOCKTAKING_ID As String
        Public Property STOCKTAKING_TYPE1 As String
        Public Property STOCKTAKING_TYPE2 As String
        Public Property STOCKTAKING_TYPE3 As String
      End Class
    End Class
  End Class
End Class

