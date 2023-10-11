Public Class CSVOption
  Public Property FilePath As String = ""
  Public Property SQList As New List(Of SQLInfo)

  Public Class SQLInfo
    Public Property FileName As String = ""
    Public Property SQL As String = ""
    Public Property DateTimeRange As New DateTimeRangeInfo
  End Class

  Public Class DateTimeRangeInfo
    Public Property FieldName As String = ""
    Public Property HourPeriod As Integer = 1
  End Class

End Class
