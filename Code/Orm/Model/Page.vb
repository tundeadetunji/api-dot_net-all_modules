Public Class Page(Of T)
    Public Sub New(records As List(Of T), currentPage As Integer, recordsPerPage As Integer, recordCount As Long, pageCount As Integer)
        Me.Records = records
        Me.CurrentPage = currentPage
        Me.RecordsPerPage = recordsPerPage
        Me.RecordCount = recordCount
        Me.PageCount = pageCount
    End Sub

    Public ReadOnly Property Records As List(Of T)
    Public ReadOnly Property CurrentPage As Integer
    Public ReadOnly Property RecordsPerPage As Integer
    Public ReadOnly Property RecordCount As Long
    Public ReadOnly Property PageCount As Integer

    Public ReadOnly Property IsFirstPage As Boolean
        Get
            Return CurrentPage = 1
        End Get
    End Property

    Public ReadOnly Property IsLastPage As Boolean
        Get
            Return CurrentPage >= PageCount
        End Get
    End Property

    Public ReadOnly Property HasPreviousPage As Boolean
        Get
            Return CurrentPage > 1
        End Get
    End Property

    Public ReadOnly Property HasNextPage As Boolean
        Get
            Return CurrentPage < PageCount
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return "Current Page: " & CurrentPage & vbCrLf &
            "Records Per Page: " & RecordsPerPage & vbCrLf &
            "Record Count: " & RecordCount & vbCrLf &
            "Page Count: " & PageCount & vbCrLf &
            "Is First Page: " & IsFirstPage & vbCrLf &
            "Is Last Page: " & IsLastPage & vbCrLf &
            "Has Previous Page: " & HasPreviousPage &
            "Has Next Page: " & HasNextPage
    End Function
End Class