Public Class User
    Public Property title As String
    Public Property firstName As String
    Public Property lastName As String
    Public Property email As String
    Public Property phone As String
    Public Property password As String
    Public Property enabled As Boolean
    Public Property registered As Boolean

    Public Overrides Function ToString() As String
        Return title & " " & firstName & " " & lastName
    End Function
End Class
