Friend NotInheritable Class NullDbLogger
    Implements IDbLogger

    Public Shared ReadOnly Property Instance As New NullDbLogger()

    Private Sub New()
    End Sub

    Public Sub LogInfo(message As String) Implements IDbLogger.LogInfo
    End Sub

    Public Sub LogWarning(message As String) Implements IDbLogger.LogWarning
    End Sub

    Public Sub LogError(message As String, Optional ex As Exception = Nothing) Implements IDbLogger.LogError
    End Sub

    'Public Sub LogSql(sql As String) Implements IDbLogger.LogSql
    'End Sub
End Class
