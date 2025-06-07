''' <summary>
''' Allows logging of SQL operations and errors in a pluggable way.
''' </summary>
Public Interface IDbLogger
    Sub LogInfo(message As String)
    Sub LogWarning(message As String)
    Sub LogError(message As String, Optional ex As Exception = Nothing)
End Interface
