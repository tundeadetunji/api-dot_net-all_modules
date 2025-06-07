Friend NotInheritable Class ConsoleDbLogger
    Implements IDbLogger
    Public Shared ReadOnly Property Instance As New ConsoleDbLogger()

    Private Sub New()
    End Sub
    Public Sub LogInfo(message As String) Implements IDbLogger.LogInfo
        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine($"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}")
        Console.ResetColor()
    End Sub

    Public Sub LogWarning(message As String) Implements IDbLogger.LogWarning
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine($"[WARN] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}")
        Console.ResetColor()
    End Sub

    Public Sub LogError(message As String, Optional ex As Exception = Nothing) Implements IDbLogger.LogError
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine($"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}")
        If ex IsNot Nothing Then
            Console.WriteLine($"       {ex.GetType().Name}: {ex.Message}")
            Console.WriteLine(ex.StackTrace)
        End If
        Console.ResetColor()
    End Sub
End Class
