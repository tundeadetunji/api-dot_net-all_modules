Imports System.Data.SqlClient
Public Class SqlServerProviderAsync
    Implements IDbProviderAsync

    Private ReadOnly _connectionString As String

    Public Sub New(connectionString As String)
        _connectionString = connectionString
    End Sub

    Public Async Function CreateConnectionAsync() As Task(Of IDbConnection) Implements IDbProviderAsync.CreateConnectionAsync
        Return Await Task.FromResult(New SqlConnection(_connectionString))
    End Function

    Public Function CreateCommandAsync(query As String, connection As SqlConnection) As Task(Of SqlCommand) Implements IDbProviderAsync.CreateCommandAsync
        Dim command As New SqlCommand(query, connection)
        Return Task.FromResult(command)
    End Function

    Public Function CreateParameterAsync(name As String, value As Object) As Task(Of SqlParameter) Implements IDbProviderAsync.CreateParameterAsync
        Dim param As New SqlParameter(name, value)
        Return Task.FromResult(param)
    End Function

    Public Function CloneParameterAsync(parameter As SqlParameter) As Task(Of SqlParameter) Implements IDbProviderAsync.CloneParameterAsync
        Dim original = CType(parameter, SqlParameter)
        Dim clone As New SqlParameter With {
            .ParameterName = original.ParameterName,
            .Value = original.Value,
            .SqlDbType = original.SqlDbType,
            .Size = original.Size,
            .Direction = original.Direction,
            .IsNullable = original.IsNullable,
            .Precision = original.Precision,
            .Scale = original.Scale,
            .SourceColumn = original.SourceColumn
        }
        Return Task.FromResult(clone)
    End Function

    Public Function GetParameterPrefix() As String Implements IDbProviderAsync.GetParameterPrefix
        Return "@"
    End Function
End Class
