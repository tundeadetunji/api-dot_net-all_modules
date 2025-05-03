Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Public Class SqlServerProvider
    Implements IDbProvider

    Private ReadOnly _connectionString As String

    Public Sub New(connectionString As String)
        _connectionString = connectionString
    End Sub

    Public Function CreateConnection() As IDbConnection Implements IDbProvider.CreateConnection
        Return New SqlConnection(_connectionString)
    End Function

    Public Function CreateCommand(query As String, connection As IDbConnection) As IDbCommand Implements IDbProvider.CreateCommand
        Return New SqlCommand(query, CType(connection, SqlConnection))
    End Function

    Public Function CreateParameter(name As String, value As Object) As IDataParameter Implements IDbProvider.CreateParameter
        Return New SqlParameter(name, value)
    End Function

    Public Function GetParameterPrefix() As String Implements IDbProvider.GetParameterPrefix
        Return "@"
    End Function

    Public Function CloneParameter(parameter As IDataParameter) As IDataParameter Implements IDbProvider.CloneParameter
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
        Return clone
    End Function
End Class
