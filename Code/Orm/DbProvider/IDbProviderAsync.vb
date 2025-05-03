Imports System.Data.SqlClient

Public Interface IDbProviderAsync
    Function GetParameterPrefix() As String
    Function CreateConnectionAsync() As Task(Of IDbConnection)
    Function CreateCommandAsync(query As String, connection As SqlConnection) As Task(Of SqlCommand)
    Function CreateParameterAsync(name As String, value As Object) As Task(Of SqlParameter)
    Function CloneParameterAsync(parameter As SqlParameter) As Task(Of SqlParameter)
End Interface
