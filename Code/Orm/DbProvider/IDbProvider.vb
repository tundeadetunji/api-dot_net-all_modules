Public Interface IDbProvider
    Function GetParameterPrefix() As String
    Function CreateConnection() As IDbConnection
    Function CreateCommand(query As String, connection As IDbConnection) As IDbCommand
    Function CreateParameter(name As String, value As Object) As IDataParameter
    Function CloneParameter(parameter As IDataParameter) As IDataParameter
End Interface
