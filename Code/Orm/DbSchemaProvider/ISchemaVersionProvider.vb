Public Interface ISchemaVersionProvider
    Sub EnsureSchemaVersionTable(connection As IDbConnection, transaction As IDbTransaction)
    Function GetCurrentSchemaVersion(connection As IDbConnection, transaction As IDbTransaction) As Integer
    Sub InsertSchemaVersion(connection As IDbConnection, transaction As IDbTransaction, version As Integer, notes As String)
End Interface
