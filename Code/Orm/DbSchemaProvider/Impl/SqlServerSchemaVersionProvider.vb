Friend Class SqlServerSchemaVersionProvider
    Implements ISchemaVersionProvider

    Private ReadOnly _logger As IDbLogger

    Public Sub New(logger As IDbLogger)
        _logger = logger
    End Sub

    Public Sub EnsureSchemaVersionTable(connection As IDbConnection, transaction As IDbTransaction) Implements ISchemaVersionProvider.EnsureSchemaVersionTable
        Dim sql = "
            IF OBJECT_ID('_SchemaVersion', 'U') IS NULL
            CREATE TABLE [_SchemaVersion] (
                [Version] INT NOT NULL,
                [AppliedAt] DATETIME NOT NULL,
                [Notes] NVARCHAR(500),
                CONSTRAINT PK_SchemaVersion PRIMARY KEY ([Version])
            );"
        ExecuteNonQuery(sql, connection, transaction)
        _logger?.LogInfo("Ensured _SchemaVersion table exists.")
    End Sub

    Public Function GetCurrentSchemaVersion(connection As IDbConnection, transaction As IDbTransaction) As Integer Implements ISchemaVersionProvider.GetCurrentSchemaVersion
        Dim sql = "SELECT MAX([Version]) FROM [_SchemaVersion];"
        Using cmd = connection.CreateCommand()
            cmd.Transaction = transaction
            cmd.CommandText = sql
            Dim result = cmd.ExecuteScalar()
            If result Is DBNull.Value OrElse result Is Nothing Then Return 0
            Return Convert.ToInt32(result)
        End Using
    End Function

    Public Sub InsertSchemaVersion(connection As IDbConnection, transaction As IDbTransaction, version As Integer, notes As String) Implements ISchemaVersionProvider.InsertSchemaVersion
        Dim sql = "
            INSERT INTO [_SchemaVersion] ([Version], [AppliedAt], [Notes]) 
            VALUES (@version, @appliedAt, @notes);"
        Using cmd = connection.CreateCommand()
            cmd.Transaction = transaction
            cmd.CommandText = sql

            Dim p1 = cmd.CreateParameter() : p1.ParameterName = "@version" : p1.Value = version : cmd.Parameters.Add(p1)
            Dim p2 = cmd.CreateParameter() : p2.ParameterName = "@appliedAt" : p2.Value = DateTime.UtcNow : cmd.Parameters.Add(p2)
            Dim p3 = cmd.CreateParameter() : p3.ParameterName = "@notes" : p3.Value = notes : cmd.Parameters.Add(p3)

            cmd.ExecuteNonQuery()
        End Using
        _logger?.LogInfo($"Inserted schema version {version}.")
    End Sub

    Private Sub ExecuteNonQuery(sql As String, connection As IDbConnection, transaction As IDbTransaction)
        Using cmd = connection.CreateCommand()
            cmd.Transaction = transaction
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        End Using
    End Sub
End Class
