Imports iNovation.Code.OrmUtils
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Threading

Friend Class SqlServerOrmAsync
    Implements IOrmAsync

#Region "init"
    Private ReadOnly _provider As IDbProviderAsync
    Public Const Id As String = "Id"

    Private Shared _instance As Lazy(Of SqlServerOrmAsync)

    Private Sub New(connectionString As String)
        _provider = New SqlServerProviderAsync(connectionString)
    End Sub

    Public Shared ReadOnly Property Instance(connectionString As String) As SqlServerOrmAsync
        Get
            If _instance Is Nothing Then
                _instance = New Lazy(Of SqlServerOrmAsync)(Function() New SqlServerOrmAsync(connectionString), LazyThreadSafetyMode.ExecutionAndPublication)
            End If
            Return _instance.Value
        End Get
    End Property
#End Region

#Region "exported"
    Public Async Sub PrepareDatabaseAsync(mode As DbPrepMode, entities As List(Of Type), Optional idColumn As String = Id) Implements IOrmAsync.PrepareDatabaseAsync
        Using connection = CType(Await _provider.CreateConnectionAsync(), SqlConnection)
            Await connection.OpenAsync()
            Using transaction = connection.BeginTransaction()

                Dim executedTables As New HashSet(Of String)

                For Each entityType In entities
                    Await CreateOrUpdateTableRecursiveAsync(entityType, idColumn, mode, connection, transaction, executedTables)
                Next

                transaction.Commit()
            End Using
        End Using
    End Sub

    Public Async Sub DeleteAllAsync(Of T)(Optional cascade As Boolean = True) Implements IOrmAsync.DeleteAllAsync
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using transaction = connection.BeginTransaction()
                Try
                    ' Check for cascade delete constraint
                    If Not cascade AndAlso Await HasCascadeDeleteConstraintAsync(tableName, connection, transaction, _provider) Then
                        Throw New InvalidOperationException($"DeleteAll on '{tableName}' was called with cascade:=False, but the database has foreign key(s) configured with ON DELETE CASCADE.")
                    End If

                    If cascade Then
                        ' Delete from child tables first
                        Dim listProps = typeT.GetProperties().
                        Where(Function(p) IsGenericList(p.PropertyType) AndAlso p.CanRead).ToList()

                        For Each listProp In listProps
                            Dim childType = listProp.PropertyType.GetGenericArguments()(0)
                            Dim childTable = childType.Name

                            Dim childSql = $"DELETE FROM [{childTable}]"
                            Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                                childCmd.Transaction = transaction
                                Await CType(childCmd, DbCommand).ExecuteNonQueryAsync()
                            End Using
                        Next
                    End If

                    Try
                        Dim deleteSql = $"DELETE FROM [{tableName}]"
                        Using fallbackCmd = Await _provider.CreateCommandAsync(deleteSql, connection)
                            fallbackCmd.Transaction = transaction
                            Await CType(fallbackCmd, DbCommand).ExecuteNonQueryAsync()
                        End Using
                    Catch ex As Exception
                        ' Optional: Log ex if needed
                    End Try

                    transaction.Commit() ' still synchronous
                Catch
                    transaction.Rollback() ' still synchronous
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Async Sub DeleteAllInTableAsync(Of T)(tableName As String, Optional cascade As Boolean = True) Implements IOrmAsync.DeleteAllInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using transaction = connection.BeginTransaction()
                Try
                    ' Enforce decision if DB is configured to cascade deletes
                    If Not cascade AndAlso Await HasCascadeDeleteConstraintAsync(tableName, connection, transaction, _provider) Then
                        Throw New InvalidOperationException($"DeleteAll on '{tableName}' was called with cascade:=False, but the database has foreign key(s) configured with ON DELETE CASCADE.")
                    End If

                    If cascade Then
                        ' Delete from all child tables first
                        Dim listProps = typeT.GetProperties().
                        Where(Function(p) IsGenericList(p.PropertyType) AndAlso p.CanRead).ToList()

                        For Each listProp In listProps
                            Dim childType = listProp.PropertyType.GetGenericArguments()(0)
                            Dim childTable = childType.Name
                            Dim fkColumn = $"{tableName}_Id"

                            Dim childSql = $"DELETE FROM [{childTable}]"
                            Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                                childCmd.Transaction = transaction
                                Await CType(childCmd, DbCommand).ExecuteNonQueryAsync()
                            End Using
                        Next
                    End If

                    Try
                        Dim deleteSql = $"DELETE FROM [{tableName}]"
                        Using fallbackCmd = Await _provider.CreateCommandAsync(deleteSql, connection)
                            fallbackCmd.Transaction = transaction
                            Await CType(fallbackCmd, DbCommand).ExecuteNonQueryAsync()
                        End Using
                    Catch ex As Exception
                        ' Optional: log the exception if you want
                    End Try

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Async Sub DeleteWhereIdAsync(Of T)(id As Object, Optional idColumn As String = "Id", Optional cascade As Boolean = True) Implements IOrmAsync.DeleteWhereIdAsync
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()
        Dim idProp = props.FirstOrDefault(Function(p) p.Name = idColumn)
        If idProp Is Nothing Then
            Throw New InvalidOperationException($"The specified id column '{idColumn}' was not found on type '{typeT.Name}'.")
        End If

        ' Detect RowVersion column
        Dim rowVersionProp = props.FirstOrDefault(Function(p) String.Equals(p.Name, "RowVersion", StringComparison.OrdinalIgnoreCase) AndAlso p.PropertyType Is GetType(Byte()))

        ' Load the current RowVersion value (if present)
        Dim rowVersionValue As Byte() = Nothing
        If rowVersionProp IsNot Nothing Then
            Dim findSql = $"SELECT [RowVersion] FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"
            Using connection = Await _provider.CreateConnectionAsync()
                Await CType(connection, DbConnection).OpenAsync()
                Using cmd = Await _provider.CreateCommandAsync(findSql, connection)
                    cmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", id))
                    Dim result = Await CType(cmd, DbCommand).ExecuteScalarAsync()
                    If result Is DBNull.Value OrElse result Is Nothing Then
                        Throw New InvalidOperationException("Record not found or missing RowVersion.")
                    End If
                    rowVersionValue = CType(result, Byte())
                End Using
            End Using
        End If

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using transaction = connection.BeginTransaction()
                Try
                    ' Cascade delete children first if needed
                    If cascade Then
                        For Each prop In props.Where(Function(p) IsGenericList(p.PropertyType))
                            Dim childType = prop.PropertyType.GetGenericArguments()(0)
                            Dim childTable = childType.Name
                            Dim fkColumn = $"{tableName}_{idColumn}"
                            Dim deleteChildSql = $"DELETE FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}id"

                            Using childCmd = Await _provider.CreateCommandAsync(deleteChildSql, connection)
                                childCmd.Transaction = transaction
                                childCmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", id))
                                Await CType(childCmd, DbCommand).ExecuteNonQueryAsync()
                            End Using
                        Next
                    End If

                    ' Delete parent record (with optional RowVersion check)
                    Dim deleteSql = $"DELETE FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"
                    If rowVersionProp IsNot Nothing Then
                        deleteSql &= $" AND [RowVersion] = {parameterPrefix}rowversion"
                    End If

                    Using cmd = Await _provider.CreateCommandAsync(deleteSql, connection)
                        cmd.Transaction = transaction
                        cmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", id))

                        If rowVersionProp IsNot Nothing Then
                            Dim param = Await _provider.CreateParameterAsync($"{parameterPrefix}rowversion", rowVersionValue)
                            param.DbType = DbType.Binary
                            cmd.Parameters.Add(param)
                        End If

                        Dim affected = Await CType(cmd, DbCommand).ExecuteNonQueryAsync()
                        If affected <> 1 Then
                            Throw New DBConcurrencyException("Delete failed due to concurrent update. Record was modified or deleted by another process.")
                        End If
                    End Using

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Async Sub DeleteWhereIdInTableAsync(Of T)(id As Object, tableName As String, Optional idColumn As String = "Id", Optional cascade As Boolean = True) Implements IOrmAsync.DeleteWhereIdInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()
        Dim idProp = props.FirstOrDefault(Function(p) p.Name = idColumn)
        If idProp Is Nothing Then
            Throw New InvalidOperationException($"The specified id column '{idColumn}' was not found on type '{typeT.Name}'.")
        End If

        ' Detect RowVersion column
        Dim rowVersionProp = props.FirstOrDefault(Function(p) String.Equals(p.Name, "RowVersion", StringComparison.OrdinalIgnoreCase) AndAlso p.PropertyType Is GetType(Byte()))

        ' Load the current RowVersion value (if present)
        Dim rowVersionValue As Byte() = Nothing
        If rowVersionProp IsNot Nothing Then
            Dim findSql = $"SELECT [RowVersion] FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"
            Using connection = Await _provider.CreateConnectionAsync()
                Await CType(connection, DbConnection).OpenAsync()
                Using cmd = Await _provider.CreateCommandAsync(findSql, connection)
                    cmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", id))
                    Dim result = Await CType(cmd, DbCommand).ExecuteScalarAsync()
                    If result Is DBNull.Value OrElse result Is Nothing Then
                        Throw New InvalidOperationException("Record not found or missing RowVersion.")
                    End If
                    rowVersionValue = CType(result, Byte())
                End Using
            End Using
        End If

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using transaction = connection.BeginTransaction()
                Try
                    ' Cascade delete children first if needed
                    If cascade Then
                        For Each prop In props.Where(Function(p) IsGenericList(p.PropertyType))
                            Dim childType = prop.PropertyType.GetGenericArguments()(0)
                            Dim childTable = childType.Name
                            Dim fkColumn = $"{tableName}_{idColumn}"
                            Dim deleteChildSql = $"DELETE FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}id"

                            Using childCmd = Await _provider.CreateCommandAsync(deleteChildSql, connection)
                                childCmd.Transaction = transaction
                                childCmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", id))
                                Await CType(childCmd, DbCommand).ExecuteNonQueryAsync()
                            End Using
                        Next
                    End If

                    ' Delete parent record (with optional RowVersion check)
                    Dim deleteSql = $"DELETE FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"
                    If rowVersionProp IsNot Nothing Then
                        deleteSql &= $" AND [RowVersion] = {parameterPrefix}rowversion"
                    End If

                    Using cmd = Await _provider.CreateCommandAsync(deleteSql, connection)
                        cmd.Transaction = transaction
                        cmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", id))

                        If rowVersionProp IsNot Nothing Then
                            Dim param = Await _provider.CreateParameterAsync($"{parameterPrefix}rowversion", rowVersionValue)
                            param.DbType = DbType.Binary
                            cmd.Parameters.Add(param)
                        End If

                        Dim affected = Await CType(cmd, DbCommand).ExecuteNonQueryAsync()
                        If affected <> 1 Then
                            Throw New DBConcurrencyException("Delete failed due to concurrent update. Record was modified or deleted by another process.")
                        End If
                    End Using

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Async Sub DeleteByAsync(Of T)(conditions As List(Of Condition), Optional cascade As Boolean = True) Implements IOrmAsync.DeleteByAsync

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        ' Build WHERE clauses and parameters
        For i = 0 To conditions.Count - 1
            Dim cond = conditions(i)
            Dim paramName = $"{parameterPrefix}p{i}"
            Dim sqlOp = GetSqlOperator(cond.SqlComparison)
            whereClauses.Add($"[{cond.Column}] {sqlOp} {paramName}")
            parameters.Add(Await _provider.CreateParameterAsync(paramName, cond.Value))
        Next

        Dim whereSql = If(whereClauses.Any(), " WHERE " & String.Join(" AND ", whereClauses), "")
        If String.IsNullOrWhiteSpace(whereSql) Then
            Throw New InvalidOperationException("DeleteByAsync must have at least one condition to avoid deleting all records.")
        End If

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using transaction = connection.BeginTransaction()

                Try
                    ' Check for unsafe DB-side cascade
                    If Not cascade AndAlso Await HasCascadeDeleteConstraintAsync(tableName, connection, transaction, _provider) Then
                        Throw New InvalidOperationException(
                        $"The database has ON DELETE CASCADE constraint(s) on table '{tableName}', but you passed cascade:=False. This would silently delete child records."
                    )
                    End If

                    ' Manual child delete if cascade = True
                    If cascade Then
                        Dim listProps = typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType) AndAlso p.CanRead).ToList()
                        For Each listProp In listProps
                            Dim childType = listProp.PropertyType.GetGenericArguments()(0)
                            Dim childTable = childType.Name
                            Dim fkColumn = $"{tableName}_Id"

                            ' Build child delete with same WHERE conditions
                            Dim childWhereSql = $" WHERE [{fkColumn}] IN (
                            SELECT [{tableName}].[Id] FROM [{tableName}]{whereSql}
                        )"
                            Dim childSql = $"DELETE FROM [{childTable}]{childWhereSql}"

                            Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                                childCmd.Transaction = transaction
                                For Each param In parameters
                                    childCmd.Parameters.Add(Await _provider.CloneParameterAsync(param))
                                Next
                                Await CType(childCmd, DbCommand).ExecuteNonQueryAsync()
                            End Using
                        Next
                    End If

                    ' Delete parent
                    Dim parentSql = $"DELETE FROM [{tableName}]{whereSql}"
                    Using cmd = Await _provider.CreateCommandAsync(parentSql, connection)
                        cmd.Transaction = transaction
                        For Each param In parameters
                            cmd.Parameters.Add(Await _provider.CloneParameterAsync(param))
                        Next
                        Await CType(cmd, DbCommand).ExecuteNonQueryAsync()
                    End Using

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Async Sub DeleteByInTableAsync(Of T)(conditions As List(Of Condition), tableName As String, Optional cascade As Boolean = True) Implements IOrmAsync.DeleteByInTableAsync

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        ' Build WHERE clauses and parameters
        For i = 0 To conditions.Count - 1
            Dim cond = conditions(i)
            Dim paramName = $"{parameterPrefix}p{i}"
            Dim sqlOp = GetSqlOperator(cond.SqlComparison)
            whereClauses.Add($"[{cond.Column}] {sqlOp} {paramName}")
            parameters.Add(Await _provider.CreateParameterAsync(paramName, cond.Value))
        Next

        Dim whereSql = If(whereClauses.Any(), " WHERE " & String.Join(" AND ", whereClauses), "")
        If String.IsNullOrWhiteSpace(whereSql) Then
            Throw New InvalidOperationException("DeleteByAsync must have at least one condition to avoid deleting all records.")
        End If

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using transaction = connection.BeginTransaction()

                Try
                    ' Check for unsafe DB-side cascade
                    If Not cascade AndAlso Await HasCascadeDeleteConstraintAsync(tableName, connection, transaction, _provider) Then
                        Throw New InvalidOperationException(
                        $"The database has ON DELETE CASCADE constraint(s) on table '{tableName}', but you passed cascade:=False. This would silently delete child records."
                    )
                    End If

                    ' Manual child delete if cascade = True
                    If cascade Then
                        Dim listProps = typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType) AndAlso p.CanRead).ToList()
                        For Each listProp In listProps
                            Dim childType = listProp.PropertyType.GetGenericArguments()(0)
                            Dim childTable = childType.Name
                            Dim fkColumn = $"{tableName}_Id"

                            ' Build child delete with same WHERE conditions
                            Dim childWhereSql = $" WHERE [{fkColumn}] IN (
                            SELECT [{tableName}].[Id] FROM [{tableName}]{whereSql}
                        )"
                            Dim childSql = $"DELETE FROM [{childTable}]{childWhereSql}"

                            Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                                childCmd.Transaction = transaction
                                For Each param In parameters
                                    childCmd.Parameters.Add(Await _provider.CloneParameterAsync(param))
                                Next
                                Await CType(childCmd, DbCommand).ExecuteNonQueryAsync()
                            End Using
                        Next
                    End If

                    ' Delete parent
                    Dim parentSql = $"DELETE FROM [{tableName}]{whereSql}"
                    Using cmd = Await _provider.CreateCommandAsync(parentSql, connection)
                        cmd.Transaction = transaction
                        For Each param In parameters
                            cmd.Parameters.Add(Await _provider.CloneParameterAsync(param))
                        Next
                        Await CType(cmd, DbCommand).ExecuteNonQueryAsync()
                    End Using

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Async Function CountByAsync(Of T)(conditions As List(Of Condition)) As Task(Of Long) Implements IOrmAsync.CountByAsync

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        Dim tableName As String = GetType(T).Name
        Dim paramPrefix As String = _provider.GetParameterPrefix()
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        ' Build WHERE clauses and parameters
        For i = 0 To conditions.Count - 1
            Dim condition = conditions(i)
            Dim paramName = $"{paramPrefix}p{i}"
            Dim sqlOp = GetSqlOperator(condition.SqlComparison)
            whereClauses.Add($"{condition.Column} {sqlOp} {paramName}")
            parameters.Add(Await _provider.CreateParameterAsync(paramName, condition.Value))
        Next

        Dim whereSql = If(whereClauses.Any(), " WHERE " & String.Join(" AND ", whereClauses), "")
        Dim query = $"SELECT COUNT(1) FROM {tableName}{whereSql}"

        Using conn = Await _provider.CreateConnectionAsync()
            Await CType(conn, DbConnection).OpenAsync()
            Using cmd = Await _provider.CreateCommandAsync(query, conn)
                For Each param In parameters
                    cmd.Parameters.Add(Await _provider.CloneParameterAsync(param))
                Next

                Dim result = Await CType(cmd, DbCommand).ExecuteScalarAsync()
                Return Convert.ToInt64(result)
            End Using
        End Using
    End Function
    Public Async Function CountByInTableAsync(Of T)(conditions As List(Of Condition), tableName As String) As Task(Of Long) Implements IOrmAsync.CountByInTableAsync

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim paramPrefix As String = _provider.GetParameterPrefix()
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        ' Build WHERE clauses and parameters
        For i = 0 To conditions.Count - 1
            Dim condition = conditions(i)
            Dim paramName = $"{paramPrefix}p{i}"
            Dim sqlOp = GetSqlOperator(condition.SqlComparison)
            whereClauses.Add($"{condition.Column} {sqlOp} {paramName}")
            parameters.Add(Await _provider.CreateParameterAsync(paramName, condition.Value))
        Next

        Dim whereSql = If(whereClauses.Any(), " WHERE " & String.Join(" AND ", whereClauses), "")
        Dim query = $"SELECT COUNT(1) FROM {tableName}{whereSql}"

        Using conn = Await _provider.CreateConnectionAsync()
            Await CType(conn, DbConnection).OpenAsync()
            Using cmd = Await _provider.CreateCommandAsync(query, conn)
                For Each param In parameters
                    cmd.Parameters.Add(Await _provider.CloneParameterAsync(param))
                Next

                Dim result = Await CType(cmd, DbCommand).ExecuteScalarAsync()
                Return Convert.ToInt64(result)
            End Using
        End Using
    End Function

    Public Async Function CountAsync(Of T)() As Task(Of Long) Implements IOrmAsync.CountAsync
        Dim tableName As String = GetType(T).Name
        Dim query As String = $"SELECT COUNT(*) FROM [{tableName}]"

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using command = Await _provider.CreateCommandAsync(query, connection)
                Dim result = Await CType(command, DbCommand).ExecuteScalarAsync()
                Return Convert.ToInt64(result)
            End Using
        End Using
    End Function
    Public Async Function CountInTableAsync(Of T)(tableName As String) As Task(Of Long) Implements IOrmAsync.CountInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim query As String = $"SELECT COUNT(*) FROM [{tableName}]"

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using command = Await _provider.CreateCommandAsync(query, connection)
                Dim result = Await CType(command, DbCommand).ExecuteScalarAsync()
                Return Convert.ToInt64(result)
            End Using
        End Using
    End Function

    Public Async Function CreateAsync(Of T)(objs As List(Of T),
                                        Optional idColumn As String = "Id",
                                        Optional IdWillAutoIncrement As Boolean = True) As Task(Of List(Of T)) Implements IOrmAsync.CreateAsync
        If objs Is Nothing OrElse objs.Count = 0 Then Return objs

        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        Using connection = CType(Await _provider.CreateConnectionAsync(), SqlConnection)
            Await connection.OpenAsync()

            Using transaction = connection.BeginTransaction()
                For Each obj In objs
                    ' === Parent Insert ===
                    Dim insertCols = New List(Of String)
                    Dim insertVals = New List(Of String)
                    Dim parameters = New List(Of IDataParameter)

                    For Each prop In props
                        If IsGenericList(prop.PropertyType) Then Continue For
                        If prop.Name = idColumn AndAlso IdWillAutoIncrement Then Continue For

                        insertCols.Add($"[{prop.Name}]")
                        insertVals.Add($"{parameterPrefix}{prop.Name}")
                        Dim rawValue = prop.GetValue(obj)
                        Dim dbValue = GetEnumDbValue(prop, rawValue)
                        parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}{prop.Name}", dbValue))
                    Next

                    Dim insertSql = $"INSERT INTO [{tableName}] ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", insertVals)})"
                    If IdWillAutoIncrement Then insertSql &= "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"

                    Using cmd = Await _provider.CreateCommandAsync(insertSql, connection)
                        cmd.Transaction = transaction
                        For Each param In parameters
                            cmd.Parameters.Add(param)
                        Next

                        If IdWillAutoIncrement Then
                            Dim newId = Convert.ToInt64(cmd.ExecuteScalar()) ' sync
                            Dim idProp = props.First(Function(p) p.Name = idColumn)
                            idProp.SetValue(obj, Convert.ChangeType(newId, idProp.PropertyType))
                        Else
                            cmd.ExecuteNonQuery()
                        End If
                    End Using

                    ' === Child Insert ===
                    For Each prop In props.Where(Function(p) IsGenericList(p.PropertyType))
                        Dim childList = CType(prop.GetValue(obj), IEnumerable)
                        If childList Is Nothing Then Continue For

                        Dim childType = prop.PropertyType.GetGenericArguments()(0)
                        Dim childTable = childType.Name
                        Dim childProps = childType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite AndAlso Not IsGenericList(p.PropertyType)).ToList()

                        Dim parentId = props.First(Function(p) p.Name = idColumn).GetValue(obj)
                        Dim fkColumn = $"{tableName}_{idColumn}"

                        For Each child In childList
                            Dim cCols = New List(Of String)
                            Dim cVals = New List(Of String)
                            Dim cParams = New List(Of IDataParameter)

                            For Each cp In childProps
                                If cp.Name = idColumn AndAlso IdWillAutoIncrement Then Continue For

                                Dim value As Object
                                If cp.Name = fkColumn Then
                                    value = parentId
                                    If cp.CanWrite Then cp.SetValue(child, parentId)
                                Else
                                    value = cp.GetValue(child)
                                End If

                                Dim dbValue = GetEnumDbValue(cp, value)
                                cCols.Add($"[{cp.Name}]")
                                cVals.Add($"{parameterPrefix}{cp.Name}")
                                cParams.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}{cp.Name}", dbValue))
                            Next

                            Dim childInsertSql = $"INSERT INTO [{childTable}] ({String.Join(", ", cCols)}) VALUES ({String.Join(", ", cVals)})"
                            If IdWillAutoIncrement Then
                                childInsertSql &= "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                            End If

                            Using childCmd = Await _provider.CreateCommandAsync(childInsertSql, connection)
                                childCmd.Transaction = transaction
                                For Each p In cParams
                                    childCmd.Parameters.Add(p)
                                Next

                                If IdWillAutoIncrement Then
                                    Dim newChildId = Convert.ToInt64(childCmd.ExecuteScalar())
                                    Dim idProp = childProps.FirstOrDefault(Function(p) p.Name = idColumn)
                                    If idProp IsNot Nothing AndAlso idProp.CanWrite Then
                                        idProp.SetValue(child, Convert.ChangeType(newChildId, idProp.PropertyType))
                                    End If
                                Else
                                    childCmd.ExecuteNonQuery()
                                End If
                            End Using
                        Next
                    Next
                Next

                transaction.Commit() ' sync
            End Using
        End Using

        Return objs
    End Function

    Public Async Function CreateInTableAsync(Of T)(objs As List(Of T), tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of List(Of T)) Implements IOrmAsync.CreateInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        If objs Is Nothing OrElse objs.Count = 0 Then Return objs

        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        Using connection = CType(Await _provider.CreateConnectionAsync(), SqlConnection)
            Await connection.OpenAsync()

            Using transaction = connection.BeginTransaction()
                For Each obj In objs
                    ' === Parent Insert ===
                    Dim insertCols = New List(Of String)
                    Dim insertVals = New List(Of String)
                    Dim parameters = New List(Of IDataParameter)

                    For Each prop In props
                        If IsGenericList(prop.PropertyType) Then Continue For
                        If prop.Name = idColumn AndAlso IdWillAutoIncrement Then Continue For

                        insertCols.Add($"[{prop.Name}]")
                        insertVals.Add($"{parameterPrefix}{prop.Name}")
                        Dim rawValue = prop.GetValue(obj)
                        Dim dbValue = GetEnumDbValue(prop, rawValue)
                        parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}{prop.Name}", dbValue))
                    Next

                    Dim insertSql = $"INSERT INTO [{tableName}] ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", insertVals)})"
                    If IdWillAutoIncrement Then insertSql &= "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"

                    Using cmd = Await _provider.CreateCommandAsync(insertSql, connection)
                        cmd.Transaction = transaction
                        For Each param In parameters
                            cmd.Parameters.Add(param)
                        Next

                        If IdWillAutoIncrement Then
                            Dim newId = Convert.ToInt64(cmd.ExecuteScalar()) ' sync
                            Dim idProp = props.First(Function(p) p.Name = idColumn)
                            idProp.SetValue(obj, Convert.ChangeType(newId, idProp.PropertyType))
                        Else
                            cmd.ExecuteNonQuery()
                        End If
                    End Using

                    ' === Child Insert ===
                    For Each prop In props.Where(Function(p) IsGenericList(p.PropertyType))
                        Dim childList = CType(prop.GetValue(obj), IEnumerable)
                        If childList Is Nothing Then Continue For

                        Dim childType = prop.PropertyType.GetGenericArguments()(0)
                        Dim childTable = childType.Name
                        Dim childProps = childType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite AndAlso Not IsGenericList(p.PropertyType)).ToList()

                        Dim parentId = props.First(Function(p) p.Name = idColumn).GetValue(obj)
                        Dim fkColumn = $"{tableName}_{idColumn}"

                        For Each child In childList
                            Dim cCols = New List(Of String)
                            Dim cVals = New List(Of String)
                            Dim cParams = New List(Of IDataParameter)

                            For Each cp In childProps
                                If cp.Name = idColumn AndAlso IdWillAutoIncrement Then Continue For

                                Dim value As Object
                                If cp.Name = fkColumn Then
                                    value = parentId
                                    If cp.CanWrite Then cp.SetValue(child, parentId)
                                Else
                                    value = cp.GetValue(child)
                                End If

                                Dim dbValue = GetEnumDbValue(cp, value)
                                cCols.Add($"[{cp.Name}]")
                                cVals.Add($"{parameterPrefix}{cp.Name}")
                                cParams.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}{cp.Name}", dbValue))
                            Next

                            Dim childInsertSql = $"INSERT INTO [{childTable}] ({String.Join(", ", cCols)}) VALUES ({String.Join(", ", cVals)})"
                            If IdWillAutoIncrement Then
                                childInsertSql &= "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                            End If

                            Using childCmd = Await _provider.CreateCommandAsync(childInsertSql, connection)
                                childCmd.Transaction = transaction
                                For Each p In cParams
                                    childCmd.Parameters.Add(p)
                                Next

                                If IdWillAutoIncrement Then
                                    Dim newChildId = Convert.ToInt64(childCmd.ExecuteScalar())
                                    Dim idProp = childProps.FirstOrDefault(Function(p) p.Name = idColumn)
                                    If idProp IsNot Nothing AndAlso idProp.CanWrite Then
                                        idProp.SetValue(child, Convert.ChangeType(newChildId, idProp.PropertyType))
                                    End If
                                Else
                                    childCmd.ExecuteNonQuery()
                                End If
                            End Using
                        Next
                    Next
                Next

                transaction.Commit() ' sync
            End Using
        End Using

        Return objs
    End Function

    Public Async Function CreateAsync(Of T)(obj As T, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of T) Implements IOrmAsync.CreateAsync
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim properties = typeT.GetProperties().
        Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        Dim idProp = properties.FirstOrDefault(Function(p) p.Name = idColumn)
        If idProp Is Nothing Then
            Throw New InvalidOperationException($"The specified id column '{idColumn}' was not found on type '{typeT.Name}'.")
        End If

        Using connection = CType(Await _provider.CreateConnectionAsync(), SqlConnection)
            Await connection.OpenAsync()
            Using transaction = connection.BeginTransaction() ' synchronous

                ' Insert parent
                Dim insertCols = New List(Of String)
                Dim insertVals = New List(Of String)
                Dim parameters = New List(Of IDataParameter)

                For Each prop In properties
                    If IsGenericList(prop.PropertyType) Then Continue For
                    If prop.Name = idColumn AndAlso IdWillAutoIncrement Then Continue For

                    Dim paramName = $"{parameterPrefix}{prop.Name}"
                    insertCols.Add($"[{prop.Name}]")
                    insertVals.Add(paramName)
                    Dim rawValue = prop.GetValue(obj)
                    Dim dbValue = GetEnumDbValue(prop, rawValue)
                    parameters.Add(Await _provider.CreateParameterAsync(paramName, dbValue))
                Next

                Dim sql = $"INSERT INTO [{tableName}] ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", insertVals)})"
                If IdWillAutoIncrement Then
                    sql += "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                End If

                Using cmd = Await _provider.CreateCommandAsync(sql, connection)
                    cmd.Transaction = transaction
                    For Each param In parameters
                        cmd.Parameters.Add(param)
                    Next

                    If IdWillAutoIncrement Then
                        Dim newId = Convert.ToInt64(cmd.ExecuteScalar()) ' no async version in .NET Framework
                        idProp.SetValue(obj, Convert.ChangeType(newId, idProp.PropertyType))
                    Else
                        cmd.ExecuteNonQuery()
                    End If
                End Using

                ' Insert children
                For Each prop In properties.Where(Function(p) IsGenericList(p.PropertyType))
                    Dim childList = CType(prop.GetValue(obj), IEnumerable)
                    If childList Is Nothing Then Continue For

                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim childProps = childType.GetProperties().
                    Where(Function(p) p.CanRead AndAlso p.CanWrite AndAlso Not IsGenericList(p.PropertyType)).ToList()

                    Dim fkColumn = $"{tableName}_{idColumn}"

                    For Each child In childList
                        ' Set FK
                        Dim fkProp = childProps.FirstOrDefault(Function(p) p.Name = fkColumn)
                        If fkProp IsNot Nothing Then
                            fkProp.SetValue(child, idProp.GetValue(obj))
                        End If

                        Dim cCols = New List(Of String)
                        Dim cVals = New List(Of String)
                        Dim cParams = New List(Of IDataParameter)

                        For Each cp In childProps
                            If cp.Name = idColumn AndAlso IdWillAutoIncrement Then Continue For

                            Dim cpName = cp.Name
                            Dim rawValue = cp.GetValue(child)
                            Dim dbValue = GetEnumDbValue(cp, rawValue)

                            cCols.Add($"[{cpName}]")
                            cVals.Add($"{parameterPrefix}{cpName}")
                            cParams.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}{cpName}", dbValue))
                        Next

                        Dim childSql = $"INSERT INTO [{childTable}] ({String.Join(", ", cCols)}) VALUES ({String.Join(", ", cVals)}); SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"

                        Using cmd = Await _provider.CreateCommandAsync(childSql, connection)
                            cmd.Transaction = transaction
                            For Each p In cParams
                                cmd.Parameters.Add(p)
                            Next

                            Dim newChildId = Convert.ToInt64(cmd.ExecuteScalar()) ' sync
                            Dim childIdProp = childProps.FirstOrDefault(Function(p) p.Name = idColumn)
                            If childIdProp IsNot Nothing Then
                                childIdProp.SetValue(child, Convert.ChangeType(newChildId, childIdProp.PropertyType))
                            End If
                        End Using
                    Next
                Next

                transaction.Commit() ' sync
            End Using
        End Using

        Return obj
    End Function
    Public Async Function CreateInTableAsync(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of T) Implements IOrmAsync.CreateInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim properties = typeT.GetProperties().
        Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        Dim idProp = properties.FirstOrDefault(Function(p) p.Name = idColumn)
        If idProp Is Nothing Then
            Throw New InvalidOperationException($"The specified id column '{idColumn}' was not found on type '{typeT.Name}'.")
        End If

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()
            Using transaction = CType(connection, DbConnection).BeginTransaction()

                ' Insert parent
                Dim insertCols = New List(Of String)
                Dim insertVals = New List(Of String)
                Dim parameters = New List(Of IDataParameter)

                For Each prop In properties
                    If IsGenericList(prop.PropertyType) Then Continue For
                    If prop.Name = idColumn AndAlso IdWillAutoIncrement Then Continue For

                    Dim paramName = $"{parameterPrefix}{prop.Name}"
                    insertCols.Add($"[{prop.Name}]")
                    insertVals.Add(paramName)

                    parameters.Add(Await _provider.CreateParameterAsync(paramName, prop.GetValue(obj)))
                    'Dim value = prop.GetValue(obj)
                    'If prop.PropertyType.IsEnum Then
                    '    value = Convert.ToInt32(value) ' or value.ToString() if you want to store enum names
                    'End If
                    'parameters.Add(Await _provider.CreateParameterAsync(paramName, value))
                Next

                Dim sql = $"INSERT INTO [{tableName}] ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", insertVals)})"
                If IdWillAutoIncrement Then
                    sql += "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                End If

                Using cmd = Await _provider.CreateCommandAsync(sql, connection)
                    cmd.Transaction = transaction
                    For Each param In parameters
                        cmd.Parameters.Add(param)
                    Next

                    If IdWillAutoIncrement Then
                        Dim newId = Convert.ToInt64(Await CType(cmd, DbCommand).ExecuteScalarAsync())
                        idProp.SetValue(obj, Convert.ChangeType(newId, idProp.PropertyType))
                    Else
                        Await CType(cmd, DbCommand).ExecuteNonQueryAsync()
                    End If
                End Using

                ' Insert children
                For Each prop In properties.Where(Function(p) IsGenericList(p.PropertyType))
                    Dim childList = CType(prop.GetValue(obj), IEnumerable)
                    If childList Is Nothing Then Continue For

                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim childProps = childType.GetProperties().
                    Where(Function(p) p.CanRead AndAlso p.CanWrite AndAlso Not IsGenericList(p.PropertyType)).ToList()

                    Dim fkColumn = $"{tableName}_{idColumn}"

                    For Each child In childList
                        ' Set FK
                        Dim fkProp = childProps.FirstOrDefault(Function(p) p.Name = fkColumn)
                        If fkProp IsNot Nothing Then
                            fkProp.SetValue(child, idProp.GetValue(obj))
                        End If

                        Dim cCols = New List(Of String)
                        Dim cVals = New List(Of String)
                        Dim cParams = New List(Of IDataParameter)

                        For Each cp In childProps
                            If cp.Name = idColumn AndAlso IdWillAutoIncrement Then Continue For

                            Dim cpName = cp.Name
                            Dim cpValue = cp.GetValue(child)

                            cCols.Add($"[{cpName}]")
                            cVals.Add($"{parameterPrefix}{cpName}")
                            cParams.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}{cpName}", cpValue))
                        Next

                        Dim childSql = $"INSERT INTO [{childTable}] ({String.Join(", ", cCols)}) VALUES ({String.Join(", ", cVals)}); SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"

                        Using cmd = Await _provider.CreateCommandAsync(childSql, connection)
                            cmd.Transaction = transaction
                            For Each p In cParams
                                cmd.Parameters.Add(p)
                            Next

                            Dim newChildId = Convert.ToInt64(Await CType(cmd, DbCommand).ExecuteScalarAsync())
                            Dim childIdProp = childProps.FirstOrDefault(Function(p) p.Name = idColumn)
                            If childIdProp IsNot Nothing Then
                                childIdProp.SetValue(child, Convert.ChangeType(newChildId, childIdProp.PropertyType))
                            End If
                        End Using
                    Next
                Next

                transaction.Commit()
            End Using
        End Using

        Return obj
    End Function

    Public Async Function CreateOrUpdateAsync(Of T)(obj As T, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of T) Implements IOrmAsync.CreateOrUpdateAsync
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim idValue = GetIdValue(obj, idColumn)

        If Await ExistsByIdInTableAsync(idValue, tableName, idColumn) Then
            Return Await UpdateAsync(Of T)(obj, idColumn)
        End If

        Dim cleanedObj = CleanOutIds(Of T)(obj)
        Return Await CreateAsync(Of T)(cleanedObj, idColumn, IdWillAutoIncrement)
    End Function
    Public Async Function CreateOrUpdateInTableAsync(Of T)(obj As T, tableName As String, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As Task(Of T) Implements IOrmAsync.CreateOrUpdateInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim idValue = GetIdValue(obj, idColumn)

        If Await ExistsByIdInTableAsync(idValue, tableName, idColumn) Then
            Return Await UpdateInTableAsync(Of T)(obj, tableName, idColumn)
        End If

        Return Await CreateInTableAsync(Of T)(CleanOutIds(Of T)(obj), tableName, idColumn, IdWillAutoIncrement)
    End Function

    Public Async Function ExistsByIdAsync(Of T)(id As Object, Optional idColumn As String = "Id") As Task(Of Boolean) Implements IOrmAsync.ExistsByIdAsync
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim paramName = $"{parameterPrefix}{idColumn}"

        Dim query = $"SELECT COUNT(1) FROM [{tableName}] WHERE [{idColumn}] = {paramName}"

        Using connection = Await _provider.CreateConnectionAsync()
            Using command = Await _provider.CreateCommandAsync(query, connection)
                command.Parameters.Add(Await _provider.CreateParameterAsync(paramName, id))
                connection.Open()
                Dim count = Convert.ToInt32(Await command.ExecuteScalarAsync()) ' <== Async read
                Return count > 0
            End Using
        End Using
    End Function
    Public Async Function ExistsByIdInTableAsync(id As Object, tableName As String, Optional idColumn As String = "Id") As Task(Of Boolean) Implements IOrmAsync.ExistsByIdInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        'Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim paramName = $"{parameterPrefix}{idColumn}"

        Dim query = $"SELECT COUNT(1) FROM [{tableName}] WHERE [{idColumn}] = {paramName}"

        Using connection = Await _provider.CreateConnectionAsync()
            Using command = Await _provider.CreateCommandAsync(query, connection)
                command.Parameters.Add(Await _provider.CreateParameterAsync(paramName, id))
                connection.Open()
                Dim count = Convert.ToInt32(Await command.ExecuteScalarAsync()) ' <== Async read
                Return count > 0
            End Using
        End Using
    End Function

    Public Async Function ExistsByAsync(Of T)(conditions As List(Of Condition)) As Task(Of Boolean) Implements IOrmAsync.ExistsByAsync

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        Dim tableName As String = GetType(T).Name
        Dim paramPrefix As String = _provider.GetParameterPrefix()
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        For i = 0 To conditions.Count - 1
            Dim condition = conditions(i)
            Dim paramName = $"{paramPrefix}p{i}"
            Dim sqlOp = GetSqlOperator(condition.SqlComparison)
            whereClauses.Add($"{condition.Column} {sqlOp} {paramName}")
            parameters.Add(Await _provider.CreateParameterAsync(paramName, condition.Value))
        Next

        Dim whereSql = String.Join(" AND ", whereClauses)
        Dim query = $"SELECT COUNT(1) FROM {tableName} WHERE {whereSql}"

        Using conn = Await _provider.CreateConnectionAsync()
            Using cmd = Await _provider.CreateCommandAsync(query, conn)
                For Each param In parameters
                    cmd.Parameters.Add(param)
                Next

                conn.Open() ' connection.OpenAsync() is not available; still sync
                Dim count = Convert.ToInt64(Await cmd.ExecuteScalarAsync())
                Return count > 0
            End Using
        End Using
    End Function

    Public Async Function ExistsByInTableAsync(Of T)(conditions As List(Of Condition), tableName As String) As Task(Of Boolean) Implements IOrmAsync.ExistsByInTableAsync

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim paramPrefix As String = _provider.GetParameterPrefix()
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        For i = 0 To conditions.Count - 1
            Dim condition = conditions(i)
            Dim paramName = $"{paramPrefix}p{i}"
            Dim sqlOp = GetSqlOperator(condition.SqlComparison)
            whereClauses.Add($"{condition.Column} {sqlOp} {paramName}")
            parameters.Add(Await _provider.CreateParameterAsync(paramName, condition.Value))
        Next

        Dim whereSql = String.Join(" AND ", whereClauses)
        Dim query = $"SELECT COUNT(1) FROM {tableName} WHERE {whereSql}"

        Using conn = Await _provider.CreateConnectionAsync()
            Using cmd = Await _provider.CreateCommandAsync(query, conn)
                For Each param In parameters
                    cmd.Parameters.Add(param)
                Next

                conn.Open() ' connection.OpenAsync() is not available; still sync
                Dim count = Convert.ToInt64(Await cmd.ExecuteScalarAsync())
                Return count > 0
            End Using
        End Using
    End Function

    Public Async Function FindAllAsync(Of T)(Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Task(Of List(Of T)) Implements IOrmAsync.FindAllAsync
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim results As New List(Of T)

        Using connection = Await _provider.CreateConnectionAsync()
            Await DirectCast(connection, DbConnection).OpenAsync()

            ' Fetch all parent rows
            Dim sql = $"SELECT * FROM [{tableName}]"
            Using command = Await _provider.CreateCommandAsync(sql, connection)
                Using reader = Await DirectCast(command, DbCommand).ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        Dim obj = Activator.CreateInstance(Of T)()
                        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite AndAlso Not IsGenericList(p.PropertyType)).ToList()

                        For Each prop In props
                            If Not ColumnExists(reader, prop.Name) Then Continue For
                            If Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then Continue For

                            Dim val = reader(prop.Name)
                            Dim safeValue = GetSafeEnumValue(prop.PropertyType, val)
                            prop.SetValue(obj, safeValue)
                        Next

                        results.Add(obj)
                    End While
                End Using
            End Using

            ' Load child collections
            For Each obj In results
                Dim props = typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType)).ToList()

                For Each prop In props
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim fkColumn = $"{tableName}_{idColumn}"
                    Dim parentId = typeT.GetProperty(idColumn).GetValue(obj)

                    Dim childList = Activator.CreateInstance(GetType(List(Of )).MakeGenericType(childType))

                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}id"
                    Using cmd = Await _provider.CreateCommandAsync(childSql, connection)
                        cmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", parentId))

                        Using reader = Await DirectCast(cmd, DbCommand).ExecuteReaderAsync()
                            While Await reader.ReadAsync()
                                Dim childObj = Activator.CreateInstance(childType)
                                Dim childProps = childType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

                                For Each cp In childProps
                                    If Not ColumnExists(reader, cp.Name) Then Continue For
                                    If Await reader.IsDBNullAsync(reader.GetOrdinal(cp.Name)) Then Continue For

                                    Dim val = reader(cp.Name)
                                    Dim safeValue = GetSafeEnumValue(cp.PropertyType, val)
                                    cp.SetValue(childObj, safeValue)
                                Next

                                CType(childList, IList).Add(childObj)
                            End While
                        End Using
                    End Using

                    prop.SetValue(obj, childList)
                Next
            Next
        End Using

        ' Sort results in-memory based on idColumn
        Dim idProp = typeT.GetProperty(idColumn)
        If idProp IsNot Nothing Then
            If ascending Then
                results = results.OrderBy(Function(x) idProp.GetValue(x)).ToList()
            Else
                results = results.OrderByDescending(Function(x) idProp.GetValue(x)).ToList()
            End If
        End If

        Return results
    End Function

    Public Async Function FindAllInTableAsync(Of T)(tableName As String, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Task(Of List(Of T)) Implements IOrmAsync.FindAllInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim results As New List(Of T)

        Using connection = Await _provider.CreateConnectionAsync()
            Await DirectCast(connection, DbConnection).OpenAsync()

            ' Fetch all parent rows
            Dim sql = $"SELECT * FROM [{tableName}]"
            Using command = Await _provider.CreateCommandAsync(sql, connection)
                Using reader = Await DirectCast(command, DbCommand).ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        Dim obj = Activator.CreateInstance(Of T)()
                        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite AndAlso Not IsGenericList(p.PropertyType)).ToList()

                        For Each prop In props
                            If Not ColumnExists(reader, prop.Name) Then Continue For
                            If Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then Continue For

                            Dim val = reader(prop.Name)
                            Dim safeValue = GetSafeEnumValue(prop.PropertyType, val)
                            prop.SetValue(obj, safeValue)
                        Next

                        results.Add(obj)
                    End While
                End Using
            End Using

            ' Load child collections
            For Each obj In results
                Dim props = typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType)).ToList()

                For Each prop In props
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim fkColumn = $"{tableName}_{idColumn}"
                    Dim parentId = typeT.GetProperty(idColumn).GetValue(obj)

                    Dim childList = Activator.CreateInstance(GetType(List(Of )).MakeGenericType(childType))

                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}id"
                    Using cmd = Await _provider.CreateCommandAsync(childSql, connection)
                        cmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", parentId))

                        Using reader = Await DirectCast(cmd, DbCommand).ExecuteReaderAsync()
                            While Await reader.ReadAsync()
                                Dim childObj = Activator.CreateInstance(childType)
                                Dim childProps = childType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

                                For Each cp In childProps
                                    If Not ColumnExists(reader, cp.Name) Then Continue For
                                    If Await reader.IsDBNullAsync(reader.GetOrdinal(cp.Name)) Then Continue For

                                    Dim val = reader(cp.Name)
                                    Dim safeValue = GetSafeEnumValue(cp.PropertyType, val)
                                    cp.SetValue(childObj, safeValue)
                                Next

                                CType(childList, IList).Add(childObj)
                            End While
                        End Using
                    End Using

                    prop.SetValue(obj, childList)
                Next
            Next
        End Using

        ' Sort results in-memory based on idColumn
        Dim idProp = typeT.GetProperty(idColumn)
        If idProp IsNot Nothing Then
            If ascending Then
                results = results.OrderBy(Function(x) idProp.GetValue(x)).ToList()
            Else
                results = results.OrderByDescending(Function(x) idProp.GetValue(x)).ToList()
            End If
        End If

        Return results
    End Function

    Public Async Function FindAllPagedAsync(Of T)(pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Task(Of Page(Of T)) Implements IOrmAsync.FindAllPagedAsync
        If pageNumber < 1 Then Throw New ArgumentOutOfRangeException(NameOf(pageNumber))
        If maxPerPage < 1 Then Throw New ArgumentOutOfRangeException(NameOf(maxPerPage))

        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim offset = (pageNumber - 1) * maxPerPage
        Dim totalSql = $"SELECT COUNT(1) FROM [{tableName}]"
        Dim totalRecords As Long
        Dim records = New List(Of T)

        Using connection = CType(Await _provider.CreateConnectionAsync(), SqlConnection)
            Await connection.OpenAsync()

            ' Get total count
            Using countCmd = Await _provider.CreateCommandAsync(totalSql, connection)
                totalRecords = Convert.ToInt64(Await countCmd.ExecuteScalarAsync())
            End Using

            If totalRecords = 0 Then
                Return New Page(Of T)(New List(Of T), pageNumber, maxPerPage, 0, 0)
            End If

            ' Fetch paged rows
            Dim sql = $"
        SELECT * FROM (
            SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS RowNum
            FROM [{tableName}]
        ) AS Paged
        WHERE RowNum BETWEEN {offset + 1} AND {offset + maxPerPage}
        "

            Using cmd = Await _provider.CreateCommandAsync(sql, connection)
                Using reader = Await cmd.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        Dim obj = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                                Dim rawVal = reader(prop.Name)
                                Dim safeVal = GetSafeEnumValue(prop.PropertyType, rawVal)
                                prop.SetValue(obj, safeVal)
                            End If
                        Next
                        records.Add(obj)
                    End While
                End Using
            End Using

            ' Load child collections
            For Each obj In records
                For Each prop In typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType) AndAlso p.CanWrite)
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim fkColumn = $"{tableName}_{idColumn}"
                    Dim parentId = typeT.GetProperty(idColumn).GetValue(obj)
                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}ParentId"

                    Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                        childCmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}ParentId", parentId))
                        Using reader = Await childCmd.ExecuteReaderAsync()
                            Dim childList = CType(Activator.CreateInstance(prop.PropertyType), IList)
                            While Await reader.ReadAsync()
                                Dim childObj = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If ColumnExists(reader, cp.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(cp.Name)) Then
                                        Dim rawVal = reader(cp.Name)
                                        Dim safeVal = GetSafeEnumValue(cp.PropertyType, rawVal)
                                        cp.SetValue(childObj, safeVal)
                                    End If
                                Next
                                childList.Add(childObj)
                            End While
                            prop.SetValue(obj, childList)
                        End Using
                    End Using
                Next
            Next
        End Using

        ' Sort records in-memory after paging
        Dim idProp = typeT.GetProperty(idColumn)
        If idProp IsNot Nothing Then
            records = If(ascending,
                     records.OrderBy(Function(x) idProp.GetValue(x)).ToList(),
                     records.OrderByDescending(Function(x) idProp.GetValue(x)).ToList())
        End If

        Dim pageCount = CInt(Math.Ceiling(totalRecords / CDbl(maxPerPage)))
        Return New Page(Of T)(records, pageNumber, maxPerPage, totalRecords, pageCount)
    End Function

    Public Async Function FindAllPagedInTableAsync(Of T)(tableName As String, pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Task(Of Page(Of T)) Implements IOrmAsync.FindAllPagedInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")

        If pageNumber < 1 Then Throw New ArgumentOutOfRangeException(NameOf(pageNumber))
        If maxPerPage < 1 Then Throw New ArgumentOutOfRangeException(NameOf(maxPerPage))

        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim offset = (pageNumber - 1) * maxPerPage
        Dim totalSql = $"SELECT COUNT(1) FROM [{tableName}]"
        Dim totalRecords As Long
        Dim records = New List(Of T)

        Using connection = CType(Await _provider.CreateConnectionAsync(), SqlConnection)
            Await connection.OpenAsync()

            ' Get total count
            Using countCmd = Await _provider.CreateCommandAsync(totalSql, connection)
                totalRecords = Convert.ToInt64(Await countCmd.ExecuteScalarAsync())
            End Using

            If totalRecords = 0 Then
                Return New Page(Of T)(New List(Of T), pageNumber, maxPerPage, 0, 0)
            End If

            ' Fetch paged rows
            Dim sql = $"
        SELECT * FROM (
            SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS RowNum
            FROM [{tableName}]
        ) AS Paged
        WHERE RowNum BETWEEN {offset + 1} AND {offset + maxPerPage}
        "

            Using cmd = Await _provider.CreateCommandAsync(sql, connection)
                Using reader = Await cmd.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        Dim obj = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                                Dim rawVal = reader(prop.Name)
                                Dim safeVal = GetSafeEnumValue(prop.PropertyType, rawVal)
                                prop.SetValue(obj, safeVal)
                            End If
                        Next
                        records.Add(obj)
                    End While
                End Using
            End Using

            ' Load child collections
            For Each obj In records
                For Each prop In typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType) AndAlso p.CanWrite)
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim fkColumn = $"{tableName}_{idColumn}"
                    Dim parentId = typeT.GetProperty(idColumn).GetValue(obj)
                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}ParentId"

                    Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                        childCmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}ParentId", parentId))
                        Using reader = Await childCmd.ExecuteReaderAsync()
                            Dim childList = CType(Activator.CreateInstance(prop.PropertyType), IList)
                            While Await reader.ReadAsync()
                                Dim childObj = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If ColumnExists(reader, cp.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(cp.Name)) Then
                                        Dim rawVal = reader(cp.Name)
                                        Dim safeVal = GetSafeEnumValue(cp.PropertyType, rawVal)
                                        cp.SetValue(childObj, safeVal)
                                    End If
                                Next
                                childList.Add(childObj)
                            End While
                            prop.SetValue(obj, childList)
                        End Using
                    End Using
                Next
            Next
        End Using

        ' Sort records in-memory after paging
        Dim idProp = typeT.GetProperty(idColumn)
        If idProp IsNot Nothing Then
            records = If(ascending,
                     records.OrderBy(Function(x) idProp.GetValue(x)).ToList(),
                     records.OrderByDescending(Function(x) idProp.GetValue(x)).ToList())
        End If

        Dim pageCount = CInt(Math.Ceiling(totalRecords / CDbl(maxPerPage)))
        Return New Page(Of T)(records, pageNumber, maxPerPage, totalRecords, pageCount)
    End Function

    Public Async Function FindByIdAsync(Of T)(id As Object, Optional idColumn As String = "Id") As Task(Of T) Implements IOrmAsync.FindByIdAsync
        Dim typeT = GetType(T)
        Dim tableName As String = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim query As String = $"SELECT * FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"

        Dim obj As T = Nothing

        Using connection = Await _provider.CreateConnectionAsync()
            Await DirectCast(connection, DbConnection).OpenAsync()

            Using command = Await _provider.CreateCommandAsync(query, CType(connection, SqlConnection))
                command.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", id))

                Using reader = Await DirectCast(command, DbCommand).ExecuteReaderAsync()
                    If Await reader.ReadAsync() Then
                        obj = Activator.CreateInstance(Of T)()

                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                                Dim rawValue = reader(prop.Name)
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                                prop.SetValue(obj, safeValue)
                            End If
                        Next
                    End If
                End Using ' reader
            End Using ' command

            ' Now fetch child collections
            If obj IsNot Nothing Then
                For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso IsGenericList(p.PropertyType))
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim foreignKey = $"{tableName}_{idColumn}"
                    Dim childList = Await GetChildRecordsAsync(connection, id, tableName, idColumn, childType, foreignKey, _provider)
                    prop.SetValue(obj, childList)
                Next
            End If
        End Using ' connection

        Return obj
    End Function

    Public Async Function FindByIdInTableAsync(Of T)(id As Object, tableName As String, Optional idColumn As String = "Id") As Task(Of T) Implements IOrmAsync.FindByIdInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim query As String = $"SELECT * FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"

        Dim obj As T = Nothing

        Using connection = Await _provider.CreateConnectionAsync()
            Await DirectCast(connection, DbConnection).OpenAsync()

            Using command = Await _provider.CreateCommandAsync(query, CType(connection, SqlConnection))
                command.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", id))

                Using reader = Await DirectCast(command, DbCommand).ExecuteReaderAsync()
                    If Await reader.ReadAsync() Then
                        obj = Activator.CreateInstance(Of T)()

                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                                Dim rawValue = reader(prop.Name)
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                                prop.SetValue(obj, safeValue)
                            End If
                        Next
                    End If
                End Using ' reader
            End Using ' command

            ' Now fetch child collections
            If obj IsNot Nothing Then
                For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso IsGenericList(p.PropertyType))
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim foreignKey = $"{tableName}_{idColumn}"
                    Dim childList = Await GetChildRecordsAsync(connection, id, tableName, idColumn, childType, foreignKey, _provider)
                    prop.SetValue(obj, childList)
                Next
            End If
        End Using ' connection

        Return obj
    End Function

    Public Async Function FindByAsync(Of T)(conditions As List(Of Condition)) As Task(Of List(Of T)) Implements IOrmAsync.FindByAsync
        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If

        Dim results As New List(Of T)
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        For i = 0 To conditions.Count - 1
            Dim c = conditions(i)
            Dim paramName = $"{parameterPrefix}param{i}"
            Dim sqlOperator = GetSqlOperator(c.SqlComparison)
            Dim clause As String

            If c.SqlComparison = SqlComparisonOperator.SqlIn Then
                Dim values = DirectCast(c.Value, IEnumerable)
                Dim subParams As New List(Of String)
                Dim idx = 0
                For Each what In values
                    Dim subParam = $"{paramName}_{idx}"
                    subParams.Add(subParam)
                    parameters.Add(Await _provider.CreateParameterAsync(subParam, what))
                    idx += 1
                Next
                clause = $"[{c.Column}] IN ({String.Join(", ", subParams)})"
            Else
                clause = $"[{c.Column}] {sqlOperator} {paramName}"
                parameters.Add(Await _provider.CreateParameterAsync(paramName, c.Value))
            End If

            whereClauses.Add(clause)
        Next

        Dim whereSql = If(whereClauses.Count > 0, "WHERE " & String.Join(" AND ", whereClauses), "")
        Dim sql = $"SELECT * FROM [{tableName}] {whereSql}"

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()

            Using command = Await _provider.CreateCommandAsync(sql, connection)
                For Each param In parameters
                    command.Parameters.Add(param)
                Next

                Using reader = Await command.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        Dim instance = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                                Dim rawValue = reader(prop.Name)
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                                prop.SetValue(instance, safeValue)
                            End If
                        Next
                        results.Add(instance)
                    End While
                End Using

                ' Load child collections
                For Each parent In results
                    For Each prop In typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType) AndAlso p.CanWrite)
                        Dim childType = prop.PropertyType.GetGenericArguments()(0)
                        Dim childTable = childType.Name
                        Dim childFk = $"{tableName}_Id"
                        Dim parentId = typeT.GetProperty("Id")?.GetValue(parent)
                        Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{childFk}] = {parameterPrefix}parentId"

                        Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                            childCmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}parentId", parentId))
                            Using childReader = Await childCmd.ExecuteReaderAsync()
                                Dim childList = CType(Activator.CreateInstance(prop.PropertyType), IList)
                                While Await childReader.ReadAsync()
                                    Dim childInstance = Activator.CreateInstance(childType)
                                    For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                        If ColumnExists(childReader, cp.Name) AndAlso Not Await childReader.IsDBNullAsync(childReader.GetOrdinal(cp.Name)) Then
                                            Dim rawVal = childReader(cp.Name)
                                            Dim safeVal = GetSafeEnumValue(cp.PropertyType, rawVal)
                                            cp.SetValue(childInstance, safeVal)
                                        End If
                                    Next
                                    childList.Add(childInstance)
                                End While
                                prop.SetValue(parent, childList)
                            End Using
                        End Using
                    Next
                Next
            End Using
        End Using

        Return results
    End Function

    Public Async Function FindByInTableAsync(Of T)(conditions As List(Of Condition), tableName As String) As Task(Of List(Of T)) Implements IOrmAsync.FindByInTableAsync

        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If

        Dim results As New List(Of T)
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        For i = 0 To conditions.Count - 1
            Dim c = conditions(i)
            Dim paramName = $"{parameterPrefix}param{i}"
            Dim sqlOperator = GetSqlOperator(c.SqlComparison)
            Dim clause As String

            If c.SqlComparison = SqlComparisonOperator.SqlIn Then
                Dim values = DirectCast(c.Value, IEnumerable)
                Dim subParams As New List(Of String)
                Dim idx = 0
                For Each what In values
                    Dim subParam = $"{paramName}_{idx}"
                    subParams.Add(subParam)
                    parameters.Add(Await _provider.CreateParameterAsync(subParam, what))
                    idx += 1
                Next
                clause = $"[{c.Column}] IN ({String.Join(", ", subParams)})"
            Else
                clause = $"[{c.Column}] {sqlOperator} {paramName}"
                parameters.Add(Await _provider.CreateParameterAsync(paramName, c.Value))
            End If

            whereClauses.Add(clause)
        Next

        Dim whereSql = If(whereClauses.Count > 0, "WHERE " & String.Join(" AND ", whereClauses), "")
        Dim sql = $"SELECT * FROM [{tableName}] {whereSql}"

        Using connection = Await _provider.CreateConnectionAsync()
            Await CType(connection, DbConnection).OpenAsync()

            Using command = Await _provider.CreateCommandAsync(sql, connection)
                For Each param In parameters
                    command.Parameters.Add(param)
                Next

                Using reader = Await command.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        Dim instance = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                                Dim rawValue = reader(prop.Name)
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                                prop.SetValue(instance, safeValue)
                            End If
                        Next
                        results.Add(instance)
                    End While
                End Using

                ' Load child collections
                For Each parent In results
                    For Each prop In typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType) AndAlso p.CanWrite)
                        Dim childType = prop.PropertyType.GetGenericArguments()(0)
                        Dim childTable = childType.Name
                        Dim childFk = $"{tableName}_Id"
                        Dim parentId = typeT.GetProperty("Id")?.GetValue(parent)
                        Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{childFk}] = {parameterPrefix}parentId"

                        Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                            childCmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}parentId", parentId))
                            Using childReader = Await childCmd.ExecuteReaderAsync()
                                Dim childList = CType(Activator.CreateInstance(prop.PropertyType), IList)
                                While Await childReader.ReadAsync()
                                    Dim childInstance = Activator.CreateInstance(childType)
                                    For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                        If ColumnExists(childReader, cp.Name) AndAlso Not Await childReader.IsDBNullAsync(childReader.GetOrdinal(cp.Name)) Then
                                            Dim rawVal = childReader(cp.Name)
                                            Dim safeVal = GetSafeEnumValue(cp.PropertyType, rawVal)
                                            cp.SetValue(childInstance, safeVal)
                                        End If
                                    Next
                                    childList.Add(childInstance)
                                End While
                                prop.SetValue(parent, childList)
                            End Using
                        End Using
                    Next
                Next
            End Using
        End Using

        Return results

    End Function

    Public Async Function FindByPagedAsync(Of T)(conditions As List(Of Condition), pageNumber As Integer, maxPerPage As Integer) As Task(Of Page(Of T)) Implements IOrmAsync.FindByPagedAsync

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()

        ' Validate page params
        If pageNumber < 1 Then pageNumber = 1
        If maxPerPage < 1 Then maxPerPage = 10

        ' Build WHERE clause
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of SqlParameter)

        For i = 0 To conditions.Count - 1
            Dim c = conditions(i)
            Dim paramName = $"{parameterPrefix}param{i}"
            Dim sqlOperator = GetSqlOperator(c.SqlComparison)
            Dim clause As String

            If c.SqlComparison = SqlComparisonOperator.SqlIn Then
                Dim values = DirectCast(c.Value, IEnumerable)
                Dim subParams As New List(Of String)
                Dim idx = 0
                For Each whatever In values
                    Dim subParam = $"{paramName}_{idx}"
                    subParams.Add(subParam)
                    parameters.Add(Await _provider.CreateParameterAsync(subParam, whatever))
                    idx += 1
                Next
                clause = $"[{c.Column}] IN ({String.Join(", ", subParams)})"
            Else
                clause = $"[{c.Column}] {sqlOperator} {paramName}"
                parameters.Add(Await _provider.CreateParameterAsync(paramName, c.Value))
            End If

            whereClauses.Add(clause)
        Next

        Dim whereSql = If(whereClauses.Count > 0, "WHERE " & String.Join(" AND ", whereClauses), "")

        Dim rowStart = (pageNumber - 1) * maxPerPage + 1
        Dim rowEnd = rowStart + maxPerPage - 1

        Dim sql = $"
    WITH Results_CTE AS (
        SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS RowNum
        FROM [{tableName}]
        {whereSql}
    )
    SELECT * FROM Results_CTE
    WHERE RowNum BETWEEN {rowStart} AND {rowEnd}"

        Dim results As New List(Of T)

        Using connection = CType(Await _provider.CreateConnectionAsync(), SqlConnection)
            Await connection.OpenAsync()

            ' Query main page results
            Using command = Await _provider.CreateCommandAsync(sql, connection)
                command.Parameters.AddRange(parameters.ToArray())

                Using reader = Await command.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        Dim instance = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                                prop.SetValue(instance, reader(prop.Name))
                            End If
                        Next
                        results.Add(instance)
                    End While
                End Using
            End Using

            ' Load child collections
            For Each parent In results
                For Each prop In typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType))
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim childFk = $"{tableName}_Id"
                    Dim parentId = typeT.GetProperty("Id")?.GetValue(parent)
                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{childFk}] = {parameterPrefix}parentId"

                    Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                        childCmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}parentId", parentId))

                        Using childReader = Await childCmd.ExecuteReaderAsync()
                            Dim childList = Activator.CreateInstance(prop.PropertyType)
                            While Await childReader.ReadAsync()
                                Dim childInstance = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If ColumnExists(childReader, cp.Name) AndAlso Not Await childReader.IsDBNullAsync(childReader.GetOrdinal(cp.Name)) Then
                                        cp.SetValue(childInstance, childReader(cp.Name))
                                    End If
                                Next
                                DirectCast(childList, IList).Add(childInstance)
                            End While
                            prop.SetValue(parent, childList)
                        End Using
                    End Using
                Next
            Next

            ' Count query
            Dim countSql = $"SELECT COUNT(1) FROM [{tableName}] {whereSql}"
            Dim totalCount As Long
            Using countCmd = Await _provider.CreateCommandAsync(countSql, connection)
                For Each param In parameters
                    countCmd.Parameters.Add(Await _provider.CloneParameterAsync(param))
                Next
                totalCount = Convert.ToInt64(Await countCmd.ExecuteScalarAsync())
            End Using

            Dim pageCount = Math.Ceiling(totalCount / maxPerPage)

            Return New Page(Of T)(results, pageNumber, maxPerPage, totalCount, CInt(pageCount))
        End Using
    End Function

    Public Async Function FindByPagedInTableAsync(Of T)(conditions As List(Of Condition), tableName As String, pageNumber As Integer, maxPerPage As Integer) As Task(Of Page(Of T)) Implements IOrmAsync.FindByPagedInTableAsync

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()

        ' Validate page params
        If pageNumber < 1 Then pageNumber = 1
        If maxPerPage < 1 Then maxPerPage = 10

        ' Build WHERE clause
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of SqlParameter)

        For i = 0 To conditions.Count - 1
            Dim c = conditions(i)
            Dim paramName = $"{parameterPrefix}param{i}"
            Dim sqlOperator = GetSqlOperator(c.SqlComparison)
            Dim clause As String

            If c.SqlComparison = SqlComparisonOperator.SqlIn Then
                Dim values = DirectCast(c.Value, IEnumerable)
                Dim subParams As New List(Of String)
                Dim idx = 0
                For Each whatever In values
                    Dim subParam = $"{paramName}_{idx}"
                    subParams.Add(subParam)
                    parameters.Add(Await _provider.CreateParameterAsync(subParam, whatever))
                    idx += 1
                Next
                clause = $"[{c.Column}] IN ({String.Join(", ", subParams)})"
            Else
                clause = $"[{c.Column}] {sqlOperator} {paramName}"
                parameters.Add(Await _provider.CreateParameterAsync(paramName, c.Value))
            End If

            whereClauses.Add(clause)
        Next

        Dim whereSql = If(whereClauses.Count > 0, "WHERE " & String.Join(" AND ", whereClauses), "")

        Dim rowStart = (pageNumber - 1) * maxPerPage + 1
        Dim rowEnd = rowStart + maxPerPage - 1

        Dim sql = $"
    WITH Results_CTE AS (
        SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS RowNum
        FROM [{tableName}]
        {whereSql}
    )
    SELECT * FROM Results_CTE
    WHERE RowNum BETWEEN {rowStart} AND {rowEnd}"

        Dim results As New List(Of T)

        Using connection = CType(Await _provider.CreateConnectionAsync(), SqlConnection)
            Await connection.OpenAsync()

            ' Query main page results
            Using command = Await _provider.CreateCommandAsync(sql, connection)
                command.Parameters.AddRange(parameters.ToArray())

                Using reader = Await command.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        Dim instance = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                                prop.SetValue(instance, reader(prop.Name))
                            End If
                        Next
                        results.Add(instance)
                    End While
                End Using
            End Using

            ' Load child collections
            For Each parent In results
                For Each prop In typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType))
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim childFk = $"{tableName}_Id"
                    Dim parentId = typeT.GetProperty("Id")?.GetValue(parent)
                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{childFk}] = {parameterPrefix}parentId"

                    Using childCmd = Await _provider.CreateCommandAsync(childSql, connection)
                        childCmd.Parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}parentId", parentId))

                        Using childReader = Await childCmd.ExecuteReaderAsync()
                            Dim childList = Activator.CreateInstance(prop.PropertyType)
                            While Await childReader.ReadAsync()
                                Dim childInstance = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If ColumnExists(childReader, cp.Name) AndAlso Not Await childReader.IsDBNullAsync(childReader.GetOrdinal(cp.Name)) Then
                                        cp.SetValue(childInstance, childReader(cp.Name))
                                    End If
                                Next
                                DirectCast(childList, IList).Add(childInstance)
                            End While
                            prop.SetValue(parent, childList)
                        End Using
                    End Using
                Next
            Next

            ' Count query
            Dim countSql = $"SELECT COUNT(1) FROM [{tableName}] {whereSql}"
            Dim totalCount As Long
            Using countCmd = Await _provider.CreateCommandAsync(countSql, connection)
                For Each param In parameters
                    countCmd.Parameters.Add(Await _provider.CloneParameterAsync(param))
                Next
                totalCount = Convert.ToInt64(Await countCmd.ExecuteScalarAsync())
            End Using

            Dim pageCount = Math.Ceiling(totalCount / maxPerPage)

            Return New Page(Of T)(results, pageNumber, maxPerPage, totalCount, CInt(pageCount))
        End Using
    End Function

    Public Async Function UpdateAsync(Of T)(obj As T, Optional idColumn As String = "Id") As Task(Of T) Implements IOrmAsync.UpdateAsync
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim properties = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()
        Dim idProp = properties.FirstOrDefault(Function(p) p.Name = idColumn)
        If idProp Is Nothing Then Throw New InvalidOperationException($"The specified id column '{idColumn}' was not found on type '{typeT.Name}'.")

        Dim idValue = idProp.GetValue(obj)

        ' Check for RowVersion property (if present and writable)
        Dim rowVersionProp = properties.FirstOrDefault(Function(p) p.Name = "RowVersion")
        Dim hasRowVersion = rowVersionProp IsNot Nothing

        Using connection = Await _provider.CreateConnectionAsync()
            connection.Open() ' IDbConnection doesn't have OpenAsync, so open synchronously
            Using transaction = connection.BeginTransaction()

                ' Update parent row
                Dim setClauses = New List(Of String)
                Dim parameters = New List(Of IDataParameter)

                For Each prop In properties
                    If IsGenericList(prop.PropertyType) OrElse prop.Name = idColumn OrElse prop.Name = "RowVersion" Then Continue For

                    Dim paramName = $"{parameterPrefix}{prop.Name}"
                    setClauses.Add($"[{prop.Name}] = {paramName}")
                    parameters.Add(Await _provider.CreateParameterAsync(paramName, prop.GetValue(obj)))
                    'Dim value = prop.GetValue(obj)
                    'If prop.PropertyType.IsEnum Then
                    '    value = Convert.ToInt32(value) ' or value.ToString() if you want to store enum names
                    'End If
                    'parameters.Add(Await _provider.CreateParameterAsync(paramName, value))
                Next

                parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", idValue))

                ' Add RowVersion condition if available
                Dim whereClause = $"[{idColumn}] = {parameterPrefix}id"
                If hasRowVersion Then
                    whereClause &= $" AND [RowVersion] = {parameterPrefix}RowVersion"
                    parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}RowVersion", rowVersionProp.GetValue(obj)))
                End If

                Dim sql = $"UPDATE [{tableName}] SET {String.Join(", ", setClauses)} WHERE {whereClause}"

                Using command = Await _provider.CreateCommandAsync(sql, connection)
                    command.Transaction = transaction
                    For Each param In parameters
                        command.Parameters.Add(param)
                    Next

                    Dim affected = Await command.ExecuteNonQueryAsync()
                    If affected <> 1 Then
                        Throw New DBConcurrencyException("The record was updated by another user or no longer exists.")
                    End If
                End Using

                ' Child object handling placeholder
                ' (you can insert/merge/update child lists here if needed later)

                transaction.Commit()
            End Using
        End Using

        Return obj
    End Function
    Public Async Function UpdateInTableAsync(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id") As Task(Of T) Implements IOrmAsync.UpdateInTableAsync
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim properties = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()
        Dim idProp = properties.FirstOrDefault(Function(p) p.Name = idColumn)
        If idProp Is Nothing Then Throw New InvalidOperationException($"The specified id column '{idColumn}' was not found on type '{typeT.Name}'.")

        Dim idValue = idProp.GetValue(obj)

        ' Check for RowVersion property (if present and writable)
        Dim rowVersionProp = properties.FirstOrDefault(Function(p) p.Name = "RowVersion")
        Dim hasRowVersion = rowVersionProp IsNot Nothing

        Using connection = Await _provider.CreateConnectionAsync()
            connection.Open() ' IDbConnection doesn't have OpenAsync, so open synchronously
            Using transaction = connection.BeginTransaction()

                ' Update parent row
                Dim setClauses = New List(Of String)
                Dim parameters = New List(Of IDataParameter)

                For Each prop In properties
                    If IsGenericList(prop.PropertyType) OrElse prop.Name = idColumn OrElse prop.Name = "RowVersion" Then Continue For

                    Dim paramName = $"{parameterPrefix}{prop.Name}"
                    setClauses.Add($"[{prop.Name}] = {paramName}")
                    parameters.Add(Await _provider.CreateParameterAsync(paramName, prop.GetValue(obj)))
                    'Dim value = prop.GetValue(obj)
                    'If prop.PropertyType.IsEnum Then
                    '    value = Convert.ToInt32(value) ' or value.ToString() if you want to store enum names
                    'End If
                    'parameters.Add(Await _provider.CreateParameterAsync(paramName, value))
                Next

                parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}id", idValue))

                ' Add RowVersion condition if available
                Dim whereClause = $"[{idColumn}] = {parameterPrefix}id"
                If hasRowVersion Then
                    whereClause &= $" AND [RowVersion] = {parameterPrefix}RowVersion"
                    parameters.Add(Await _provider.CreateParameterAsync($"{parameterPrefix}RowVersion", rowVersionProp.GetValue(obj)))
                End If

                Dim sql = $"UPDATE [{tableName}] SET {String.Join(", ", setClauses)} WHERE {whereClause}"

                Using command = Await _provider.CreateCommandAsync(sql, connection)
                    command.Transaction = transaction
                    For Each param In parameters
                        command.Parameters.Add(param)
                    Next

                    Dim affected = Await command.ExecuteNonQueryAsync()
                    If affected <> 1 Then
                        Throw New DBConcurrencyException("The record was updated by another user or no longer exists.")
                    End If
                End Using

                ' Child object handling placeholder
                ' (you can insert/merge/update child lists here if needed later)

                transaction.Commit()
            End Using
        End Using

        Return obj
    End Function


#End Region


End Class
