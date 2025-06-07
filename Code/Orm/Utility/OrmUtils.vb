Imports System.Data.Common
Imports System.Reflection
Imports System.Xml.Serialization

Friend Class OrmUtils

#Region "props"
    Private Const Id As String = "Id"
#End Region


#Region "sync"

    Friend Shared Function GetEnumStorageType(prop As PropertyInfo) As EnumStorageType
        Dim attr = prop.GetCustomAttribute(Of EnumStorageAttribute)()
        If attr IsNot Nothing Then
            Return attr.StorageType
        End If
        Return EnumStorageType.Ordinal ' default
    End Function


    Friend Shared Function MapToSqlType(prop As PropertyInfo) As String
        ' Attribute override
        Dim attr = prop.GetCustomAttribute(Of SqlTypeAttribute)()
        If attr IsNot Nothing Then Return attr.SqlType

        ' Detect and unwrap Nullable<T>
        Dim rawType = Nullable.GetUnderlyingType(prop.PropertyType)
        If rawType Is Nothing Then rawType = prop.PropertyType

        ' Handle Enum storage
        If rawType.IsEnum Then
            Dim storageType = GetEnumStorageType(prop)
            Return If(storageType = EnumStorageType.String, "NVARCHAR(100)", "INT")
        End If

        ' Handle known types
        If rawType = GetType(Int16) Then Return "SMALLINT"
        If rawType = GetType(Int32) Then Return "INT"
        If rawType = GetType(Int64) Then Return "BIGINT"
        If rawType = GetType(Byte) Then Return "TINYINT"
        If rawType = GetType(Boolean) Then Return "BIT"
        If rawType = GetType(Decimal) Then Return "DECIMAL(18, 2)"
        If rawType = GetType(Double) Then Return "FLOAT"
        If rawType = GetType(Single) Then Return "REAL"
        If rawType = GetType(DateTime) Then Return "DATETIME"
        If rawType = GetType(DateTimeOffset) Then Return "DATETIMEOFFSET"
        If rawType = GetType(TimeSpan) Then Return "TIME"
        If rawType = GetType(Guid) Then Return "UNIQUEIDENTIFIER"
        If rawType = GetType(Byte()) Then Return "VARBINARY(MAX)"
        If rawType = GetType(Char) Then Return "NCHAR(1)"
        If rawType = GetType(String) Then Return "NVARCHAR(255)"
        If rawType = GetType(Char()) Then Return "NVARCHAR(MAX)" ' Treat char array as string

        ' Unknown fallback (potential log point)
        Return "NVARCHAR(MAX)"
    End Function

    Friend Shared Sub CreateOrUpdateTableRecursive(entityType As Type, idColumn As String, mode As DbPrepMode, connection As IDbConnection, transaction As IDbTransaction, executedTables As HashSet(Of String), logger As IDbLogger)
        Dim tableName = entityType.Name
        If executedTables.Contains(tableName) Then Return

        ' First ensure dependencies are created
        For Each prop In entityType.GetProperties()
            If IsGenericList(prop.PropertyType) Then
                Dim itemType = prop.PropertyType.GetGenericArguments()(0)

                If ShouldRecurseIntoType(itemType) Then
                    ' Many-to-Many detection
                    Dim inverseProp = itemType.GetProperties().FirstOrDefault(Function(p) IsGenericList(p.PropertyType) AndAlso p.PropertyType.GetGenericArguments()(0) Is entityType)
                    If inverseProp IsNot Nothing Then
                        ' Detected M2M, create join table
                        Dim joinTable = $"{entityType.Name}_{itemType.Name}"
                        Dim thisKey = $"{entityType.Name}_Id"
                        Dim otherKey = $"{itemType.Name}_Id"
                        Dim sql_query = $"CREATE TABLE IF NOT EXISTS [{joinTable}]([{thisKey}] INT NOT NULL, [{otherKey}] INT NOT NULL);"
                        logger?.LogInfo($"Creating join table: {joinTable}")
                        ExecuteNonQuery(sql_query, connection, transaction)
                        Continue For
                    End If

                    ' Otherwise, recurse for One-to-Many
                    CreateOrUpdateTableRecursive(itemType, idColumn, mode, connection, transaction, executedTables, logger)
                End If
            End If
        Next

        If executedTables.Contains(tableName) Then Return

        Dim sql As String
        Select Case mode
            Case DbPrepMode.Create
                sql = BuildCreateTableSql(entityType, idColumn)
            Case DbPrepMode.Update
                sql = BuildAlterTableSql(entityType, idColumn, connection, transaction)
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(mode))
        End Select

        If Not String.IsNullOrWhiteSpace(sql) Then
            logger?.LogInfo($"{mode} table: {tableName}")
            ExecuteNonQuery(sql, connection, transaction)
        End If

        ' Add fallback FK if needed
        For Each prop In entityType.GetProperties()
            If IsGenericList(prop.PropertyType) Then
                Dim itemType = prop.PropertyType.GetGenericArguments()(0)
                If Not itemType.GetProperties().Any(Function(p) p.Name = $"{entityType.Name}_Id") Then
                    Dim alter = $"ALTER TABLE [{itemType.Name}] ADD [{entityType.Name}_Id] INT;"
                    logger?.LogInfo($"Adding fallback foreign key: {alter}")
                    ExecuteNonQuery(alter, connection, transaction)
                End If
            End If
        Next

        executedTables.Add(tableName)
    End Sub
    Friend Shared Sub ExecuteNonQuery(sql As String, connection As IDbConnection, transaction As IDbTransaction)
        Dim statements = sql.Split(New String() {";" & Environment.NewLine, ";"}, StringSplitOptions.RemoveEmptyEntries)
        For Each stmt In statements
            Using cmd = connection.CreateCommand()
                cmd.Transaction = transaction
                cmd.CommandText = stmt.Trim()
                cmd.ExecuteNonQuery()
            End Using
        Next
    End Sub


    Friend Shared Function GetSqlOperator(op As SqlComparisonOperator) As String
        Select Case op
            Case SqlComparisonOperator.SqlEquals : Return "="
            Case SqlComparisonOperator.SqlNotEquals : Return "<>"
            Case SqlComparisonOperator.SqlGreaterThan : Return ">"
            Case SqlComparisonOperator.SqlLessThan : Return "<"
            Case SqlComparisonOperator.SqlGreaterThanOrEqual : Return ">="
            Case SqlComparisonOperator.SqlLessThanOrEqual : Return "<="
            Case SqlComparisonOperator.SqlLike : Return "LIKE"
            Case SqlComparisonOperator.SqlIn : Return "IN"
            Case Else : Throw New ArgumentOutOfRangeException()
        End Select
    End Function

    Friend Shared Function HasCascadeDeleteConstraint(tableName As String, connection As IDbConnection, transaction As IDbTransaction, _provider As IDbProvider) As Boolean
        Dim sql As String = "
        SELECT COUNT(*) 
        FROM sys.foreign_keys fk
        JOIN sys.tables parent ON fk.parent_object_id = parent.object_id
        WHERE parent.name = @TableName
          AND fk.delete_referential_action_desc = 'CASCADE'
    "

        Using cmd = _provider.CreateCommand(sql, connection)
            cmd.Transaction = transaction
            cmd.Parameters.Add(_provider.CreateParameter("@TableName", tableName))
            Dim result = Convert.ToInt32(cmd.ExecuteScalar())
            Return result > 0
        End Using
    End Function


    Friend Shared Function GetChildRecords(connection As IDbConnection, parentId As Object, parentTable As String, idColumn As String, childType As Type, foreignKeyColumn As String, _provider As IDbProvider) As Object
        Dim childTable As String = childType.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim query As String = $"SELECT * FROM [{childTable}] WHERE [{foreignKeyColumn}] = {parameterPrefix}parentId"

        Dim listType = GetType(List(Of )).MakeGenericType(childType)
        Dim list = Activator.CreateInstance(listType)

        Using command = _provider.CreateCommand(query, connection)
            command.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}parentId", parentId))
            Using reader = command.ExecuteReader()
                While reader.Read()
                    Dim child = Activator.CreateInstance(childType)
                    For Each prop In childType.GetProperties().Where(Function(p) p.CanWrite)
                        If Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
                            Dim rawValue = reader(prop.Name)
                            Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                            prop.SetValue(child, safeValue)
                        End If
                    Next
                    listType.GetMethod("Add").Invoke(list, {child})
                End While
            End Using
        End Using

        Return list
    End Function
    Friend Shared Function ColumnExists(reader As IDataReader, columnName As String) As Boolean
        For i As Integer = 0 To reader.FieldCount - 1
            If reader.GetName(i).Equals(columnName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    Friend Shared Function GetSafeEnumValue(targetType As Type, rawValue As Object) As Object
        If rawValue Is Nothing OrElse rawValue Is DBNull.Value Then
            Return Nothing
        End If

        If targetType.IsEnum Then
            Dim enumUnderlyingType = [Enum].GetUnderlyingType(targetType)

            Try
                ' Handle numeric types stored as strings or numbers
                Dim numericValue = Convert.ChangeType(rawValue, enumUnderlyingType)
                Return [Enum].ToObject(targetType, numericValue)
            Catch ex1 As Exception
                ' If the value is a string enum name like "Active"
                If TypeOf rawValue Is String Then
                    Try
                        Return [Enum].Parse(targetType, rawValue.ToString(), ignoreCase:=True)
                    Catch ex2 As Exception
                        Throw New ArgumentException($"Could not parse enum value '{rawValue}' for enum type '{targetType.Name}'", ex2)
                    End Try
                End If

                Throw New ArgumentException($"Cannot convert '{rawValue}' to enum '{targetType.Name}'", ex1)
            End Try
        End If

        ' Not an enum, do regular conversion
        Return Convert.ChangeType(rawValue, targetType)
    End Function

    Friend Shared Function IsGenericList(type As Type) As Boolean
        Return type.IsGenericType AndAlso type.GetGenericTypeDefinition() = GetType(List(Of ))
    End Function
    Friend Shared Function GetEnumDbValue(prop As PropertyInfo, value As Object) As Object
        If value Is Nothing Then Return DBNull.Value

        Dim type = Nullable.GetUnderlyingType(prop.PropertyType)
        If type Is Nothing Then type = prop.PropertyType

        If type.IsEnum Then
            Dim attr = prop.GetCustomAttribute(Of EnumStorageAttribute)()
            If attr IsNot Nothing AndAlso attr.StorageType = EnumStorageType.String Then
                Return value.ToString()
            End If
        End If

        Return value
    End Function

    Friend Shared Function GetIdValue(obj As Object, idColumn As String) As Object
        Dim prop = obj.GetType().GetProperty(idColumn)
        If prop Is Nothing Then
            Throw New ArgumentException($"Property '{idColumn}' not found on type '{obj.GetType().Name}'")
        End If
        Return prop.GetValue(obj)
    End Function

    Friend Shared Function IsChildCollection(prop As PropertyInfo) As Boolean
        Return GetType(IEnumerable).IsAssignableFrom(prop.PropertyType) AndAlso
           prop.PropertyType.IsGenericType AndAlso
           Not prop.PropertyType Is GetType(String)
    End Function

    Friend Shared Function CleanOutIds(Of T)(obj As T, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T
        If Not IdWillAutoIncrement Then
            Return obj
        End If

        Dim typeT = GetType(T)
        Dim newObj = Activator.CreateInstance(Of T)()
        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        For Each prop In props
            If IsGenericList(prop.PropertyType) Then
                Dim list = CType(prop.GetValue(obj), IEnumerable)
                If list Is Nothing Then Continue For

                Dim elementType = prop.PropertyType.GetGenericArguments()(0)
                Dim newList = CType(Activator.CreateInstance(GetType(List(Of )).MakeGenericType(elementType)), IList)

                For Each item In list
                    Dim newItem = Activator.CreateInstance(elementType)
                    Dim itemProps = elementType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

                    For Each itemProp In itemProps
                        If itemProp.Name = Id OrElse itemProp.Name.EndsWith("_" & idColumn) Then
                            ' Leave it blank/null
                            itemProp.SetValue(newItem, Nothing)
                        Else
                            itemProp.SetValue(newItem, itemProp.GetValue(item))
                        End If
                    Next

                    newList.Add(newItem)
                Next

                prop.SetValue(newObj, newList)
            Else
                If prop.Name = idColumn Then
                    ' Null out identity field
                    prop.SetValue(newObj, Nothing)
                Else
                    prop.SetValue(newObj, prop.GetValue(obj))
                End If
            End If
        Next

        Return newObj
    End Function


#End Region

#Region "async"

    Friend Shared Async Function GetChildRecordsAsync(connection As IDbConnection, parentId As Object, parentTable As String, parentIdColumn As String, childType As Type, foreignKey As String, provider As IDbProviderAsync) As Task(Of IList)
        Dim parameterPrefix = provider.GetParameterPrefix()
        Dim childTable = childType.Name
        Dim query = $"SELECT * FROM [{childTable}] WHERE [{foreignKey}] = {parameterPrefix}ParentId"

        Dim list = CType(Activator.CreateInstance(GetType(List(Of )).MakeGenericType(childType)), IList)

        Using cmd = Await provider.CreateCommandAsync(query, connection)
            cmd.Parameters.Add(Await provider.CreateParameterAsync($"{parameterPrefix}ParentId", parentId))

            ' 🔥 Here is the fix:
            Using reader = Await DirectCast(cmd, DbCommand).ExecuteReaderAsync()
                While Await reader.ReadAsync()
                    Dim childObj = Activator.CreateInstance(childType)
                    For Each prop In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                        If ColumnExists(reader, prop.Name) AndAlso Not Await reader.IsDBNullAsync(reader.GetOrdinal(prop.Name)) Then
                            prop.SetValue(childObj, Convert.ChangeType(reader(prop.Name), prop.PropertyType))
                        End If
                    Next
                    list.Add(childObj)
                End While
            End Using
        End Using

        Return list
    End Function

    Friend Shared Async Function CreateOrUpdateTableRecursiveAsync(entityType As Type, idColumn As String, mode As DbPrepMode, connection As DbConnection, transaction As DbTransaction, executedTables As HashSet(Of String)) As Task
        Dim tableName = entityType.Name
        If executedTables.Contains(tableName) Then Return
        executedTables.Add(tableName)

        Dim properties = entityType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        ' Recursively handle child collections first
        For Each prop In properties.Where(Function(p) IsGenericList(p.PropertyType))
            Dim childType = prop.PropertyType.GetGenericArguments()(0)
            Await CreateOrUpdateTableRecursiveAsync(childType, idColumn, mode, connection, transaction, executedTables)
        Next

        If mode = DbPrepMode.Create Then
            Dim dropSql = $"IF OBJECT_ID('[{tableName}]', 'U') IS NOT NULL DROP TABLE [{tableName}];"
            Await ExecuteNonQueryAsync(dropSql, connection, transaction)
        End If

        ' Verify ID column exists
        If Not properties.Any(Function(p) p.Name = idColumn) Then
            Throw New InvalidOperationException($"Type {entityType.Name} must define a property named '{idColumn}' as primary key.")
        End If

        ' Build column definitions
        Dim columnDefs As New List(Of String)
        For Each prop In properties
            If IsGenericList(prop.PropertyType) Then Continue For

            Dim columnName = prop.Name
            Dim dbType = MapToSqlType(prop)

            If columnName = idColumn Then
                columnDefs.Add($"[{columnName}] INT IDENTITY(1,1) PRIMARY KEY")
            ElseIf columnName = "RowVersion" AndAlso prop.PropertyType = GetType(Byte()) Then
                columnDefs.Add("[RowVersion] ROWVERSION")
            Else
                columnDefs.Add($"[{columnName}] {dbType}")
            End If
        Next

        ' Add foreign key columns for child types
        For Each prop In properties
            If IsGenericList(prop.PropertyType) Then
                Dim childType = prop.PropertyType.GetGenericArguments()(0)
                Dim childTable = childType.Name
                Dim fkColumn = $"{tableName}_{idColumn}"
                Dim alterSql = $"IF COL_LENGTH('[{childTable}]', '{fkColumn}') IS NULL BEGIN " &
                           $"ALTER TABLE [{childTable}] ADD [{fkColumn}] INT; " &
                           $"ALTER TABLE [{childTable}] ADD CONSTRAINT FK_{childTable}_{fkColumn} " &
                           $"FOREIGN KEY([{fkColumn}]) REFERENCES [{tableName}]([{idColumn}]) ON DELETE CASCADE; END"
                Await ExecuteNonQueryAsync(alterSql, connection, transaction)
            End If
        Next

        ' Create or update table
        If mode = DbPrepMode.Create Then
            Dim createSql = $"CREATE TABLE [{tableName}] ({String.Join(", ", columnDefs)});"
            Await ExecuteNonQueryAsync(createSql, connection, transaction)
        ElseIf mode = DbPrepMode.Update Then
            ' Get existing columns
            Dim existingColumns As New HashSet(Of String)
            Using cmd = connection.CreateCommand()
                cmd.Transaction = transaction
                cmd.CommandText = $"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tableName}'"
                Using reader = Await cmd.ExecuteReaderAsync()
                    While Await reader.ReadAsync()
                        existingColumns.Add(reader.GetString(0))
                    End While
                End Using
            End Using

            For Each def In columnDefs
                Dim columnName = def.Split(" "c)(0).Trim("["c, "]"c)
                If columnName = idColumn OrElse existingColumns.Contains(columnName) Then Continue For

                Dim alterSql = $"ALTER TABLE [{tableName}] ADD {def};"
                Await ExecuteNonQueryAsync(alterSql, connection, transaction)
            Next
        End If
    End Function
    Friend Shared Async Function HasCascadeDeleteConstraintAsync(tableName As String, connection As IDbConnection, transaction As IDbTransaction, _provider As IDbProvider) As Task(Of Boolean)
        Dim sql As String = "
        SELECT COUNT(*) 
        FROM sys.foreign_keys fk
        JOIN sys.tables parent ON fk.parent_object_id = parent.object_id
        WHERE parent.name = @TableName
          AND fk.delete_referential_action_desc = 'CASCADE'
    "

        Using cmd = _provider.CreateCommand(sql, connection)
            cmd.Transaction = transaction
            cmd.Parameters.Add(_provider.CreateParameter("@TableName", tableName))
            Dim result = Convert.ToInt32(Await CType(cmd, DbCommand).ExecuteScalarAsync())
            Return result > 0
        End Using
    End Function

#End Region

#Region "support async"
    Private Shared Async Function ExecuteNonQueryAsync(sql As String, connection As DbConnection, transaction As DbTransaction) As Task
        Using cmd = connection.CreateCommand()
            cmd.Transaction = transaction
            cmd.CommandText = sql
            Await cmd.ExecuteNonQueryAsync()
        End Using
    End Function

#End Region

#Region "support sync"
    Private Shared ReadOnly WellKnownScalarTypes As HashSet(Of Type) = New HashSet(Of Type) From {
    GetType(Boolean),
    GetType(Byte),
    GetType(SByte),
    GetType(Short),
    GetType(UShort),
    GetType(Integer),
    GetType(UInteger),
    GetType(Long),
    GetType(ULong),
    GetType(Single),
    GetType(Double),
    GetType(Decimal),
    GetType(Char),
    GetType(String),
    GetType(DateTime),
    GetType(DateTimeOffset),
    GetType(TimeSpan),
    GetType(Guid)
}

    Private Shared Function IsScalarType(t As Type) As Boolean
        Return WellKnownScalarTypes.Contains(t) OrElse t.IsEnum
    End Function

    Private Shared Function ShouldRecurseIntoType(type As Type) As Boolean
        If type Is Nothing Then Return False

        ' Unwrap Nullable(Of T)
        If Nullable.GetUnderlyingType(type) IsNot Nothing Then
            type = Nullable.GetUnderlyingType(type)
        End If

        If type.IsPrimitive OrElse type.IsEnum Then Return False

        If type = GetType(String) OrElse
       type = GetType(DateTime) OrElse
       type = GetType(Decimal) OrElse
       type = GetType(Guid) OrElse
       type = GetType(Byte()) Then
            Return False
        End If

        ' Avoid descending into common wrappers
        If type.Namespace IsNot Nothing AndAlso
       (type.Namespace.StartsWith("System") OrElse type.Namespace.StartsWith("Microsoft")) Then
            If Not type.IsClass Then Return False
        End If

        Return True
    End Function

    Private Shared Function BuildCreateTableSql(entityType As Type, idColumn As String) As String
        Dim tableName As String = entityType.Name
        Dim columns As New List(Of String)

        For Each prop In entityType.GetProperties()
            ' Skip collections
            If GetType(System.Collections.IEnumerable).IsAssignableFrom(prop.PropertyType) AndAlso prop.PropertyType IsNot GetType(String) Then
                Continue For
            End If

            Dim columnName = prop.Name
            Dim sqlType = MapToSqlType(prop)

            If columnName.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                columns.Add($"[{columnName}] {sqlType} PRIMARY KEY IDENTITY(1,1)")
            ElseIf columnName.Equals("RowVersion", StringComparison.OrdinalIgnoreCase) Then
                columns.Add($"[{columnName}] ROWVERSION")
            Else
                columns.Add($"[{columnName}] {sqlType}")
            End If
        Next

        Dim columnsSql = String.Join(", ", columns)
        Return $"CREATE TABLE [{tableName}] ({columnsSql});"
    End Function
    Private Shared Function BuildAlterTableSql(entityType As Type, idColumn As String, connection As IDbConnection, transaction As IDbTransaction) As String
        Dim tableName As String = entityType.Name
        Dim existingColumns As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ' Fetch current columns
        Using cmd = connection.CreateCommand()
            cmd.Transaction = transaction
            cmd.CommandText = $"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tableName}'"
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    existingColumns.Add(reader.GetString(0))
                End While
            End Using
        End Using

        Dim alterStatements As New List(Of String)

        For Each prop In entityType.GetProperties()
            ' Skip collections
            If GetType(System.Collections.IEnumerable).IsAssignableFrom(prop.PropertyType) AndAlso prop.PropertyType IsNot GetType(String) Then
                Continue For
            End If

            Dim columnName = prop.Name

            If Not existingColumns.Contains(columnName) Then
                Dim sqlType = MapToSqlType(prop)

                If columnName.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                    Continue For ' Skip ID column (already exists)
                ElseIf columnName.Equals("RowVersion", StringComparison.OrdinalIgnoreCase) Then
                    alterStatements.Add($"ALTER TABLE [{tableName}] ADD [{columnName}] ROWVERSION;")
                Else
                    alterStatements.Add($"ALTER TABLE [{tableName}] ADD [{columnName}] {sqlType};")
                End If
            End If
        Next

        Return String.Join(Environment.NewLine, alterStatements)
    End Function

#End Region


End Class
