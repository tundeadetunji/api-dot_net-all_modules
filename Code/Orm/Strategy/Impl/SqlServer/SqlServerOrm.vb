Imports iNovation.Code.OrmUtils
Imports System.Threading
Friend Class SqlServerOrm
    Implements IOrm

#Region "init"
    Private ReadOnly _provider As IDbProvider
    Public Const Id As String = "Id"

    Private Shared _instance As Lazy(Of SqlServerOrm)

    Private Sub New(connectionString As String)
        _provider = New SqlServerProvider(connectionString)
    End Sub

    Public Shared ReadOnly Property Instance(connectionString As String) As SqlServerOrm
        Get
            If _instance Is Nothing Then
                _instance = New Lazy(Of SqlServerOrm)(Function() New SqlServerOrm(connectionString), LazyThreadSafetyMode.ExecutionAndPublication)
            End If
            Return _instance.Value
        End Get
    End Property
#End Region

#Region "PrepareDatabase"

    Public Sub PrepareDatabase(mode As DbPrepMode, entities As List(Of Type), Optional idColumn As String = Id) Implements IOrm.PrepareDatabase
        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

                Dim executedTables As New HashSet(Of String)

                For Each entityType In entities
                    CreateOrUpdateTableRecursive(entityType, idColumn, mode, connection, transaction, executedTables)
                Next

                transaction.Commit()
            End Using
        End Using
    End Sub
#End Region

#Region "FindById"
    Public Function FindByIdInTable(Of T)(id As Object, tableName As String, Optional idColumn As String = Id) As T Implements IOrm.FindByIdInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim query As String = $"SELECT * FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"

        Dim obj As T = Nothing

        Using connection = _provider.CreateConnection()
            connection.Open()

            Using command = _provider.CreateCommand(query, connection)
                command.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", id))

                Using reader = command.ExecuteReader()
                    If reader.Read() Then
                        obj = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
                                Dim rawValue = reader(prop.Name)
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                                prop.SetValue(obj, safeValue)
                            End If
                        Next
                    End If
                End Using ' reader
            End Using ' command

            ' Now that the reader is closed, we can safely fetch child records
            If obj IsNot Nothing Then
                For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso IsGenericList(p.PropertyType))
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim foreignKey = $"{tableName}_{idColumn}" ' e.g., Person_Id
                    Dim childList = GetChildRecords(connection, id, tableName, idColumn, childType, foreignKey, _provider)
                    prop.SetValue(obj, childList)
                Next
            End If
        End Using ' connection

        Return obj
    End Function

    Public Function FindById(Of T)(id As Object, Optional idColumn As String = Id) As T Implements IOrm.FindById
        Dim typeT = GetType(T)
        Dim tableName As String = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim query As String = $"SELECT * FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"

        Dim obj As T = Nothing

        Using connection = _provider.CreateConnection()
            connection.Open()

            Using command = _provider.CreateCommand(query, connection)
                command.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", id))

                Using reader = command.ExecuteReader()
                    If reader.Read() Then
                        obj = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
                                Dim rawValue = reader(prop.Name)
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                                prop.SetValue(obj, safeValue)
                            End If
                        Next
                    End If
                End Using ' reader
            End Using ' command

            ' Now that the reader is closed, we can safely fetch child records
            If obj IsNot Nothing Then
                For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso IsGenericList(p.PropertyType))
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim foreignKey = $"{tableName}_{idColumn}" ' e.g., Person_Id
                    Dim childList = GetChildRecords(connection, id, tableName, idColumn, childType, foreignKey, _provider)
                    prop.SetValue(obj, childList)
                Next
            End If
        End Using ' connection

        Return obj
    End Function

#End Region

#Region "Update"
    Public Function Update(Of T)(obj As T, Optional idColumn As String = Id) As T Implements IOrm.Update
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

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

                ' Update parent row
                Dim setClauses = New List(Of String)
                Dim parameters = New List(Of IDataParameter)

                For Each prop In properties
                    If IsGenericList(prop.PropertyType) OrElse prop.Name = idColumn OrElse prop.Name = "RowVersion" Then Continue For

                    Dim paramName = $"{parameterPrefix}{prop.Name}"
                    setClauses.Add($"[{prop.Name}] = {paramName}")
                    parameters.Add(_provider.CreateParameter(paramName, prop.GetValue(obj)))
                Next

                parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", idValue))

                ' Add RowVersion condition if available
                Dim whereClause = $"[{idColumn}] = {parameterPrefix}id"
                If hasRowVersion Then
                    whereClause &= $" AND [RowVersion] = {parameterPrefix}RowVersion"
                    parameters.Add(_provider.CreateParameter($"{parameterPrefix}RowVersion", rowVersionProp.GetValue(obj)))
                End If

                Dim sql = $"UPDATE [{tableName}] SET {String.Join(", ", setClauses)} WHERE {whereClause}"

                Using command = _provider.CreateCommand(sql, connection)
                    command.Transaction = transaction
                    For Each param In parameters
                        command.Parameters.Add(param)
                    Next
                    Dim affected = command.ExecuteNonQuery()
                    If affected <> 1 Then
                        Throw New DBConcurrencyException("The record was updated by another user or no longer exists.")
                    End If
                End Using

                ' Child object handling (same as before — can enhance later if needed)
                ' ...

                transaction.Commit()
            End Using
        End Using

        Return obj
    End Function

    Public Function UpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = Id) As T Implements IOrm.UpdateInTable
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

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

                ' Update parent row
                Dim setClauses = New List(Of String)
                Dim parameters = New List(Of IDataParameter)

                For Each prop In properties
                    If IsGenericList(prop.PropertyType) OrElse prop.Name = idColumn OrElse prop.Name = "RowVersion" Then Continue For

                    Dim paramName = $"{parameterPrefix}{prop.Name}"
                    setClauses.Add($"[{prop.Name}] = {paramName}")
                    parameters.Add(_provider.CreateParameter(paramName, prop.GetValue(obj)))
                Next

                parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", idValue))

                ' Add RowVersion condition if available
                Dim whereClause = $"[{idColumn}] = {parameterPrefix}id"
                If hasRowVersion Then
                    whereClause &= $" AND [RowVersion] = {parameterPrefix}RowVersion"
                    parameters.Add(_provider.CreateParameter($"{parameterPrefix}RowVersion", rowVersionProp.GetValue(obj)))
                End If

                Dim sql = $"UPDATE [{tableName}] SET {String.Join(", ", setClauses)} WHERE {whereClause}"

                Using command = _provider.CreateCommand(sql, connection)
                    command.Transaction = transaction
                    For Each param In parameters
                        command.Parameters.Add(param)
                    Next
                    Dim affected = command.ExecuteNonQuery()
                    If affected <> 1 Then
                        Throw New DBConcurrencyException("The record was updated by another user or no longer exists.")
                    End If
                End Using

                ' Child object handling (same as before — can enhance later if needed)
                ' ...

                transaction.Commit()
            End Using
        End Using

        Return obj
    End Function

#End Region

#Region "Create"
    Public Function Create(Of T)(obj As T, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T Implements IOrm.Create
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim properties = typeT.GetProperties().
        Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        Dim idProp = properties.FirstOrDefault(Function(p) p.Name = idColumn)
        If idProp Is Nothing Then
            Throw New InvalidOperationException($"The specified id column '{idColumn}' was not found on type '{typeT.Name}'.")
        End If

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

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
                    parameters.Add(_provider.CreateParameter(paramName, dbValue))
                Next

                Dim sql = $"INSERT INTO [{tableName}] ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", insertVals)})"
                If IdWillAutoIncrement Then
                    sql += "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                End If

                Using cmd = _provider.CreateCommand(sql, connection)
                    cmd.Transaction = transaction
                    For Each param In parameters
                        cmd.Parameters.Add(param)
                    Next

                    If IdWillAutoIncrement Then
                        Dim newId = Convert.ToInt64(cmd.ExecuteScalar())
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
                            cParams.Add(_provider.CreateParameter($"{parameterPrefix}{cpName}", dbValue))
                        Next

                        Dim childSql = $"INSERT INTO [{childTable}] ({String.Join(", ", cCols)}) VALUES ({String.Join(", ", cVals)}); SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"

                        Using cmd = _provider.CreateCommand(childSql, connection)
                            cmd.Transaction = transaction
                            For Each p In cParams
                                cmd.Parameters.Add(p)
                            Next

                            Dim newChildId = Convert.ToInt64(cmd.ExecuteScalar())
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
    Public Function CreateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T Implements IOrm.CreateInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()

        Dim properties = typeT.GetProperties().
        Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        Dim idProp = properties.FirstOrDefault(Function(p) p.Name = idColumn)
        If idProp Is Nothing Then
            Throw New InvalidOperationException($"The specified id column '{idColumn}' was not found on type '{typeT.Name}'.")
        End If

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

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
                    parameters.Add(_provider.CreateParameter(paramName, dbValue))
                Next

                Dim sql = $"INSERT INTO [{tableName}] ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", insertVals)})"
                If IdWillAutoIncrement Then
                    sql += "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                End If

                Using cmd = _provider.CreateCommand(sql, connection)
                    cmd.Transaction = transaction
                    For Each param In parameters
                        cmd.Parameters.Add(param)
                    Next

                    If IdWillAutoIncrement Then
                        Dim newId = Convert.ToInt64(cmd.ExecuteScalar())
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
                            cParams.Add(_provider.CreateParameter($"{parameterPrefix}{cpName}", dbValue))
                        Next

                        Dim childSql = $"INSERT INTO [{childTable}] ({String.Join(", ", cCols)}) VALUES ({String.Join(", ", cVals)}); SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"

                        Using cmd = _provider.CreateCommand(childSql, connection)
                            cmd.Transaction = transaction
                            For Each p In cParams
                                cmd.Parameters.Add(p)
                            Next

                            Dim newChildId = Convert.ToInt64(cmd.ExecuteScalar())
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
    Public Function Create(Of T)(objs As List(Of T), Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As List(Of T) Implements IOrm.Create
        If objs Is Nothing OrElse objs.Count = 0 Then Return objs

        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim tableName = typeT.Name
        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()
                For Each obj In objs
                    ' Insert parent
                    Dim columns = New List(Of String)
                    Dim values = New List(Of String)
                    Dim parameters = New List(Of IDataParameter)

                    For Each prop In props
                        If IsGenericList(prop.PropertyType) OrElse (prop.Name = idColumn AndAlso IdWillAutoIncrement) Then Continue For

                        columns.Add($"[{prop.Name}]")
                        values.Add($"{parameterPrefix}{prop.Name}")

                        Dim rawValue = prop.GetValue(obj)
                        Dim dbValue = GetEnumDbValue(prop, rawValue)
                        parameters.Add(_provider.CreateParameter($"{parameterPrefix}{prop.Name}", dbValue))
                    Next

                    Dim insertSql = $"INSERT INTO [{tableName}] ({String.Join(", ", columns)}) VALUES ({String.Join(", ", values)})"
                    If IdWillAutoIncrement Then
                        insertSql &= "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                    End If

                    Using command = _provider.CreateCommand(insertSql, connection)
                        command.Transaction = transaction
                        For Each param In parameters
                            command.Parameters.Add(param)
                        Next

                        If IdWillAutoIncrement Then
                            Dim newId = Convert.ToInt64(command.ExecuteScalar())
                            Dim idProp = props.First(Function(p) p.Name = idColumn)
                            idProp.SetValue(obj, Convert.ChangeType(newId, idProp.PropertyType))
                        Else
                            command.ExecuteNonQuery()
                        End If
                    End Using

                    ' Insert child collections
                    For Each prop In props.Where(Function(p) IsGenericList(p.PropertyType))
                        Dim childList = CType(prop.GetValue(obj), IEnumerable)
                        If childList Is Nothing Then Continue For

                        Dim childType = prop.PropertyType.GetGenericArguments()(0)
                        Dim childTable = childType.Name
                        Dim childProps = childType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite AndAlso Not IsGenericList(p.PropertyType)).ToList()

                        Dim fkColumn = $"{tableName}_{idColumn}"
                        Dim parentId = props.First(Function(p) p.Name = idColumn).GetValue(obj)

                        For Each child In childList
                            Dim insertCols = New List(Of String)
                            Dim insertVals = New List(Of String)
                            Dim insertParams = New List(Of IDataParameter)

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
                                insertCols.Add($"[{cp.Name}]")
                                insertVals.Add($"{parameterPrefix}{cp.Name}")
                                insertParams.Add(_provider.CreateParameter($"{parameterPrefix}{cp.Name}", dbValue))
                            Next

                            Dim insertChildSql = $"INSERT INTO [{childTable}] ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", insertVals)})"
                            If IdWillAutoIncrement Then
                                insertChildSql &= "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                            End If

                            Using cmd = _provider.CreateCommand(insertChildSql, connection)
                                cmd.Transaction = transaction
                                For Each p In insertParams
                                    cmd.Parameters.Add(p)
                                Next

                                If IdWillAutoIncrement Then
                                    Dim newChildId = Convert.ToInt64(cmd.ExecuteScalar())
                                    Dim idProp = childProps.FirstOrDefault(Function(p) p.Name = idColumn)
                                    If idProp IsNot Nothing AndAlso idProp.CanWrite Then
                                        idProp.SetValue(child, Convert.ChangeType(newChildId, idProp.PropertyType))
                                    End If
                                Else
                                    cmd.ExecuteNonQuery()
                                End If
                            End Using
                        Next
                    Next
                Next

                transaction.Commit()
            End Using
        End Using

        Return objs
    End Function
    Public Function CreateInTable(Of T)(objs As List(Of T), tableName As String, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As List(Of T) Implements IOrm.CreateInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        If objs Is Nothing OrElse objs.Count = 0 Then Return objs

        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()
                For Each obj In objs
                    ' Insert parent
                    Dim columns = New List(Of String)
                    Dim values = New List(Of String)
                    Dim parameters = New List(Of IDataParameter)

                    For Each prop In props
                        If IsGenericList(prop.PropertyType) OrElse (prop.Name = idColumn AndAlso IdWillAutoIncrement) Then Continue For

                        columns.Add($"[{prop.Name}]")
                        values.Add($"{parameterPrefix}{prop.Name}")

                        Dim rawValue = prop.GetValue(obj)
                        Dim dbValue = GetEnumDbValue(prop, rawValue)
                        parameters.Add(_provider.CreateParameter($"{parameterPrefix}{prop.Name}", dbValue))
                    Next

                    Dim insertSql = $"INSERT INTO [{tableName}] ({String.Join(", ", columns)}) VALUES ({String.Join(", ", values)})"
                    If IdWillAutoIncrement Then
                        insertSql &= "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                    End If

                    Using command = _provider.CreateCommand(insertSql, connection)
                        command.Transaction = transaction
                        For Each param In parameters
                            command.Parameters.Add(param)
                        Next

                        If IdWillAutoIncrement Then
                            Dim newId = Convert.ToInt64(command.ExecuteScalar())
                            Dim idProp = props.First(Function(p) p.Name = idColumn)
                            idProp.SetValue(obj, Convert.ChangeType(newId, idProp.PropertyType))
                        Else
                            command.ExecuteNonQuery()
                        End If
                    End Using

                    ' Insert child collections
                    For Each prop In props.Where(Function(p) IsGenericList(p.PropertyType))
                        Dim childList = CType(prop.GetValue(obj), IEnumerable)
                        If childList Is Nothing Then Continue For

                        Dim childType = prop.PropertyType.GetGenericArguments()(0)
                        Dim childTable = childType.Name
                        Dim childProps = childType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite AndAlso Not IsGenericList(p.PropertyType)).ToList()

                        Dim fkColumn = $"{tableName}_{idColumn}"
                        Dim parentId = props.First(Function(p) p.Name = idColumn).GetValue(obj)

                        For Each child In childList
                            Dim insertCols = New List(Of String)
                            Dim insertVals = New List(Of String)
                            Dim insertParams = New List(Of IDataParameter)

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
                                insertCols.Add($"[{cp.Name}]")
                                insertVals.Add($"{parameterPrefix}{cp.Name}")
                                insertParams.Add(_provider.CreateParameter($"{parameterPrefix}{cp.Name}", dbValue))
                            Next

                            Dim insertChildSql = $"INSERT INTO [{childTable}] ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", insertVals)})"
                            If IdWillAutoIncrement Then
                                insertChildSql &= "; SELECT CAST(SCOPE_IDENTITY() AS BIGINT);"
                            End If

                            Using cmd = _provider.CreateCommand(insertChildSql, connection)
                                cmd.Transaction = transaction
                                For Each p In insertParams
                                    cmd.Parameters.Add(p)
                                Next

                                If IdWillAutoIncrement Then
                                    Dim newChildId = Convert.ToInt64(cmd.ExecuteScalar())
                                    Dim idProp = childProps.FirstOrDefault(Function(p) p.Name = idColumn)
                                    If idProp IsNot Nothing AndAlso idProp.CanWrite Then
                                        idProp.SetValue(child, Convert.ChangeType(newChildId, idProp.PropertyType))
                                    End If
                                Else
                                    cmd.ExecuteNonQuery()
                                End If
                            End Using
                        Next
                    Next
                Next

                transaction.Commit()
            End Using
        End Using

        Return objs
    End Function

#End Region

#Region "CreateOrUpdate"
    Public Function CreateOrUpdate(Of T)(obj As T, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T Implements IOrm.CreateOrUpdate
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim idValue = GetIdValue(obj, idColumn)

        If ExistsByIdInTable(idValue, tableName, idColumn) Then
            Return Update(Of T)(obj, idColumn)
        End If

        Return Create(Of T)(CleanOutIds(Of T)(obj), idColumn, IdWillAutoIncrement)

    End Function

    Public Function CreateOrUpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T Implements IOrm.CreateOrUpdateInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim idValue = GetIdValue(obj, idColumn)

        If ExistsByIdInTable(idValue, tableName, idColumn) Then
            Return UpdateInTable(Of T)(obj, tableName, idColumn)
        End If

        Return CreateInTable(Of T)(CleanOutIds(Of T)(obj), tableName, idColumn, IdWillAutoIncrement)

    End Function

#End Region

#Region "Delete"

    Public Sub DeleteWhereId(Of T)(id As Object, Optional idColumn As String = "Id", Optional cascade As Boolean = True) Implements IOrm.DeleteWhereId
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

        ' Load the current RowVersion value (if present) so we can use it in the DELETE
        Dim rowVersionValue As Byte() = Nothing
        If rowVersionProp IsNot Nothing Then
            Dim findSql = $"SELECT [RowVersion] FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"
            Using connection = _provider.CreateConnection()
                connection.Open()
                Using cmd = _provider.CreateCommand(findSql, connection)
                    cmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", id))
                    Dim result = cmd.ExecuteScalar()
                    If result Is DBNull.Value OrElse result Is Nothing Then
                        Throw New InvalidOperationException("Record not found or missing RowVersion.")
                    End If
                    rowVersionValue = CType(result, Byte())
                End Using
            End Using
        End If

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

                ' Cascade delete children first if needed
                If cascade Then
                    For Each prop In props.Where(Function(p) IsGenericList(p.PropertyType))
                        Dim childType = prop.PropertyType.GetGenericArguments()(0)
                        Dim childTable = childType.Name
                        Dim fkColumn = $"{tableName}_{idColumn}"
                        Dim deleteChildSql = $"DELETE FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}id"

                        Using childCmd = _provider.CreateCommand(deleteChildSql, connection)
                            childCmd.Transaction = transaction
                            childCmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", id))
                            childCmd.ExecuteNonQuery()
                        End Using
                    Next
                End If

                ' Delete parent with optional RowVersion
                Dim deleteSql = $"DELETE FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"
                If rowVersionProp IsNot Nothing Then
                    deleteSql &= $" AND [RowVersion] = {parameterPrefix}rowversion"
                End If

                Using cmd = _provider.CreateCommand(deleteSql, connection)
                    cmd.Transaction = transaction
                    cmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", id))

                    If rowVersionProp IsNot Nothing Then
                        Dim param = _provider.CreateParameter($"{parameterPrefix}rowversion", rowVersionValue)
                        param.DbType = DbType.Binary
                        cmd.Parameters.Add(param)
                    End If

                    Dim affected = cmd.ExecuteNonQuery()
                    If affected <> 1 Then
                        Throw New DBConcurrencyException("Delete failed due to concurrent update. Record was modified or deleted by another process.")
                    End If
                End Using

                transaction.Commit()
            End Using
        End Using
    End Sub

    Public Sub DeleteWhereIdInTable(Of T)(id As Object, tableName As String, Optional idColumn As String = Id, Optional cascade As Boolean = True) Implements IOrm.DeleteWhereIdInTable
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

        ' Load the current RowVersion value (if present) so we can use it in the DELETE
        Dim rowVersionValue As Byte() = Nothing
        If rowVersionProp IsNot Nothing Then
            Dim findSql = $"SELECT [RowVersion] FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"
            Using connection = _provider.CreateConnection()
                connection.Open()
                Using cmd = _provider.CreateCommand(findSql, connection)
                    cmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", id))
                    Dim result = cmd.ExecuteScalar()
                    If result Is DBNull.Value OrElse result Is Nothing Then
                        Throw New InvalidOperationException("Record not found or missing RowVersion.")
                    End If
                    rowVersionValue = CType(result, Byte())
                End Using
            End Using
        End If

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

                ' Cascade delete children first if needed
                If cascade Then
                    For Each prop In props.Where(Function(p) IsGenericList(p.PropertyType))
                        Dim childType = prop.PropertyType.GetGenericArguments()(0)
                        Dim childTable = childType.Name
                        Dim fkColumn = $"{tableName}_{idColumn}"
                        Dim deleteChildSql = $"DELETE FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}id"

                        Using childCmd = _provider.CreateCommand(deleteChildSql, connection)
                            childCmd.Transaction = transaction
                            childCmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", id))
                            childCmd.ExecuteNonQuery()
                        End Using
                    Next
                End If

                ' Delete parent with optional RowVersion
                Dim deleteSql = $"DELETE FROM [{tableName}] WHERE [{idColumn}] = {parameterPrefix}id"
                If rowVersionProp IsNot Nothing Then
                    deleteSql &= $" AND [RowVersion] = {parameterPrefix}rowversion"
                End If

                Using cmd = _provider.CreateCommand(deleteSql, connection)
                    cmd.Transaction = transaction
                    cmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", id))

                    If rowVersionProp IsNot Nothing Then
                        Dim param = _provider.CreateParameter($"{parameterPrefix}rowversion", rowVersionValue)
                        param.DbType = DbType.Binary
                        cmd.Parameters.Add(param)
                    End If

                    Dim affected = cmd.ExecuteNonQuery()
                    If affected <> 1 Then
                        Throw New DBConcurrencyException("Delete failed due to concurrent update. Record was modified or deleted by another process.")
                    End If
                End Using

                transaction.Commit()
            End Using
        End Using
    End Sub

    Public Sub DeleteAll(Of T)(Optional cascade As Boolean = True) Implements IOrm.DeleteAll
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()
                Try
                    ' Enforce decision if DB is configured to cascade deletes
                    If Not cascade AndAlso HasCascadeDeleteConstraint(tableName, connection, transaction, _provider) Then
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
                            Using childCmd = _provider.CreateCommand(childSql, connection)
                                childCmd.Transaction = transaction
                                childCmd.ExecuteNonQuery()
                            End Using
                        Next
                    End If

                    Try
                        Dim deleteSql = $"DELETE FROM [{tableName}]"
                        Using fallbackCmd = _provider.CreateCommand(deleteSql, connection)
                            fallbackCmd.Transaction = transaction
                            fallbackCmd.ExecuteNonQuery()
                        End Using
                    Catch ex As Exception
                    End Try

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub
    Public Sub DeleteAllInTable(Of T)(tableName As String, Optional cascade As Boolean = True) Implements IOrm.DeleteAllInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()
                Try
                    ' Enforce decision if DB is configured to cascade deletes
                    If Not cascade AndAlso HasCascadeDeleteConstraint(tableName, connection, transaction, _provider) Then
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
                            Using childCmd = _provider.CreateCommand(childSql, connection)
                                childCmd.Transaction = transaction
                                childCmd.ExecuteNonQuery()
                            End Using
                        Next
                    End If

                    Try
                        Dim deleteSql = $"DELETE FROM [{tableName}]"
                        Using fallbackCmd = _provider.CreateCommand(deleteSql, connection)
                            fallbackCmd.Transaction = transaction
                            fallbackCmd.ExecuteNonQuery()
                        End Using
                    Catch ex As Exception
                    End Try

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub


#End Region

#Region "Exists"

    Public Function ExistsById(Of T)(id As Object, Optional idColumn As String = Id) As Boolean Implements IOrm.ExistsById
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim paramName = $"{parameterPrefix}{idColumn}"

        Dim query = $"SELECT COUNT(1) FROM [{tableName}] WHERE [{idColumn}] = {paramName}"

        Using connection = _provider.CreateConnection()
            Using command = _provider.CreateCommand(query, connection)
                command.Parameters.Add(_provider.CreateParameter(paramName, id))
                connection.Open()
                Dim count = Convert.ToInt32(command.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    Public Function ExistsByIdInTable(id As Object, tableName As String, Optional idColumn As String = Id) As Boolean Implements IOrm.ExistsByIdInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim paramName = $"{parameterPrefix}{idColumn}"

        Dim query = $"SELECT COUNT(1) FROM [{tableName}] WHERE [{idColumn}] = {paramName}"

        Using connection = _provider.CreateConnection()
            Using command = _provider.CreateCommand(query, connection)
                command.Parameters.Add(_provider.CreateParameter(paramName, id))
                connection.Open()
                Dim count = Convert.ToInt32(command.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

#End Region

#Region "FindAll"

    Public Function FindAll(Of T)(Optional idColumn As String = "Id", Optional ascending As Boolean = True) As List(Of T) Implements IOrm.FindAll
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim results As New List(Of T)

        Using connection = _provider.CreateConnection()
            connection.Open()

            ' Fetch all parent rows
            Dim sql = $"SELECT * FROM [{tableName}]"
            Using command = _provider.CreateCommand(sql, connection)
                Using reader = command.ExecuteReader()
                    While reader.Read()
                        Dim obj = Activator.CreateInstance(Of T)()
                        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

                        For Each prop In props
                            If IsGenericList(prop.PropertyType) Then Continue For ' skip child collections here
                            If Not ColumnExists(reader, prop.Name) Then Continue For

                            Dim val = reader(prop.Name)
                            If val IsNot DBNull.Value Then
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, val)
                                prop.SetValue(obj, safeValue)
                            End If
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
                    Dim childProps = childType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

                    Dim fkColumn = $"{tableName}_{idColumn}"
                    Dim parentId = typeT.GetProperty(idColumn).GetValue(obj)

                    Dim childList = Activator.CreateInstance(GetType(List(Of )).MakeGenericType(childType))

                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}id"
                    Using cmd = _provider.CreateCommand(childSql, connection)
                        cmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", parentId))

                        Using reader = cmd.ExecuteReader()
                            While reader.Read()
                                Dim childObj = Activator.CreateInstance(childType)

                                For Each cp In childProps
                                    If Not ColumnExists(reader, cp.Name) Then Continue For

                                    Dim val = reader(cp.Name)
                                    If val IsNot DBNull.Value Then
                                        cp.SetValue(childObj, Convert.ChangeType(val, cp.PropertyType))
                                    End If
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
    Public Function FindAllInTable(Of T)(tableName As String, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As List(Of T) Implements IOrm.FindAllInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim results As New List(Of T)

        Using connection = _provider.CreateConnection()
            connection.Open()

            ' Fetch all parent rows
            Dim sql = $"SELECT * FROM [{tableName}]"
            Using command = _provider.CreateCommand(sql, connection)
                Using reader = command.ExecuteReader()
                    While reader.Read()
                        Dim obj = Activator.CreateInstance(Of T)()
                        Dim props = typeT.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

                        For Each prop In props
                            If IsGenericList(prop.PropertyType) Then Continue For ' skip child collections here
                            If Not ColumnExists(reader, prop.Name) Then Continue For

                            Dim val = reader(prop.Name)
                            If val IsNot DBNull.Value Then
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, val)
                                prop.SetValue(obj, safeValue)
                            End If
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
                    Dim childProps = childType.GetProperties().Where(Function(p) p.CanRead AndAlso p.CanWrite).ToList()

                    Dim fkColumn = $"{tableName}_{idColumn}"
                    Dim parentId = typeT.GetProperty(idColumn).GetValue(obj)

                    Dim childList = Activator.CreateInstance(GetType(List(Of )).MakeGenericType(childType))

                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{fkColumn}] = {parameterPrefix}id"
                    Using cmd = _provider.CreateCommand(childSql, connection)
                        cmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}id", parentId))

                        Using reader = cmd.ExecuteReader()
                            While reader.Read()
                                Dim childObj = Activator.CreateInstance(childType)

                                For Each cp In childProps
                                    If Not ColumnExists(reader, cp.Name) Then Continue For

                                    Dim val = reader(cp.Name)
                                    If val IsNot DBNull.Value Then
                                        cp.SetValue(childObj, Convert.ChangeType(val, cp.PropertyType))
                                    End If
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

#End Region

#Region "FindAllPaged"

    Public Function FindAllPaged(Of T)(pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = Id, Optional ascending As Boolean = True) As Page(Of T) Implements IOrm.FindAllPaged
        If pageNumber < 1 Then Throw New ArgumentOutOfRangeException(NameOf(pageNumber))
        If maxPerPage < 1 Then Throw New ArgumentOutOfRangeException(NameOf(maxPerPage))

        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim offset = (pageNumber - 1) * maxPerPage
        Dim totalSql = $"SELECT COUNT(1) FROM [{tableName}]"
        Dim totalRecords As Long
        Dim records = New List(Of T)

        Using connection = _provider.CreateConnection()
            connection.Open()

            Using countCmd = _provider.CreateCommand(totalSql, connection)
                totalRecords = Convert.ToInt64(countCmd.ExecuteScalar())
            End Using

            If totalRecords = 0 Then
                Return New Page(Of T)(New List(Of T), pageNumber, maxPerPage, 0, 0)
            End If

            Dim sql = $"
            SELECT * FROM (
                SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS RowNum
                FROM [{tableName}]
            ) AS Paged
            WHERE RowNum BETWEEN {offset + 1} AND {offset + maxPerPage}
        "

            Using cmd = _provider.CreateCommand(sql, connection)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim obj = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
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
                    Using childCmd = _provider.CreateCommand(childSql, connection)
                        childCmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}ParentId", parentId))
                        Using reader = childCmd.ExecuteReader()
                            Dim childList = CType(Activator.CreateInstance(prop.PropertyType), IList)
                            While reader.Read()
                                Dim childObj = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If ColumnExists(reader, cp.Name) AndAlso Not reader.IsDBNull(reader.GetOrdinal(cp.Name)) Then
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
    Public Function FindAllPagedInTable(Of T)(tableName As String, pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = Id, Optional ascending As Boolean = True) As Page(Of T) Implements IOrm.FindAllPagedInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        If pageNumber < 1 Then Throw New ArgumentOutOfRangeException(NameOf(pageNumber))
        If maxPerPage < 1 Then Throw New ArgumentOutOfRangeException(NameOf(maxPerPage))

        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim offset = (pageNumber - 1) * maxPerPage
        Dim totalSql = $"SELECT COUNT(1) FROM [{tableName}]"
        Dim totalRecords As Long
        Dim records = New List(Of T)

        Using connection = _provider.CreateConnection()
            connection.Open()

            Using countCmd = _provider.CreateCommand(totalSql, connection)
                totalRecords = Convert.ToInt64(countCmd.ExecuteScalar())
            End Using

            If totalRecords = 0 Then
                Return New Page(Of T)(New List(Of T), pageNumber, maxPerPage, 0, 0)
            End If

            Dim sql = $"
            SELECT * FROM (
                SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS RowNum
                FROM [{tableName}]
            ) AS Paged
            WHERE RowNum BETWEEN {offset + 1} AND {offset + maxPerPage}
        "

            Using cmd = _provider.CreateCommand(sql, connection)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim obj = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If ColumnExists(reader, prop.Name) AndAlso Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
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
                    Using childCmd = _provider.CreateCommand(childSql, connection)
                        childCmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}ParentId", parentId))
                        Using reader = childCmd.ExecuteReader()
                            Dim childList = CType(Activator.CreateInstance(prop.PropertyType), IList)
                            While reader.Read()
                                Dim childObj = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If ColumnExists(reader, cp.Name) AndAlso Not reader.IsDBNull(reader.GetOrdinal(cp.Name)) Then
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

#End Region

#Region "Count"
    Public Function Count(Of T)() As Long Implements IOrm.Count
        Dim tableName As String = GetType(T).Name
        Dim query As String = $"SELECT COUNT(*) FROM [{tableName}]"

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using command = _provider.CreateCommand(query, connection)
                Dim result = command.ExecuteScalar()
                Return Convert.ToInt64(result)
            End Using
        End Using
    End Function
    Public Function CountInTable(Of T)(tableName As String) As Long Implements IOrm.CountInTable
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")

        Dim query As String = $"SELECT COUNT(*) FROM [{tableName}]"

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using command = _provider.CreateCommand(query, connection)
                Dim result = command.ExecuteScalar()
                Return Convert.ToInt64(result)
            End Using
        End Using
    End Function

#End Region

#Region "FindBy"

    Public Function FindBy(Of T)(conditions As List(Of Condition)) As List(Of T) Implements IOrm.FindBy

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
                For Each whatever In values
                    Dim subParam = $"{paramName}_{idx}"
                    subParams.Add(subParam)
                    parameters.Add(_provider.CreateParameter(subParam, whatever))
                    idx += 1
                Next
                clause = $"[{c.Column}] IN ({String.Join(", ", subParams)})"
            Else
                clause = $"[{c.Column}] {sqlOperator} {paramName}"
                parameters.Add(_provider.CreateParameter(paramName, c.Value))
            End If

            whereClauses.Add(clause)
        Next

        Dim whereSql = If(whereClauses.Count > 0, "WHERE " & String.Join(" AND ", whereClauses), "")
        Dim sql = $"SELECT * FROM [{tableName}] {whereSql}"

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using command = _provider.CreateCommand(sql, connection)
                For Each param In parameters
                    command.Parameters.Add(param)
                Next

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        Dim instance = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
                                Dim rawValue = reader(prop.Name)
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                                prop.SetValue(instance, safeValue)
                            End If
                        Next
                        results.Add(instance)
                    End While
                End Using
            End Using

            For Each parent In results
                For Each prop In typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType))
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim childFk = $"{tableName}_Id"
                    Dim parentId = typeT.GetProperty("Id")?.GetValue(parent)

                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{childFk}] = {parameterPrefix}parentId"
                    Using childCmd = _provider.CreateCommand(childSql, connection)
                        childCmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}parentId", parentId))
                        Using childReader = childCmd.ExecuteReader()
                            Dim childList = Activator.CreateInstance(prop.PropertyType)
                            While childReader.Read()
                                Dim childInstance = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If Not childReader.IsDBNull(childReader.GetOrdinal(cp.Name)) Then
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
        End Using

        Return results
    End Function

    Public Function FindByInTable(Of T)(conditions As List(Of Condition), tableName As String) As List(Of T) Implements IOrm.FindByInTable


        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If


        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")

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
                For Each whatever In values
                    Dim subParam = $"{paramName}_{idx}"
                    subParams.Add(subParam)
                    parameters.Add(_provider.CreateParameter(subParam, whatever))
                    idx += 1
                Next
                clause = $"[{c.Column}] IN ({String.Join(", ", subParams)})"
            Else
                clause = $"[{c.Column}] {sqlOperator} {paramName}"
                parameters.Add(_provider.CreateParameter(paramName, c.Value))
            End If

            whereClauses.Add(clause)
        Next

        Dim whereSql = If(whereClauses.Count > 0, "WHERE " & String.Join(" AND ", whereClauses), "")
        Dim sql = $"SELECT * FROM [{tableName}] {whereSql}"

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using command = _provider.CreateCommand(sql, connection)
                For Each param In parameters
                    command.Parameters.Add(param)
                Next

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        Dim instance = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
                                Dim rawValue = reader(prop.Name)
                                Dim safeValue = GetSafeEnumValue(prop.PropertyType, rawValue)
                                prop.SetValue(instance, safeValue)
                            End If
                        Next
                        results.Add(instance)
                    End While
                End Using
            End Using

            For Each parent In results
                For Each prop In typeT.GetProperties().Where(Function(p) IsGenericList(p.PropertyType))
                    Dim childType = prop.PropertyType.GetGenericArguments()(0)
                    Dim childTable = childType.Name
                    Dim childFk = $"{tableName}_Id"
                    Dim parentId = typeT.GetProperty("Id")?.GetValue(parent)

                    Dim childSql = $"SELECT * FROM [{childTable}] WHERE [{childFk}] = {parameterPrefix}parentId"
                    Using childCmd = _provider.CreateCommand(childSql, connection)
                        childCmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}parentId", parentId))
                        Using childReader = childCmd.ExecuteReader()
                            Dim childList = Activator.CreateInstance(prop.PropertyType)
                            While childReader.Read()
                                Dim childInstance = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If Not childReader.IsDBNull(childReader.GetOrdinal(cp.Name)) Then
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
        End Using

        Return results
    End Function

#End Region

#Region "FindByPaged"
    Public Function FindByPagedInTable(Of T)(conditions As List(Of Condition), tableName As String, pageNumber As Integer, maxPerPage As Integer) As Page(Of T) Implements IOrm.FindByPagedInTable


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
                For Each whatever In values
                    Dim subParam = $"{paramName}_{idx}"
                    subParams.Add(subParam)
                    parameters.Add(_provider.CreateParameter(subParam, whatever))
                    idx += 1
                Next
                clause = $"[{c.Column}] IN ({String.Join(", ", subParams)})"
            Else
                clause = $"[{c.Column}] {sqlOperator} {paramName}"
                parameters.Add(_provider.CreateParameter(paramName, c.Value))
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

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using command = _provider.CreateCommand(sql, connection)
                For Each param In parameters
                    command.Parameters.Add(param)
                Next

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        Dim instance = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
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
                    Using childCmd = _provider.CreateCommand(childSql, connection)
                        childCmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}parentId", parentId))
                        Using childReader = childCmd.ExecuteReader()
                            Dim childList = Activator.CreateInstance(prop.PropertyType)
                            While childReader.Read()
                                Dim childInstance = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If Not childReader.IsDBNull(childReader.GetOrdinal(cp.Name)) Then
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

            ' Get total count
            Dim countSql = $"SELECT COUNT(1) FROM [{tableName}] {whereSql}"
            Dim totalCount As Long
            Using countCmd = _provider.CreateCommand(countSql, connection)
                For Each param In parameters
                    countCmd.Parameters.Add(_provider.CloneParameter(param))
                Next
                totalCount = Convert.ToInt64(countCmd.ExecuteScalar())
            End Using

            Dim pageCount = Math.Ceiling(totalCount / maxPerPage)

            Return New Page(Of T)(results, pageNumber, maxPerPage, totalCount, CInt(pageCount))
        End Using
    End Function

    Public Function FindByPaged(Of T)(conditions As List(Of Condition), pageNumber As Integer, maxPerPage As Integer) As Page(Of T) Implements IOrm.FindByPaged

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
                For Each whatever In values
                    Dim subParam = $"{paramName}_{idx}"
                    subParams.Add(subParam)
                    parameters.Add(_provider.CreateParameter(subParam, whatever))
                    idx += 1
                Next
                clause = $"[{c.Column}] IN ({String.Join(", ", subParams)})"
            Else
                clause = $"[{c.Column}] {sqlOperator} {paramName}"
                parameters.Add(_provider.CreateParameter(paramName, c.Value))
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

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using command = _provider.CreateCommand(sql, connection)
                For Each param In parameters
                    command.Parameters.Add(param)
                Next

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        Dim instance = Activator.CreateInstance(Of T)()
                        For Each prop In typeT.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                            If Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
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
                    Using childCmd = _provider.CreateCommand(childSql, connection)
                        childCmd.Parameters.Add(_provider.CreateParameter($"{parameterPrefix}parentId", parentId))
                        Using childReader = childCmd.ExecuteReader()
                            Dim childList = Activator.CreateInstance(prop.PropertyType)
                            While childReader.Read()
                                Dim childInstance = Activator.CreateInstance(childType)
                                For Each cp In childType.GetProperties().Where(Function(p) p.CanWrite AndAlso Not IsGenericList(p.PropertyType))
                                    If Not childReader.IsDBNull(childReader.GetOrdinal(cp.Name)) Then
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

            ' Get total count
            Dim countSql = $"SELECT COUNT(1) FROM [{tableName}] {whereSql}"
            Dim totalCount As Long
            Using countCmd = _provider.CreateCommand(countSql, connection)
                For Each param In parameters
                    countCmd.Parameters.Add(_provider.CloneParameter(param))
                Next
                totalCount = Convert.ToInt64(countCmd.ExecuteScalar())
            End Using

            Dim pageCount = Math.Ceiling(totalCount / maxPerPage)

            Return New Page(Of T)(results, pageNumber, maxPerPage, totalCount, CInt(pageCount))
        End Using
    End Function

#End Region

#Region "ExistsBy"
    Public Function ExistsBy(Of T)(conditions As List(Of Condition)) As Boolean Implements IOrm.ExistsBy

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
            parameters.Add(_provider.CreateParameter(paramName, condition.Value))
        Next

        Dim whereSql = String.Join(" AND ", whereClauses)
        Dim query = $"SELECT COUNT(1) FROM {tableName} WHERE {whereSql}"

        Using conn = _provider.CreateConnection()
            Using cmd = _provider.CreateCommand(query, conn)
                For Each param In parameters
                    cmd.Parameters.Add(param)
                Next
                conn.Open()
                Dim count = Convert.ToInt64(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    Public Function ExistsByInTable(Of T)(conditions As List(Of Condition), tableName As String) As Boolean Implements IOrm.ExistsByInTable

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
            parameters.Add(_provider.CreateParameter(paramName, condition.Value))
        Next

        Dim whereSql = String.Join(" AND ", whereClauses)
        Dim query = $"SELECT COUNT(1) FROM {tableName} WHERE {whereSql}"

        Using conn = _provider.CreateConnection()
            Using cmd = _provider.CreateCommand(query, conn)
                For Each param In parameters
                    cmd.Parameters.Add(param)
                Next
                conn.Open()
                Dim count = Convert.ToInt64(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

#End Region

#Region "CountBy"
    Public Function CountBy(Of T)(conditions As List(Of Condition)) As Long Implements IOrm.CountBy

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
            parameters.Add(_provider.CreateParameter(paramName, condition.Value))
        Next

        Dim whereSql = If(whereClauses.Any(), " WHERE " & String.Join(" AND ", whereClauses), "")
        Dim query = $"SELECT COUNT(1) FROM {tableName}{whereSql}"

        Using conn = _provider.CreateConnection()
            Using cmd = _provider.CreateCommand(query, conn)
                For Each param In parameters
                    cmd.Parameters.Add(param)
                Next
                conn.Open()
                Return Convert.ToInt64(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Public Function CountByInTable(Of T)(conditions As List(Of Condition), tableName As String) As Long Implements IOrm.CountByInTable

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
            parameters.Add(_provider.CreateParameter(paramName, condition.Value))
        Next

        Dim whereSql = If(whereClauses.Any(), " WHERE " & String.Join(" AND ", whereClauses), "")
        Dim query = $"SELECT COUNT(1) FROM {tableName}{whereSql}"

        Using conn = _provider.CreateConnection()
            Using cmd = _provider.CreateCommand(query, conn)
                For Each param In parameters
                    cmd.Parameters.Add(param)
                Next
                conn.Open()
                Return Convert.ToInt64(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

#End Region

#Region "DeleteBy"

    Public Sub DeleteBy(Of T)(conditions As List(Of Condition), Optional cascade As Boolean = True) Implements IOrm.DeleteBy

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        Dim typeT = GetType(T)
        Dim tableName = typeT.Name
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        For i = 0 To conditions.Count - 1
            Dim cond = conditions(i)
            Dim paramName = $"{parameterPrefix}p{i}"
            Dim sqlOp = GetSqlOperator(cond.SqlComparison)
            whereClauses.Add($"[{cond.Column}] {sqlOp} {paramName}")
            parameters.Add(_provider.CreateParameter(paramName, cond.Value))
        Next

        Dim whereSql = If(whereClauses.Any(), " WHERE " & String.Join(" AND ", whereClauses), "")
        If String.IsNullOrWhiteSpace(whereSql) Then
            Throw New InvalidOperationException("DeleteBy must have at least one condition to avoid deleting all records.")
        End If

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

                ' Check for unsafe DB-side cascade
                If Not cascade AndAlso HasCascadeDeleteConstraint(tableName, connection, transaction, _provider) Then
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

                        ' Build child delete with same WHERE conditions (on FK)
                        Dim childWhereSql = $" WHERE [{fkColumn}] IN (
                        SELECT [{tableName}].[Id] FROM [{tableName}]{whereSql}
                    )"
                        Dim childSql = $"DELETE FROM [{childTable}]{childWhereSql}"

                        Using childCmd = _provider.CreateCommand(childSql, connection)
                            childCmd.Transaction = transaction
                            For Each param In parameters
                                childCmd.Parameters.Add(_provider.CloneParameter(param))
                            Next
                            childCmd.ExecuteNonQuery()
                        End Using
                    Next
                End If

                ' Delete parent
                Dim parentSql = $"DELETE FROM [{tableName}]{whereSql}"
                Using cmd = _provider.CreateCommand(parentSql, connection)
                    cmd.Transaction = transaction
                    For Each param In parameters
                        cmd.Parameters.Add(_provider.CloneParameter(param))
                    Next
                    cmd.ExecuteNonQuery()
                End Using

                transaction.Commit()
            End Using
        End Using
    End Sub

    Public Sub DeleteByInTable(Of T)(conditions As List(Of Condition), tableName As String, Optional cascade As Boolean = True) Implements IOrm.DeleteByInTable

        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.")
        End If
        If String.IsNullOrEmpty(tableName) Then Throw New ArgumentException("Table Name cannot be null.")
        Dim typeT = GetType(T)
        Dim parameterPrefix = _provider.GetParameterPrefix()
        Dim whereClauses As New List(Of String)
        Dim parameters As New List(Of IDataParameter)

        For i = 0 To conditions.Count - 1
            Dim cond = conditions(i)
            Dim paramName = $"{parameterPrefix}p{i}"
            Dim sqlOp = GetSqlOperator(cond.SqlComparison)
            whereClauses.Add($"[{cond.Column}] {sqlOp} {paramName}")
            parameters.Add(_provider.CreateParameter(paramName, cond.Value))
        Next

        Dim whereSql = If(whereClauses.Any(), " WHERE " & String.Join(" AND ", whereClauses), "")
        If String.IsNullOrWhiteSpace(whereSql) Then
            Throw New InvalidOperationException("DeleteBy must have at least one condition to avoid deleting all records.")
        End If

        Using connection = _provider.CreateConnection()
            connection.Open()
            Using transaction = connection.BeginTransaction()

                ' Check for unsafe DB-side cascade
                If Not cascade AndAlso HasCascadeDeleteConstraint(tableName, connection, transaction, _provider) Then
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

                        ' Build child delete with same WHERE conditions (on FK)
                        Dim childWhereSql = $" WHERE [{fkColumn}] IN (
                        SELECT [{tableName}].[Id] FROM [{tableName}]{whereSql}
                    )"
                        Dim childSql = $"DELETE FROM [{childTable}]{childWhereSql}"

                        Using childCmd = _provider.CreateCommand(childSql, connection)
                            childCmd.Transaction = transaction
                            For Each param In parameters
                                childCmd.Parameters.Add(_provider.CloneParameter(param))
                            Next
                            childCmd.ExecuteNonQuery()
                        End Using
                    Next
                End If

                ' Delete parent
                Dim parentSql = $"DELETE FROM [{tableName}]{whereSql}"
                Using cmd = _provider.CreateCommand(parentSql, connection)
                    cmd.Transaction = transaction
                    For Each param In parameters
                        cmd.Parameters.Add(_provider.CloneParameter(param))
                    Next
                    cmd.ExecuteNonQuery()
                End Using

                transaction.Commit()
            End Using
        End Using
    End Sub

#End Region
End Class