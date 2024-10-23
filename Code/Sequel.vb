Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Data.Common

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Text

Public Class Sequel
    ''' <summary>
    ''' Commits record to MS Access database or Excel database.
    ''' </summary>
    ''' <param name="query">The SQL query.</param>
    ''' <param name="connection_string">The server connection string.</param>
    ''' <param name="parameters_keys_values_">Values to put in table. If saving image from FileUpload control, use FileBytes attribute of the FileUpload.</param>
    ''' <returns>True if successful, False if not.</returns>
    Public Shared Function CommitAccess(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing) As Boolean
        ' Dim d_c As New DataConnectionDesktop
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = parameters_keys_values_

        'If DB_Is_SQL_ = True Then
        '    CommitSQLRecord(query, connection_string, select_parameter_keys_values)
        '    Return True
        '    Exit Function
        'End If

        Try
            Dim insert_query As String = query
            Using insert_conn As New OleDbConnection(connection_string)
                Using insert_comm As New OleDbCommand()
                    With insert_comm
                        .Connection = insert_conn
                        .CommandText = insert_query
                        If select_parameter_keys_values IsNot Nothing Then
                            For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                                .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                            Next
                        End If
                    End With
                    Try
                        insert_conn.Open()
                        insert_comm.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End Using
            End Using
            Return True
        Catch ex As Exception
        End Try

    End Function

    Public Shared Function CommitSequel(query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As Boolean
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            Dim insert_query As String = query
            Using insert_conn As New SqlConnection(connection_string)
                Using insert_comm As New SqlCommand()
                    With insert_comm
                        .Connection = insert_conn
                        .CommandTimeout = 0
                        .CommandType = CommandType.Text
                        .CommandText = insert_query
                        If select_parameter_keys_values IsNot Nothing Then
                            For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                                .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                            Next
                        End If
                    End With
                    Try
                        insert_conn.Open()
                        insert_comm.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End Using
            End Using
            Return True
        Catch ex As Exception
        End Try

        '		Dim Entries_Insert As String = "INSERT INTO ENTRIES (EntryBy, ID, Category, [Description], Flag, [Title], Entry, DateAdded, TimeAdded, TitleID, Picture, PictureExtension, Topic) VALUES (@EntryBy, @ID, @Category, [@Description], @Flag, [@Title], @Entry, @DateAdded, @TimeAdded, @TitleID, @Picture, @PictureExtension, @Topic)"
        '		Dim entries_parameters_() = {"EntryBy", TitleBy.Text.Trim, "ID", EntryID.Text.Trim, "Category", Category.Text.Trim, "[Description]", Description.Text.Trim, "Flag", cFlag.Text.Trim, "[Title]", EntryTitle.Text.Trim, "Entry", NewEntry.Text.Trim, "DateAdded", date_, "TimeAdded", time_, "TitleID", TitleID.Text.Trim, "Picture", stream.GetBuffer(), "PictureExtension", PictureExtension.Text.Trim, "Topic", Topic.Text.Trim}
        '		d.CommitRecord(Entries_Insert, a_con, entries_parameters_)

    End Function
    Public Shared Function CommitExcel(query As String, connection_string As String) As Boolean
        Return CommitAccess(query, connection_string)
    End Function
    'Public Shared Function CommitExcel(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing) As Boolean
    '    Return CommitAccess(query, connection_string, parameters_keys_values_)
    'End Function
    ''' <summary>
    ''' Gets the data in a table. Traverse with .rows (Collection/List(Of String) and .columns (Collection/List(Of String). Can return a single row or all rows depending on the result of 'query'.
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values"></param>
    ''' <returns></returns>
    Public Shared Function QDataTable(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As DataTable

        Try
            Dim connection As New SqlConnection(connection_string)
            Dim sql As String = query

            Dim Command = New SqlCommand(sql, connection)
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If

            Dim da As New SqlDataAdapter(Command)
            Dim dt As New DataTable

            da.Fill(dt)
            Return dt
        Catch ex As Exception
            'Throw New Exception("Error fetching data: " & vbCrLf & ex.ToString)
            Return Nothing
        End Try

    End Function

    Public Shared Function QDataTableFromExcel(query As String, connection_string As String) As DataTable
        Dim con As New OleDbConnection(connection_string)
        Dim adapter As New OleDbDataAdapter(query, con)
        Dim dataSet As New DataSet()

        Try
            With con
                .Open()
                adapter.Fill(dataSet)
                con.Close()
            End With
            Return dataSet.Tables(0)
        Catch ex As Exception
        Finally
            Try
                con.Close()
            Catch ex As Exception

            End Try
        End Try

    End Function

    Public Shared Function QTablesInDatabase(connection_string As String) As List(Of String)
        Dim result As List(Of String) = New List(Of String)
        Try
            result = General.QList(General.BuildSelectString("sys.Tables", {"name"}, Nothing, "name"), connection_string)
        Catch ex As Exception
        End Try
        Return result
    End Function

    Public Shared Function QFieldsInTable(query As String, connection_string As String)
        Dim l As New List(Of String)
        Dim dt As DataTable = QDataTable(query, connection_string)
        If dt Is Nothing Then Return l
        With dt
            If .Columns.Count > 0 Then
                For i As Integer = 0 To .Columns.Count - 1
                    l.Add(.Columns.Item(i).ColumnName)
                Next
            End If
        End With

        Return l
    End Function
    ''' <summary>
    ''' Gets columns and their data types.
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values"></param>
    ''' <returns></returns>
    Public Shared Function QDataTypes(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As Dictionary(Of String, String)
        Dim dt As DataTable = QDataTable(query, connection_string)
        Dim result As New Dictionary(Of String, String)
        If dt Is Nothing Then Return result
        Try

            With dt
                If .Columns.Count > 0 Then
                    For i As Integer = 0 To .Columns.Count - 1
                        result.Add(.Columns(i).ColumnName, .Columns(i).DataType.Name)
                    Next
                End If
            End With
        Catch ex As Exception
            Throw New Exception("Error fetching data: " & vbCrLf & ex.ToString)
        End Try
        Return result
    End Function
    'Public Shared Function QFieldsInTable(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing)
    '    Dim l As New List(Of String)
    '    Dim dt As DataTable = QDataTable(query, connection_string, select_parameter_keys_values)

    '    With dt
    '        If .Columns.Count > 0 Then
    '            For i As Integer = 0 To .Columns.Count - 1
    '                l.Add(.Columns.Item(i).ColumnName)
    '            Next
    '        End If
    '    End With

    '    Return l
    'End Function

    Public Shared Function QTablesInDatabaseFromExcel(connection_string As String) As List(Of String)
        Dim result As New List(Of String)
        Try

            Dim con As New OleDbConnection(connection_string)
            With con
                .Open()

                Dim d As DataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                For i = 0 To d.Rows.Count - 1
                    result.Add(d.Rows(i).Item("TABLE_NAME").ToString)
                Next

                .Close()
            End With
        Catch ex As Exception
        End Try
        Return result
    End Function

    Public Shared Function QFieldsInTableFromExcel(query As String, connection_string As String) As List(Of String)
        Dim l As New List(Of String)
        Dim dt As DataTable = QDataTableFromExcel(query, connection_string)

        With dt
            If .Columns.Count > 0 Then
                For i As Integer = 0 To .Columns.Count - 1
                    l.Add(.Columns.Item(i).ColumnName)
                Next
            End If
        End With

        Return l
    End Function
    ''' <summary>
    ''' Same as QObject
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values"></param>
    ''' <returns></returns>
    Public Shared Function QRow(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As Dictionary(Of String, Object)
        Return QObject(query, connection_string, select_parameter_keys_values)
    End Function

    ''' <summary>
    ''' Returns Dictionary of first row returned by query
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values"></param>
    ''' <returns></returns>
    Public Shared Function QObject(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As Dictionary(Of String, Object)
        Dim dt As DataTable = QDataTable(query, connection_string, select_parameter_keys_values)
        If dt Is Nothing Then Return Nothing
        If dt.Rows.Count < 1 Then Return Nothing
        Dim l As New Dictionary(Of String, Object)

        With dt
            'For i = 0 To .Rows.Count - 1
            For i = 0 To 0
                For j As Integer = 0 To .Columns.Count - 1
                    l.Add(.Columns.Item(j).ColumnName, .Rows(i).Item(j))
                Next
            Next
        End With

        Return l
    End Function

    ''' <summary>
    ''' Same as QObjects.
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values"></param>
    ''' <returns></returns>
    Public Shared Function QRows(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As List(Of Dictionary(Of String, Object))
        Return QObjects(query, connection_string, select_parameter_keys_values)
    End Function

    ''' <summary>
    ''' Returns list of dictionaries (each of which correspond to row of the table, according to the output of query)
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values"></param>
    ''' <returns></returns>

    Public Shared Function QObjects(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As List(Of Dictionary(Of String, Object))
        Dim l As New List(Of Dictionary(Of String, Object))
        Dim dt As DataTable = QDataTable(query, connection_string, select_parameter_keys_values)
        If dt Is Nothing Then Return l
        If dt.Rows.Count < 1 Then Return l
        With dt
            For i = 0 To .Rows.Count - 1
                Dim r As New Dictionary(Of String, Object)
                For j As Integer = 0 To .Columns.Count - 1
                    r.Add(.Columns.Item(j).ColumnName, .Rows(i).Item(j))
                Next
                l.Add(r)
            Next
        End With
        Return l
    End Function

    ''' <summary>
    ''' Returns Type originally saved in database. Assigns column names to private and public properties.
    ''' Assumes the name of the table is the same as the name of the type.
    ''' Joins or transactions are not allowed for.
    ''' May throw error if any of the properties is Read-only.
    ''' Prefer using equivalent in iNovation.Code.SequelOrm.*
    ''' </summary>
    ''' <typeparam name="T">The return type.</typeparam>
    ''' <param name="Entity">Instance of the return type.</param>
    ''' <returns></returns>
    Public Shared Function QEntity(Of T As Class)(Entity As T, connection_string As String, Id As Object, Optional IdColumnName As String = "Id") As T

        Dim table As String = getClassName(Entity)
        Dim container As PropContainer = SortProperties(Entity, False)
        Dim noncollection As Dictionary(Of String, Object) = container.NonCollection

        Dim tempContainer As PropContainer = SortProperties(Entity, False)
        Dim tempNoncollection As Dictionary(Of String, Object) = tempContainer.NonCollection
        removeIdKey(tempNoncollection, IdColumnName)

        Dim kv = General.DictionaryKeyValueToArray(noncollection)
        Dim query As String = General.BuildSelectString(table, Nothing, {IdColumnName})
        Dim r = QRow(query, connection_string, {IdColumnName, Id})
        Return QMap(Of T)(r, Entity)

    End Function

    ''' <summary>
    ''' Inserts a new record, directly from the type.
    ''' Assumes the name of the table is the same as the name of the type.
    ''' Joins or transactions are not allowed for.
    ''' Prefer using equivalent in iNovation.Code.SequelOrm.*
    ''' </summary>
    ''' <typeparam name="T">The type.</typeparam>
    ''' <param name="Entity">Instance of the type.</param>
    ''' <param name="connection_string"></param>
    ''' <param name="IdIsAutoIncrement">Id auto-increments by 1 (default).</param>
    Public Shared Sub QEntityCreate(Of T As Class)(Entity As T, connection_string As String, Optional IdIsAutoIncrement As Boolean = True)
        Dim container As PropContainer = SortProperties(Entity, IdIsAutoIncrement)
        Dim collection As Dictionary(Of String, Object) = container.Collection
        Dim noncollection As Dictionary(Of String, Object) = container.NonCollection

        Dim table As String = getClassName(Entity)
        Dim result As StringBuilder = New StringBuilder

        CommitSequel(createSaveQuery(table, noncollection), connection_string)
    End Sub

    ''' <summary>
    ''' Updates a record, directly from the type.
    ''' Assumes the name of the table is the same as the name of the type.
    ''' Joins or transactions are not allowed for.
    ''' Prefer using equivalent in iNovation.Code.SequelOrm.*
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Entity"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="IdColumnName"></param>
    ''' <param name="Id"></param>

    Public Shared Sub QEntityUpdate(Of T As Class)(Entity As T, connection_string As String, Id As Object, Optional IdColumnName As String = "Id")
        Dim table As String = getClassName(Entity)
        Dim container As PropContainer = SortProperties(Entity, False)
        Dim noncollection As Dictionary(Of String, Object) = container.NonCollection

        Dim tempContainer As PropContainer = SortProperties(Entity, False)
        Dim tempNoncollection As Dictionary(Of String, Object) = tempContainer.NonCollection
        removeIdKey(tempNoncollection, IdColumnName)

        Dim kv = General.DictionaryKeyValueToArray(noncollection)
        Dim query As String = General.BuildUpdateString(table, tempNoncollection.Keys.ToArray, {IdColumnName})
        CommitSequel(query, connection_string, kv)

    End Sub

    Private Shared Sub removeIdKey(dict As Dictionary(Of String, Object), k As String)
        dict.Remove(k.ToUpper)
        dict.Remove(k)
        dict.Remove(k.ToLower)
    End Sub

    ''' <summary>
    ''' Deletes a record, directly from the type.
    ''' Assumes the name of the table is the same as the name of the type.
    ''' Joins or transactions are not allowed for.
    ''' Prefer using equivalent in iNovation.Code.SequelOrm.*
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Entity"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="IdColumnName"></param>
    ''' <param name="Id"></param>
    Public Shared Sub QEntityDelete(Of T As Class)(Entity As T, connection_string As String, Id As Object, Optional IdColumnName As String = "Id")
        Dim table As String = getClassName(Entity)

        Dim query As String = General.DeleteString_CONDITIONAL(table, {IdColumnName, "="})
        CommitSequel(query, connection_string, {IdColumnName, Id})
    End Sub

    Public Shared Function createDatabaseQuery(ByVal databaseName As String) As String
        Return $"CREATE DATABASE {databaseName}"
    End Function

    ''' <summary>
    ''' Ensures the database exists beforehand.
    ''' Transactions are not allowed for.
    ''' Prefer using equivalent in iNovation.Code.SequelOrm.*
    ''' </summary>
    ''' <param name="database_name"></param>
    ''' <param name="connection_string"></param>
    Public Shared Sub QEntityEnsureDatabase(database_name As String, connection_string As String)
        Using connection As New SqlConnection(connection_string)
            connection.Open()

            ' Check if the database exists
            Dim dbExists As Boolean = False
            Dim checkDbQuery As String = $"SELECT database_id FROM sys.databases WHERE name = '{database_name}'"

            Using command As New SqlCommand(checkDbQuery, connection)
                Dim result = command.ExecuteScalar()
                dbExists = (result IsNot Nothing)
            End Using

            ' Create the database if it does not exist
            If Not dbExists Then
                Using command As New SqlCommand(createDatabaseQuery(database_name), connection)
                    command.ExecuteNonQuery()
                    'Console.WriteLine($"Database '{DatabaseName}' created.")
                End Using
            Else
                'Console.WriteLine($"Database '{DatabaseName}' already exists.")
            End If
        End Using
    End Sub
    ''' <summary>
    ''' Prefer using equivalent in iNovation.Code.SequelOrm.*
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Entity"></param>
    ''' <param name="idPropertyName"></param>
    ''' <param name="stringTargetType"></param>
    ''' <returns></returns>
    Public Shared Function CreateTableQuery(Of T As Class)(ByVal Entity As T,
                                                      Optional ByVal idPropertyName As String = "Id",
                                                      Optional stringTargetType As String = "NVARCHAR(MAX)") As String
        Dim tableName As String = Entity.GetType().Name
        Dim columns As New List(Of String)
        Dim hasIdProperty As Boolean = False

        ' Check if the object has an "Id" property
        For Each propertyInfo As PropertyInfo In Entity.GetType().GetProperties(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            Dim columnName As String = propertyInfo.Name
            Dim dataType As String = GetSqlDataType(propertyInfo.PropertyType, stringTargetType)

            ' If the property is the "Id" property, mark it as the primary key
            If columnName.Equals(idPropertyName, StringComparison.OrdinalIgnoreCase) Then
                columns.Add($"{columnName} INT PRIMARY KEY IDENTITY(1,1)")
                hasIdProperty = True
            Else
                columns.Add($"{columnName} {dataType}")
            End If
        Next

        ' If the object does not have an "Id" property, add it as the primary key
        If Not hasIdProperty Then
            columns.Insert(0, $"{idPropertyName} INT PRIMARY KEY IDENTITY(1,1)")
        End If

        Dim query As New StringBuilder()
        query.AppendFormat("CREATE TABLE {0} (", tableName)
        query.Append(String.Join(", ", columns))
        query.Append(")")

        Return query.ToString()
    End Function
    Private Shared Function GetSqlDataType(propertyType As Type, Optional stringTargetType As String = "NVARCHAR(MAX)") As String
        Select Case propertyType.Name
            Case "String"
                Return stringTargetType
            Case "Integer", "Int32"
                Return "INT"
            Case "Long", "Int64"
                Return "BIGINT"
            Case "Boolean"
                Return "BIT"
            Case "DateTime"
                Return "DATETIME"
            Case "Decimal"
                Return "DECIMAL(18, 2)"
            Case "Double"
                Return "FLOAT"
            Case Else
                Throw New ArgumentException($"Unsupported data type: {propertyType.Name}")
        End Select
    End Function
    ''' <summary>
    ''' Ensures the table is in the database beforehand. You can call this before any CRUD operation.
    ''' Transactions are not allowed for.
    ''' Prefer using equivalent in iNovation.Code.SequelOrm.*
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Entity"></param>
    ''' <param name="database_name"></param>
    ''' <param name="connection_string"></param>
    Public Shared Sub QEntityEnsureTable(Of T As Class)(ByVal Entity As T, database_name As String, connection_string As String, Optional ByVal idPropertyName As String = "Id", Optional stringTargetType As String = "NVARCHAR(MAX)")
        QEntityEnsureDatabase(database_name, connection_string)

        Using connection As New SqlConnection(connection_string)
            connection.Open()
            connection.ChangeDatabase(database_name)

            ' Check if the table exists
            Dim tableExists As Boolean = False
            Dim checkTableQuery As String = $"SELECT COUNT(*) FROM information_schema.tables WHERE table_name = '{getClassName(Entity)}'"

            Using command As New SqlCommand(checkTableQuery, connection)
                Dim result = command.ExecuteScalar()
                tableExists = (Convert.ToInt32(result) > 0)
            End Using

            ' Create the table if it does not exist
            If Not tableExists Then
                Using command As New SqlCommand(CreateTableQuery(Entity, idPropertyName, stringTargetType), connection)
                    command.ExecuteNonQuery()
                    'Console.WriteLine($"Table '{tableName}' created with the provided query.")
                End Using
            Else
                'Console.WriteLine($"Table '{tableName}' already exists.")
            End If
        End Using
    End Sub

    Private Shared Function getClassName(entity As Object)
        Return entity.GetType.Name
    End Function

    ''' <summary>
    ''' Returns type directly from the value obtained from calling Sequel.QRow()
    ''' Assumes the name of the type is the same as the name of the type.
    ''' Joins or transactions are not allowed for.
    ''' Prefer using equivalent in iNovation.Code.SequelOrm.*
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="dictionary">Value from calling QRow</param>
    ''' <param name="Entity"></param>
    ''' <returns></returns>
    Public Shared Function QMap(Of T As Class)(dictionary As Dictionary(Of String, Object), Entity As T) As T
        If dictionary Is Nothing Then
            Throw New ArgumentNullException(NameOf(dictionary))
        End If

        If Entity Is Nothing Then
            Throw New ArgumentNullException(NameOf(Entity))
        End If

        Dim type = GetType(T)
        Dim properties = type.GetProperties(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        Dim value As Object = Nothing

        For Each prop In properties
            If dictionary.TryGetValue(prop.Name, value) Then
                If Not prop.CanWrite Then
                    Throw New InvalidOperationException($"Property {prop.Name} is readonly and cannot be set.")
                End If
                prop.SetValue(Entity, value)
            End If
        Next

        ' Handle private fields
        Dim fields = type.GetFields(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)

        For Each field In fields
            If dictionary.TryGetValue(field.Name, value) Then
                field.SetValue(Entity, value)
            End If
        Next

        Return Entity

        'Dim mapper As SequelOrm = New SequelOrm
        'Dim amy As User = mapper.MapDictionaryToObject(dict, New User)


    End Function

    Private Structure KvContainer
        Public keys As Array
        Public keys_values As Array
    End Structure

    Private Shared Function CreateKvContainer(Of T)(ByVal Entity As T) As KvContainer
        ' Get all public and private properties of the object
        Dim properties As PropertyInfo() = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)

        ' Initialize arrays to hold property names and property name-value pairs
        Dim keys As New List(Of String)()
        Dim keys_params As New List(Of String)()

        ' Iterate through each property
        For Each prop As PropertyInfo In properties
            ' Get the property name
            Dim propertyName As String = prop.Name
            ' Get the property value
            Dim propertyValue As Object = prop.GetValue(Entity)

            ' Add the property name to keys array
            keys.Add(propertyName)
            ' Add the property name and value to keys_params array
            keys_params.Add($"{propertyName}: {propertyValue}")
        Next

        ' Convert lists to arrays
        Return New KvContainer With {.keys = keys.ToArray(), .keys_values = keys_params.ToArray()}
    End Function
#Region "JSON"

    ''' <summary>
    ''' Converts a row to JSON Object. query must return one row only.
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values_"></param>
    ''' <returns>JObject</returns>
    Public Shared Function RowToJSON(query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As JObject
        Try
            Dim dt As DataTable = QDataTable(query, connection_string, select_parameter_keys_values_)
            Dim d As New Dictionary(Of String, Object)
            With dt
                For i = 0 To .Rows.Count - 1
                    For j = 0 To .Columns.Count - 1
                        d.Add(.Columns.Item(j).ColumnName, .Rows(i).Item(j))
                    Next
                Next
            End With
            Dim jsonObject As JObject = JObject.Parse(JsonConvert.SerializeObject(d))
            Return jsonObject

        Catch ex As Exception

        End Try
    End Function
    ''' <summary>
    ''' Converts all rows (i.e. table) to objects and returns a list (JArray) containing them
    ''' </summary>
    ''' <param name="t_"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="pk"></param>
    ''' <param name="select_params"></param>
    ''' <param name="where_keys"></param>
    ''' <param name="select_parameter_keys_values_"></param>
    ''' <param name="OrderByField"></param>
    ''' <param name="order_by"></param>
    ''' <returns>JArray</returns>
    Public Shared Function RowsToJSON(t_ As String, connection_string As String, pk As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional select_parameter_keys_values_ As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As General.OrderBy = General.OrderBy.ASC) As JArray
        Dim overall_q = General.BuildSelectString(t_, select_params, where_keys, OrderByField, order_by)
        Dim dt As DataTable = QDataTable(overall_q, connection_string, select_parameter_keys_values_)
        Dim l As New List(Of JObject)
        For i = 0 To dt.Rows.Count - 1
            Dim pk_query = General.BuildSelectString(t_, {pk})
            Dim pk_value As Object = dt.Rows(i).Item(pk)
            Dim pk_kv = {pk, pk_value}
            Dim q = General.BuildSelectString(t_, Nothing, {pk})
            ''Dim dt_ As DataTable = GetDataTable(q, connection_string, {pk, pk_value})
            Dim jobject As JObject = RowToJSON(q, connection_string, pk_kv)
            l.Add(jobject)
        Next
        Dim j As JArray = JArray.Parse(JsonConvert.SerializeObject(l))
        Return j
    End Function
    Public Enum OutputAs
        JObject = 0
        String_ = 1
    End Enum
    Private Shared Function SampleJSON(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing, Optional table_ As String = "whatever", Optional output_as As OutputAs = OutputAs.String_) As Object
        Dim t As DataTable = QDataTable(query, connection_string, select_parameter_keys_values)
        Dim s As String = "{""table"": """ & table_ & """, ""shoulddisplay"": ""all"", ""shouldupdate"": ""id"", ""shoulddelete"": ""enabled"", "
        Dim r
        With t
            For i = 0 To .Columns.Count - 1
                s &= """" & .Columns.Item(i).ColumnName & """: "
                If .Columns.Item(i).DataType = GetType(Boolean) Then
                    s &= True.ToString.ToLower
                ElseIf .Columns.Item(i).DataType = GetType(Date) Then
                    s &= """" & Date.Now & """"
                ElseIf .Columns.Item(i).DataType = GetType(Byte) Then
                    s &= """byte_value"""
                ElseIf .Columns.Item(i).DataType = GetType(TimeSpan) Then
                    s &= """timespan_value"""
                Else
                    s &= """" & .Columns.Item(i).ColumnName.ToLower & """"
                End If
                If i <> .Columns.Count - 1 Then
                    s &= ", "
                End If
            Next
        End With
        s &= "}"
        If output_as = OutputAs.JObject Then
            r = JObject.Parse(s)
        ElseIf output_as = OutputAs.String_ Then
            r = s
        End If
        Return r
    End Function
#End Region

#Region "CSV"
    Public Shared Function ConvertDataToCSV(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing) As String
        If Mid(query, 1, Len("select")).ToLower <> "select" And Mid(query, 1, Len("update")).ToLower <> "update" Then
            Return Nothing
        End If

        Dim data As DataTable = QDataTable(query, connection_string, parameters_keys_values_)
        Dim header As String = ""
        Dim content As String = ""

        With data
            'header
            For h = 0 To .Columns.Count - 1
                header &= .Columns(h).ColumnName
                If h <> .Columns.Count Then header &= ","
            Next

            For i = 0 To .Rows.Count - 1
                For j = 0 To .Columns.Count - 1
                    'content
                    content &= .Rows(i).Item(j)
                    If j <> .Columns.Count Then content &= ","
                Next
                content &= vbCrLf
            Next
        End With

        Return header.Trim & vbCrLf & content.Trim
    End Function
    'Public Shared Sub CommitData(where_to_place_data As CommitDataTargetInfo, query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing)
    '    If Mid(query, 1, Len("select")).ToLower <> "select" And Mid(query, 1, Len("update")).ToLower <> "update" Then
    '        Return
    '    End If

    '    Dim data As DataTable = GetDataTable(query, connection_string, parameters_keys_values_)
    '    Dim header As String = ""
    '    Dim content As String = ""

    '    With data
    '        'header
    '        For h = 0 To .Columns.Count - 1
    '            header &= .Columns(h).ColumnName
    '            If h <> .Columns.Count Then header &= ","
    '        Next

    '        For i = 0 To .Rows.Count - 1
    '            For j = 0 To .Columns.Count - 1
    '                'content
    '                content &= .Rows(i).Item(j)
    '                If j <> .Columns.Count Then content &= ","
    '            Next
    '            content &= vbCrLf
    '        Next
    '    End With

    '    WriteText(where_to_place_data.filename, header.Trim & vbCrLf & content.Trim)
    'End Sub

#End Region

#Region "Entity ORM related"

    Public Structure PropContainer
        Public Property Collection As Dictionary(Of String, Object)
        Public Property NonCollection As Dictionary(Of String, Object)
    End Structure


    Private Shared ReadOnly Property Id As String = "Id"
    Private Shared ReadOnly Property SemiColon As String = ";"

    Private Shared Function SortProperties(classInstance As Object, IdIsAutoIncrement As Boolean) As PropContainer
        Dim nonCollectionFields As New Dictionary(Of String, Object)
        Dim collectionFields As New Dictionary(Of String, Object)

        For Each field As PropertyInfo In classInstance.GetType().GetProperties(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
            Dim fieldValue As Object = field.GetValue(classInstance)
            Dim propertyType As Type = field.PropertyType

            If field.Name.Equals(Id, StringComparison.OrdinalIgnoreCase) AndAlso IdIsAutoIncrement Then
                ' do not add the "Id" property if IdIsAutoIncrement is true
                Continue For
            End If

            If propertyType.IsGenericType AndAlso propertyType.GetGenericTypeDefinition() = GetType(List(Of )) OrElse propertyType.IsEnum Then
                collectionFields.Add(field.Name, fieldValue)
            Else
                nonCollectionFields.Add(field.Name, fieldValue)
            End If
        Next

        Return New PropContainer With {.Collection = collectionFields, .NonCollection = nonCollectionFields}
    End Function

    Public Shared Function GetPropertyValue(instance As Object, propertyName As String) As Object
        Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(propertyName, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        If propertyInfo IsNot Nothing Then
            Return propertyInfo.GetValue(instance)
        Else
            Throw New ArgumentException($"Property '{propertyName}' not found in type {instance.GetType().Name}")
        End If
    End Function

    Public Shared Function GetProperties(instance As Object, parentEntityName As String, primaryKey As Integer) As Dictionary(Of String, Object)
        Dim properties As New Dictionary(Of String, Object)
        For Each prop As PropertyInfo In instance.GetType().GetProperties(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
            If prop.Name.Equals(parentEntityName & Id) Then
                properties.Add(prop.Name, primaryKey)
            Else
                properties.Add(prop.Name, prop.GetValue(instance))
            End If
        Next
        Return properties
    End Function

    Public Shared Function HasProperty(ByVal Entity As Object, ByVal propertyName As String) As Boolean
        Dim type As Type = Entity.GetType()
        Dim propertyInfo As PropertyInfo = type.GetProperty(propertyName, BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.NonPublic)

        Return propertyInfo IsNot Nothing
    End Function
    Public Shared Function createSaveQuery(tableName As String, data As Dictionary(Of String, Object)) As String
        Dim query As New StringBuilder
        Dim values As New List(Of String)

        query.Append($"INSERT INTO {tableName} (")
        For Each key As String In data.Keys
            query.Append($"{key}, ")
        Next
        query.Length -= 2 ' Remove the last comma and space
        query.Append(") VALUES (")

        For Each value As Object In data.Values
            If value Is Nothing Then
                values.Add("NULL")
            ElseIf TypeOf value Is String Then
                values.Add($"'{value.ToString().Replace("'", "''")}'")
            ElseIf TypeOf value Is Date Then
                values.Add($"'{value.ToString("yyyy-MM-dd HH:mm:ss")}'")
            Else
                values.Add(value.ToString())
            End If
        Next

        query.Append(String.Join(", ", values))
        query.Append(")")

        Return query.Append(SemiColon).ToString()
    End Function

#End Region

End Class
