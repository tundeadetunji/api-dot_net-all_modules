Imports System
Imports System.Collections.Generic
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Reflection

Public Class SequelOrm

#Region "props, fields and initialization"

    Private Shared _instance As SequelOrm
    Private Shared ReadOnly _lock As New Object()
    Private ReadOnly _connectionString As String
    Private ReadOnly _databaseName As String
    Private Const Id As String = "Id"

    Private Sub New(connectionString As String, databaseName As String)
        _connectionString = connectionString
        _databaseName = databaseName
    End Sub

#End Region

#Region "exported"

    ''' <summary>
    ''' Lightweight ORM.
    ''' Picks up only non read-only properties, may throw exception if any property is read-only.
    ''' You may need to install System.Data.SqlClient from Nuget if using >=.Net 5.
    ''' </summary>
    ''' <param name="connectionString"></param>
    ''' <param name="databaseName"></param>
    ''' <returns></returns>
    Public Shared Function GetInstance(connectionString As String, databaseName As String) As SequelOrm
        If _instance Is Nothing Then
            SyncLock _lock
                If _instance Is Nothing Then
                    _instance = New SequelOrm(connectionString, databaseName)
                End If
            End SyncLock
        End If
        Return _instance
    End Function

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Public Function FindById(Of T)(id As Object, Optional idColumn As String = Id) As T
        ' Ensure that the id is not null
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        ' Create the SQL query with parameterized id
        Dim query As String = $"SELECT * FROM {GetType(T).Name} WHERE {idColumn} = @{idColumn}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add the parameter based on the provided id
                command.Parameters.AddWithValue("@" & idColumn, id)
                connection.Open()

                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        Return MapToObject(Of T)(reader)
                    End If
                End Using
            End Using
        End Using

        Return Nothing
    End Function

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Public Function FindByIdInTable(Of T)(id As Object, tableName As String, Optional idColumn As String = Id) As T
        ' Ensure that the id is not null
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        ' Create the SQL query with parameterized id
        Dim query As String = $"SELECT * FROM {tableName} WHERE {idColumn} = @{idColumn}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add the parameter based on the provided id
                command.Parameters.AddWithValue("@" & idColumn, id)
                connection.Open()

                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        Return MapToObject(Of T)(reader)
                    End If
                End Using
            End Using
        End Using

        Return Nothing
    End Function

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Public Function FindByInTable(Of T)(conditions As Dictionary(Of String, Object), tableName As String) As List(Of T)
        Dim whereClause As String = String.Join(" AND ", conditions.Keys)
        Dim query As String = $"SELECT * FROM {tableName} WHERE {whereClause}"
        Dim results As New List(Of T)()
        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                For Each kvp In conditions
                    command.Parameters.AddWithValue("@" & kvp.Key, kvp.Value)
                Next
                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        results.Add(MapToObject(Of T)(reader))
                    End While
                End Using
            End Using
        End Using
        Return results
    End Function

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <returns></returns>
    Public Function FindBy(Of T)(conditions As Dictionary(Of String, Object)) As List(Of T)
        Dim whereClause As String = String.Join(" AND ", conditions.Keys)
        Dim query As String = $"SELECT * FROM {GetType(T).Name} WHERE {whereClause}"
        Dim results As New List(Of T)()
        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                For Each kvp In conditions
                    command.Parameters.AddWithValue("@" & kvp.Key, kvp.Value)
                Next
                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        results.Add(MapToObject(Of T)(reader))
                    End While
                End Using
            End Using
        End Using
        Return results
    End Function

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Public Function CreateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T
        Dim properties = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.Instance)
        Dim columns As New List(Of String)()
        Dim values As New List(Of String)()
        Dim idExists As Boolean = False
        Dim idValue As Object = Nothing

        ' Loop through properties to build columns and values
        For Each prop In properties
            If prop.Name.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                idExists = True
                idValue = prop.GetValue(obj)
            Else
                columns.Add(prop.Name)
                values.Add("@" & prop.Name)
            End If
        Next

        ' Build the SQL insert query
        Dim query As String
        If IdWillAutoIncrement Then
            query = $"INSERT INTO {tableName} ({String.Join(", ", columns)}) VALUES ({String.Join(", ", values)}); SELECT SCOPE_IDENTITY() AS NewId"
        Else
            query = $"INSERT INTO {tableName} ({idColumn}, {String.Join(", ", columns)}) VALUES (@{idColumn}, {String.Join(", ", values)})"
        End If

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add the idColumn parameter only if IdWillAutoIncrement is false
                If Not IdWillAutoIncrement AndAlso idExists Then
                    command.Parameters.AddWithValue("@" & idColumn, idValue)
                End If

                ' Add parameters for other properties
                For Each prop In properties
                    If Not prop.Name.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                        command.Parameters.AddWithValue("@" & prop.Name, prop.GetValue(obj))
                    End If
                Next

                connection.Open()
                Dim result = command.ExecuteScalar()
                If IdWillAutoIncrement Then
                    idValue = result
                End If
            End Using
        End Using

        ' Retrieve the newly created record
        Return GetRecordById(Of T)(idValue, idColumn, tableName)
    End Function

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Public Function Create(Of T)(obj As T, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T
        Dim properties = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.Instance)
        Dim columns As New List(Of String)()
        Dim values As New List(Of String)()
        Dim idExists As Boolean = False
        Dim idValue As Object = Nothing

        ' Loop through properties to build columns and values
        For Each prop In properties
            If prop.Name.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                idExists = True
                idValue = prop.GetValue(obj)
            Else
                columns.Add(prop.Name)
                values.Add("@" & prop.Name)
            End If
        Next

        ' Build the SQL insert query
        Dim query As String
        If IdWillAutoIncrement Then
            query = $"INSERT INTO {GetType(T).Name} ({String.Join(", ", columns)}) VALUES ({String.Join(", ", values)}); SELECT SCOPE_IDENTITY() AS NewId"
        Else
            query = $"INSERT INTO {GetType(T).Name} ({idColumn}, {String.Join(", ", columns)}) VALUES (@{idColumn}, {String.Join(", ", values)})"
        End If

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add the idColumn parameter only if IdWillAutoIncrement is false
                If Not IdWillAutoIncrement AndAlso idExists Then
                    command.Parameters.AddWithValue("@" & idColumn, idValue)
                End If

                ' Add parameters for other properties
                For Each prop In properties
                    If Not prop.Name.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                        command.Parameters.AddWithValue("@" & prop.Name, prop.GetValue(obj))
                    End If
                Next

                connection.Open()
                Dim result = command.ExecuteScalar()
                If IdWillAutoIncrement Then
                    idValue = result
                End If
            End Using
        End Using

        ' Retrieve the newly created record
        Return GetRecordById(Of T)(idValue, idColumn, GetType(T).Name)
    End Function

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Public Function UpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = Id) As T
        Dim properties = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.Instance)
        Dim columns As New List(Of String)()
        Dim values As New List(Of String)()
        Dim idValue As Object = Nothing

        ' Loop through properties to build columns and values
        For Each prop In properties
            If prop.Name.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                idValue = prop.GetValue(obj)
            Else
                columns.Add(prop.Name)
                values.Add("@" & prop.Name)
            End If
        Next

        ' Build the SQL update query
        Dim query As String = $"UPDATE {tableName} SET {String.Join(", ", columns.Select(Function(c) $"{c} = @{c}"))} WHERE {idColumn} = @{idColumn}; SELECT * FROM {GetType(T).Name} WHERE {idColumn} = @{idColumn}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add the idColumn parameter
                command.Parameters.AddWithValue("@" & idColumn, idValue)

                ' Add parameters for other properties
                For Each prop In properties
                    If Not prop.Name.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                        command.Parameters.AddWithValue("@" & prop.Name, prop.GetValue(obj))
                    End If
                Next

                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.Read() Then
                    ' Populate the updated record
                    Dim updatedRecord As T = Activator.CreateInstance(Of T)()
                    For Each prop In properties
                        prop.SetValue(updatedRecord, reader(prop.Name))
                    Next
                    Return updatedRecord
                End If
            End Using
        End Using

        ' Return null if no record was updated
        Return Nothing
    End Function

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Public Function Update(Of T)(obj As T, Optional idColumn As String = Id) As T
        Dim properties = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.Instance)
        Dim columns As New List(Of String)()
        Dim values As New List(Of String)()
        Dim idValue As Object = Nothing

        ' Loop through properties to build columns and values
        For Each prop In properties
            If prop.Name.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                idValue = prop.GetValue(obj)
            Else
                columns.Add(prop.Name)
                values.Add("@" & prop.Name)
            End If
        Next

        ' Build the SQL update query
        Dim query As String = $"UPDATE {GetType(T).Name} SET {String.Join(", ", columns.Select(Function(c) $"{c} = @{c}"))} WHERE {idColumn} = @{idColumn}; SELECT * FROM {GetType(T).Name} WHERE {idColumn} = @{idColumn}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                ' Add the idColumn parameter
                command.Parameters.AddWithValue("@" & idColumn, idValue)

                ' Add parameters for other properties
                For Each prop In properties
                    If Not prop.Name.Equals(idColumn, StringComparison.OrdinalIgnoreCase) Then
                        command.Parameters.AddWithValue("@" & prop.Name, prop.GetValue(obj))
                    End If
                Next

                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.Read() Then
                    ' Populate the updated record
                    Dim updatedRecord As T = Activator.CreateInstance(Of T)()
                    For Each prop In properties
                        prop.SetValue(updatedRecord, reader(prop.Name))
                    Next
                    Return updatedRecord
                End If
            End Using
        End Using

        ' Return null if no record was updated
        Return Nothing
    End Function

    Public Function Exists(id As Object, Optional idColumn As String = Id) As Boolean
        s
    End Function

    Public Function existsintable(id As Object, tableName As String, Optional idColumn As String = Id) As Boolean
        e
    End Function

    Public Function CreateOrUpdate(Of T)(obj As T, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T
        dfs
    End Function

    Public Function CreateOrUpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T
        dfs
    End Function


    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Public Function FindAllInTable(Of T)(tableName As String) As List(Of T)
        Dim query As String = $"SELECT * FROM {tableName}"
        Dim results As New List(Of T)()

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        results.Add(MapToObject(Of T)(reader))
                    End While
                End Using
            End Using
        End Using
        Return results
    End Function

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <returns></returns>
    Public Function FindAll(Of T)() As List(Of T)
        Dim query As String = $"SELECT * FROM {GetType(T).Name}"
        Dim results As New List(Of T)()

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        results.Add(MapToObject(Of T)(reader))
                    End While
                End Using
            End Using
        End Using
        Return results
    End Function


    Public Function Delete()
        dfs
    End Function

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    Public Sub DeleteWhereIdInTable(Of T)(id As Object, tableName As String, Optional idColumn As String = Id)
        ' Ensure that the id is not null
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        Dim query As String = $"DELETE FROM {tableName} WHERE {idColumn} = @{idColumn}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@" & idColumn, id)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="idColumn"></param>
    Public Sub DeleteWhereId(Of T)(id As Object, Optional idColumn As String = Id)
        ' Ensure that the id is not null
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        Dim query As String = $"DELETE FROM {GetType(T).Name} WHERE {idColumn} = @{idColumn}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@" & idColumn, id)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="tableName"></param>
    Public Sub DeleteAllInTable(Of T)(tableName As String)

        Dim query As String = $"DELETE FROM {tableName}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    Public Sub DeleteAll(Of T)()

        Dim query As String = $"DELETE FROM {GetType(T).Name}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

#End Region

#Region "support"


    Private Function MapToObject(Of T)(reader As SqlDataReader) As T
        Dim obj As T = Activator.CreateInstance(Of T)()
        Dim properties = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)

        For Each prop In properties
            If reader(prop.Name) IsNot DBNull.Value Then
                prop.SetValue(obj, reader(prop.Name))
            End If
        Next
        Return obj
    End Function

    Private Function Same(a As String, b As String) As Boolean
        Return String.Equals(a, b, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function GetRecordById(Of T)(idValue As Object, idColumn As String, tableName As String) As T
        Dim query As String = $"SELECT * FROM {tableName} WHERE {idColumn} = @Id"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@Id", idValue)

                connection.Open()
                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        Dim record As T = Activator.CreateInstance(Of T)()
                        Dim properties = GetType(T).GetProperties(BindingFlags.Public Or BindingFlags.Instance)

                        For Each prop In properties
                            If Not reader.IsDBNull(reader.GetOrdinal(prop.Name)) Then
                                prop.SetValue(record, reader(prop.Name))
                            End If
                        Next

                        Return record
                    End If
                End Using
            End Using
        End Using

        ' Return default value if no record is found
        Return Nothing
    End Function

#End Region

End Class