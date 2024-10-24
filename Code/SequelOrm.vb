Imports System
Imports System.Collections.Generic
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Text
Imports Microsoft.VisualBasic.Devices

''' <summary>
''' Lightweight ORM. Joins are coming in subsequent versions.
''' <br />
''' Picks up only non read-only, non Collection-type properties, may throw exception if any property is read-only or is IEnumerable/Collection/List/Enum.
''' The methods in this class are not asynchronous, do not employ locking mechanisms, and not covering transactions.
''' Joins are on the roadmap.
''' <br />
''' You may need to install System.Data.SqlClient from Nuget if using >=.Net 5.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public Class SequelOrm

#Region "consts and initialization related"

    Private Shared _instance As SequelOrm
    Private Shared ReadOnly _lockInit As New Object()
    Private ReadOnly _connectionString As String
    Private ReadOnly _databaseName As String
    Private Const Id As String = "Id"
    Private Const RowVersion As String = "RowVersion"

    Private Sub New(connectionString As String, databaseName As String)
        _connectionString = connectionString
        _databaseName = databaseName
    End Sub
    Public Shared Function GetInstance(connectionString As String, databaseName As String) As SequelOrm
        If _instance Is Nothing Then
            SyncLock _lockInit
                If _instance Is Nothing Then
                    _instance = New SequelOrm(connectionString, databaseName)
                End If
            End SyncLock
        End If
        Return _instance
    End Function

#End Region

#Region "exported"

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
        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.", NameOf(conditions))
        End If

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
        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.", NameOf(conditions))
        End If

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

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Public Function ExistsById(Of T)(id As Object, Optional idColumn As String = Id) As Boolean
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        Dim tableName As String = GetType(T).Name

        Dim query As String = $"SELECT COUNT(1) FROM {tableName} WHERE {idColumn} = @{idColumn}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@" & idColumn, id)
                connection.Open()
                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                Return count > 0 ' Return true if count is greater than 0
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <param name="id"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Public Function ExistsByIdInTable(id As Object, tableName As String, Optional idColumn As String = Id) As Boolean
        'Todo
        ' Ensure that the id is not null
        If id Is Nothing Then
            Throw New ArgumentNullException(NameOf(id), "ID cannot be null.")
        End If

        Dim query As String = $"SELECT COUNT(1) FROM {tableName} WHERE {idColumn} = @{idColumn}"

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@" & idColumn, id)
                connection.Open()
                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                Return count > 0 ' Return true if count is greater than 0
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <returns></returns>
    Public Function ExistsBy(Of T)(conditions As Dictionary(Of String, Object)) As Boolean
        Dim tableName As String = GetType(T).Name

        ' Validate input
        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.", NameOf(conditions))
        End If

        ' Construct the SQL query
        Dim query As New StringBuilder($"SELECT COUNT(1) FROM {tableName} WHERE ")
        Dim parameters As New List(Of String)()

        For Each condition In conditions
            parameters.Add($"{condition.Key} = @{condition.Key}")
        Next

        query.Append(String.Join(" AND ", parameters))

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query.ToString(), connection)
                ' Add parameters to the command
                For Each condition In conditions
                    command.Parameters.AddWithValue("@" & condition.Key, condition.Value)
                Next

                connection.Open()

                ' Execute the query and check if any records exist
                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                Return count > 0 ' Return true if count is greater than 0
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Public Function ExistsByInTable(Of T)(conditions As Dictionary(Of String, Object), tableName As String) As Boolean
        ' Validate input
        If conditions Is Nothing OrElse conditions.Count = 0 Then
            Throw New ArgumentException("Conditions cannot be null or empty.", NameOf(conditions))
        End If

        ' Construct the SQL query
        Dim query As New StringBuilder($"SELECT COUNT(1) FROM {tableName} WHERE ")
        Dim parameters As New List(Of String)()

        For Each condition In conditions
            parameters.Add($"{condition.Key} = @{condition.Key}")
        Next

        query.Append(String.Join(" AND ", parameters))

        Using connection As New SqlConnection(_connectionString)
            Using command As New SqlCommand(query.ToString(), connection)
                ' Add parameters to the command
                For Each condition In conditions
                    command.Parameters.AddWithValue("@" & condition.Key, condition.Value)
                Next

                connection.Open()

                ' Execute the query and check if any records exist
                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                Return count > 0 ' Return true if count is greater than 0
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Creates record if it doesn't exist, Updates otherwise.
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Public Function CreateOrUpdate(Of T)(obj As T, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T
        Dim propertyInfo As PropertyInfo = obj.GetType().GetProperty(idColumn, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        Dim id As Object
        If propertyInfo IsNot Nothing Then
            id = propertyInfo.GetValue(obj)
        Else
            Throw New ArgumentException($"Property {idColumn} not found in type {obj.GetType().Name}")
        End If

        Return If(ExistsById(Of T)(id, idColumn),
            Update(Of T)(obj, idColumn),
            Create(Of T)(obj, idColumn, IdWillAutoIncrement))

    End Function

    ''' <summary>
    ''' Creates record if it doesn't exist, Updates otherwise.
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Public Function CreateOrUpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = Id, Optional IdWillAutoIncrement As Boolean = True) As T
        Dim propertyInfo As PropertyInfo = obj.GetType().GetProperty(idColumn, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        Dim id As Object
        If propertyInfo IsNot Nothing Then
            id = propertyInfo.GetValue(obj)
        Else
            Throw New ArgumentException($"Property {idColumn} not found in type {obj.GetType().Name}")
        End If

        Return If(ExistsByIdInTable(id, tableName, idColumn),
            UpdateInTable(Of T)(obj, tableName, idColumn),
            CreateInTable(Of T)(obj, tableName, idColumn, IdWillAutoIncrement))
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