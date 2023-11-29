Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Imports System.Data.OleDb
Imports System.Data.SqlClient

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
            Throw New Exception("Error fetching data: " & vbCrLf & ex.ToString)
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
    ''' Returns Dictionary of first row returned by query
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values"></param>
    ''' <returns></returns>
    Public Shared Function QObject(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As Dictionary(Of String, Object)
        Dim l As New Dictionary(Of String, Object)
        Dim dt As DataTable = QDataTable(query, connection_string, select_parameter_keys_values)

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
    ''' Returns list of dictionaries (each of which correspond to row of the table, according to the output of query)
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values"></param>
    ''' <returns></returns>

    Public Shared Function QObjects(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As List(Of Dictionary(Of String, Object))
        Dim l As New List(Of Dictionary(Of String, Object))
        Dim dt As DataTable = QDataTable(query, connection_string, select_parameter_keys_values)
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


End Class
