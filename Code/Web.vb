'from iNovation
Imports iNovation.Code.General
Imports System.Data.OleDb
Imports iNovation.Code.Values

'from Assemblies
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls
Imports System.Net
Imports System.IO

''' <summary>
''' This class contains methods geared toward web components in ASP.NET. It is equivalent to Desktop.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public Class Web

#Region "Web"

    Public Shared Function toTitleCase(str As String) As String
        Return str.Substring(0, 1).ToUpper + str.Substring(1)
    End Function

    Public Shared Function stringToEnum(type As String) As String
        Return type.Replace("_", "")
    End Function


    '''' <summary>
    '''' Same as NModule.W.GenderDrop.
    '''' </summary>
    '''' <param name="d_"></param>
    '''' <param name="FirstIsEmpty"></param>
    'Public Sub GenderDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
    '    If d_.Items.Count > 0 Then Exit Sub

    '    Dim web_methods_ As New Methods

    '    Clear(d_)

    '    If FirstIsEmpty Then d_.Items.Add("")

    '    Dim l As List(Of String) = web_methods_.GenderList()
    '    For i As Integer = 0 To l.Count - 1
    '        d_.Items.Add(l(i).ToString)
    '    Next
    'End Sub

    Public Sub DropTextBoolean(d_ As DropDownList, Optional pattern_ As String = "always/never", Optional FirstIsEmpty As Boolean = True)
        If d_.Items.Count > 0 Then Exit Sub

        Clear(d_)

        If FirstIsEmpty Then d_.Items.Add("")

        With d_.Items
            Select Case pattern_.Trim.ToLower
                Case ""
                    .Add("Always")
                    .Add("Never")
                Case "yes/no"
                    .Add("Yes")
                    .Add("No")
                Case "always/never"
                    .Add("Always")
                    .Add("Never")
                Case "on/off"
                    .Add("On")
                    .Add("Off")
                Case "1/0"
                    .Add("1")
                    .Add("0")
                Case "true/false"
                    .Add("True")
                    .Add("False")
            End Select

        End With

    End Sub

    Private Sub CData(c_ As CheckBoxList, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing)
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            c_.DataSource = Nothing
        Catch ex As Exception
        End Try

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

            c_.DataSource = dt
            c_.DataTextField = data_text_field
            If data_value_field.Length > 0 Then c_.DataValueField = data_value_field
            c_.DataBind()
        Catch
        End Try


        '		d.GData(gPayment, Payment_, g_con)

        '		Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
        '		d.GData(gPayment, School_, m_con, select_parameter_keys_values)

    End Sub

    Private Sub CData_(c_ As CheckBoxList, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys As ListBox = Nothing, Optional select_parameter_values As ListBox = Nothing)
        Try
            c_.DataSource = Nothing
        Catch ex As Exception

        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys IsNot Nothing And select_parameter_values IsNot Nothing And select_parameter_keys.Items.Count > 0 And select_parameter_values.Items.Count > 0 Then
                select_parameter_keys.SelectedIndex = 0 : select_parameter_values.SelectedIndex = 0
                For i% = 0 To select_parameter_keys.Items.Count - 1
                    Command.Parameters.AddWithValue(select_parameter_keys.Items.Item(i).Value, select_parameter_values.Items.Item(i).Value)
                Next
                select_parameter_keys.SelectedIndex += 1
                select_parameter_values.SelectedIndex += 1
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        c_.DataSource = dt
        c_.DataTextField = data_text_field
        If data_value_field.Length > 0 Then c_.DataValueField = data_value_field
        c_.DataBind()

    End Sub

    Private Sub ClearList(l_ As ListBox)
        Try
            l_.DataSource = Nothing
        Catch ex As Exception
        End Try
        l_.Items.Clear()
    End Sub

    Private Sub ClearDropDown(d_ As DropDownList)
        Try
            d_.DataSource = Nothing
        Catch ex As Exception
        End Try
        d_.Items.Clear()
    End Sub

    'Public Sub Clear_(Optional l_ As ListBox = Nothing, Optional d_ As DropDownList = Nothing)
    '    Try
    '        If l_ IsNot Nothing Then
    '            ClearList(l_)
    '        End If
    '    Catch
    '    End Try
    '    Try
    '        If d_ IsNot Nothing Then
    '            ClearDropDown(d_)
    '        End If
    '    Catch
    '    End Try
    'End Sub

    ''' <summary>
    ''' Binds SQL Server database table to GridView.
    ''' </summary>
    ''' <param name="g_">GridView to bind to.</param>
    ''' <param name="query">The SQL query.</param>
    ''' <param name="connection_string">The server connection string.</param>
    ''' <param name="select_parameter_keys_values_">The SQL select parameters.</param>
    Private Sub GData(g_ As GridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing)
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            g_.DataSource = Nothing
        Catch ex As Exception
        End Try

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

            g_.DataSource = dt
            g_.DataBind()
        Catch
        End Try


        '		d.GData(gPayment, Payment_, g_con)

        '		Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
        '		d.GData(gPayment, School_, m_con, select_parameter_keys_values)

    End Sub

    ''' <summary>
    ''' Binds SQL Server database table to GridView. Returns the GridView. Same as GData.
    ''' </summary>
    ''' <param name="g_">GridView to bind to.</param>
    ''' <param name="query">The SQL query.</param>
    ''' <param name="connection_string">The server connection string.</param>
    ''' <param name="select_parameter_keys_values_">The SQL select parameters.</param>
    ''' <return>g_</return>
    Public Shared Function Display(g_ As GridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As GridView
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            g_.DataSource = Nothing
        Catch ex As Exception
        End Try

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

            g_.DataSource = dt
            g_.DataBind()
        Catch
        End Try

        Return g_
        '		d.GData(gPayment, Payment_, g_con)

        '		Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
        '		d.GData(gPayment, School_, m_con, select_parameter_keys_values)

    End Function

    Public Shared Function DisplayFromExcel(g_ As GridView, query As String, connection_string As String) As GridView
        Try
            With g_
                .DataSource = Nothing
            End With
        Catch ex As Exception

        End Try
        Dim table As DataTable = iNovation.Code.Sequel.QDataTableFromExcel(query, connection_string)
        g_.DataSource = table
        g_.DataBind()
        Return g_
    End Function
    'Public Shared Function DisplayFromExcel(g_ As GridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As GridView
    '    Try
    '        With g_
    '            .DataSource = Nothing
    '        End With
    '    Catch ex As Exception

    '    End Try
    '    Dim table As DataTable = iNovation.Code.Sequel.QDataTableFromExcel(query, connection_string, select_parameter_keys_values_)
    '    g_.DataSource = table
    '    g_.DataBind()
    '    Return g_
    'End Function

    ''' <summary>
    ''' Binds DropDownList to SQL database column.
    ''' </summary>
    ''' <param name="d_"></param>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="data_text_field"></param>
    ''' <param name="data_value_field"></param>
    ''' <param name="select_parameter_keys_values_"></param>
    ''' <param name="DontIfFull">Ignores the function if d_ already has items</param>
    ''' <returns>d_</returns>
    Private Shared Function DData(d_ As DropDownList, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing, Optional DontIfFull As Boolean = False) As DropDownList
        If DontIfFull = True Then
            If d_.Items.Count > 0 Then Return d_
        End If
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            d_.DataSource = Nothing
        Catch ex As Exception

        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        d_.DataSource = dt
        d_.DataTextField = data_text_field
        If data_value_field.Length > 0 Then d_.DataValueField = data_value_field
        d_.DataBind()
        Return d_
    End Function

    ''' <summary>
    ''' Binds DropDownList to List.
    ''' </summary>
    ''' <param name="d_"></param>
    ''' <param name="dont_if_full">Ignores the function if d_ already has items.</param>
    ''' <param name="object_">List.</param>
    ''' <returns>d_</returns>
    Private Shared Function DData(d_ As DropDownList, object_ As Object, Optional dont_if_full As Boolean = False) As DropDownList
        If dont_if_full And d_.Items.Count > 0 Then Return d_
        Try
            Dim l As New List(Of String)
            l = CType(object_, List(Of String))
            With d_
                .DataSource = l
                .DataBind()
            End With
        Catch
        End Try
        Return d_
    End Function

    Private Shared Function LData(l_ As ListBox, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing, Optional DontIfFull As Boolean = False) As ListBox
        If DontIfFull = True Then
            If l_.Items.Count > 0 Then Return l_
        End If
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            l_.DataSource = Nothing
        Catch ex As Exception

        End Try

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        l_.DataSource = dt
        l_.DataTextField = data_text_field
        If data_value_field.Length > 0 Then l_.DataValueField = data_value_field
        l_.DataBind()

        Return l_
    End Function

    ''' <summary>
    ''' Commits record to SQL Server database.
    ''' </summary>
    ''' <param name="query">The SQL query.</param>
    ''' <param name="connection_string">The server connection string.</param>
    ''' <param name="parameters_keys_values_">Values to put in table.</param>
    ''' <returns>True if successful, False if not.</returns>
    Private Function CommitRecord(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing) As Boolean
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = parameters_keys_values_
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
            Return False
        End Try

        '		Dim Entries_Insert As String = "INSERT INTO ENTRIES (EntryBy, ID, Category, [Description], Flag, [Title], Entry, DateAdded, TimeAdded, TitleID, Picture, PictureExtension, Topic) VALUES (@EntryBy, @ID, @Category, [@Description], @Flag, [@Title], @Entry, @DateAdded, @TimeAdded, @TitleID, @Picture, @PictureExtension, @Topic)"
        '		Dim entries_parameters_() = {"EntryBy", TitleBy.Text.Trim, "ID", EntryID.Text.Trim, "Category", Category.Text.Trim, "[Description]", Description.Text.Trim, "Flag", cFlag.Text.Trim, "[Title]", EntryTitle.Text.Trim, "Entry", NewEntry.Text.Trim, "DateAdded", date_, "TimeAdded", time_, "TitleID", TitleID.Text.Trim, "Picture", stream.GetBuffer(), "PictureExtension", PictureExtension.Text.Trim, "Topic", Topic.Text.Trim}
        '		d.CommitRecord(Entries_Insert, m_con, entries_parameters_)

    End Function

    ''' <summary>
    ''' Commits record to SQL Server database. Same as CommitRecord.
    ''' </summary>
    ''' <param name="query">The SQL query.</param>
    ''' <param name="connection_string">The server connection string.</param>
    ''' <param name="parameters_keys_values_">Values to put in table.</param>
    ''' <returns>True if successful, False if not.</returns>
    'Public Shared Function CommitSequel(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing) As Boolean
    '    Dim select_parameter_keys_values() = {}
    '    select_parameter_keys_values = parameters_keys_values_
    '    Try
    '        Dim insert_query As String = query
    '        Using insert_conn As New SqlConnection(connection_string)
    '            Using insert_comm As New SqlCommand()
    '                With insert_comm
    '                    .Connection = insert_conn
    '                    .CommandTimeout = 0
    '                    .CommandType = CommandType.Text
    '                    .CommandText = insert_query
    '                    If select_parameter_keys_values IsNot Nothing Then
    '                        For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
    '                            .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
    '                        Next
    '                    End If
    '                End With
    '                Try
    '                    insert_conn.Open()
    '                    insert_comm.ExecuteNonQuery()
    '                Catch ex As Exception
    '                End Try
    '            End Using
    '        End Using
    '        Return True
    '    Catch ex As Exception
    '        Return False
    '    End Try

    '    '		Dim Entries_Insert As String = "INSERT INTO ENTRIES (EntryBy, ID, Category, [Description], Flag, [Title], Entry, DateAdded, TimeAdded, TitleID, Picture, PictureExtension, Topic) VALUES (@EntryBy, @ID, @Category, [@Description], @Flag, [@Title], @Entry, @DateAdded, @TimeAdded, @TitleID, @Picture, @PictureExtension, @Topic)"
    '    '		Dim entries_parameters_() = {"EntryBy", TitleBy.Text.Trim, "ID", EntryID.Text.Trim, "Category", Category.Text.Trim, "[Description]", Description.Text.Trim, "Flag", cFlag.Text.Trim, "[Title]", EntryTitle.Text.Trim, "Entry", NewEntry.Text.Trim, "DateAdded", date_, "TimeAdded", time_, "TitleID", TitleID.Text.Trim, "Picture", stream.GetBuffer(), "PictureExtension", PictureExtension.Text.Trim, "Topic", Topic.Text.Trim}
    '    '		d.CommitRecord(Entries_Insert, m_con, entries_parameters_)

    'End Function

    'Public Shared Function PictureFromStream(picture_ As System.Drawing.Image, Optional file_extension As String = ".jpg") As Byte()
    '    Dim photo_ As System.Drawing.Image = picture_
    '    Dim stream_ As New IO.MemoryStream

    '    If photo_ IsNot Nothing Then
    '        Select Case LCase(file_extension)
    '            Case ".jpg"
    '                photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
    '            Case ".jpeg"
    '                photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
    '            Case ".png"
    '                photo_.Save(stream_, Imaging.ImageFormat.Png)
    '            Case ".gif"
    '                photo_.Save(stream_, Imaging.ImageFormat.Gif)
    '            Case ".bmp"
    '                photo_.Save(stream_, Imaging.ImageFormat.Bmp)
    '            Case ".tif"
    '                photo_.Save(stream_, Imaging.ImageFormat.Tiff)
    '            Case ".ico"
    '                photo_.Save(stream_, Imaging.ImageFormat.Icon)
    '            Case Else
    '                photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
    '        End Select
    '    End If
    '    Return stream_.GetBuffer
    'End Function

    Public Shared Function ExtractHeader(select_keys As String()) As List(Of String)
        Dim a As String() = select_keys
        Dim l As New List(Of String)
        Dim x As Integer
        For i As Integer = 0 To a.Length - 1
            If a(i).ToLower.Contains("as") Then
                For j As Integer = 1 To a(i).Length
                    If Mid(a(i), j, 4).ToLower = " as " Then
                        x = j
                        l.Add(a(i).Substring(x + 4).Replace("'", "").Trim)
                    End If
                Next
            Else
                l.Add(a(i))
            End If
        Next
        Return l
    End Function

    Public Shared Function RandomColor(l As List(Of Integer)) As String
        Dim r As Integer = RandomList(0, 256, l)
        Dim g As Integer = RandomList(0, 256, l)
        Dim b As Integer = RandomList(0, 256, l)
        Dim return_ As String = "rgb(" & r & ", " & g & ", " & b & ")"
        Return return_
    End Function

    Public Shared Function RandomColor(l As List(Of Integer), Optional alpha_ As Byte = Nothing) As String
        Dim a As Byte = alpha_
        If a < 0 Then a = 0
        If a > 1 Then a = 1

        Dim r As Integer = RandomList(0, 256, l)
        Dim g As Integer = RandomList(0, 256, l)
        Dim b As Integer = RandomList(0, 256, l)
        Dim return_ As String = "rgba(" & r & ", " & g & ", " & b & ", " & a & ")"
        Return return_
    End Function


    Public Shared Sub StatusDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
        If d_.Items.Count > 0 Then Exit Sub

        Clear(d_)

        If FirstIsEmpty Then d_.Items.Add("")

        Dim l As List(Of String) = StatusList()
        For i As Integer = 0 To l.Count - 1
            d_.Items.Add(l(i).ToString)
        Next
    End Sub

    Public Shared Sub EnableControl(div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional state_ As Boolean = True)
        div_.Visible = state_
    End Sub


    ''' <summary>
    ''' Populates DropDownList with numbers.
    ''' </summary>
    ''' <param name="d_">DropDownList to fill</param>
    ''' <param name="end_">Ending Number</param>
    ''' <param name="start_">Beginning Number</param>
    ''' <param name="firstItemIsEmpty"></param>
    Public Shared Sub NumberDrop(d_ As DropDownList, end_ As Integer, Optional start_ As Integer = 0, Optional firstItemIsEmpty As Boolean = False)
        If d_.Items.Count > 0 Then Exit Sub

        With d_.Items
            Try
                If firstItemIsEmpty = True And d_.DataSource Is Nothing Then .Add("")
            Catch
            End Try
            For i As Integer = start_ To end_
                .Add(i)
            Next
        End With
    End Sub
    ''' <summary>
    ''' Discontinued.
    ''' </summary>
    ''' <param name="c_"></param>
    Public Shared Sub ClearText(c_ As Object)
        Try
            If TypeOf c_ Is DropDownList Then CType(c_, DropDownList).Text = ""
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Sub Clear(c_ As Array, Optional default_string As String = "")
        For i As Integer = 0 To c_.Length - 1
            Clear(c_(i), default_string)
        Next
    End Sub
    Public Shared Sub Clear(c_ As Object, Optional default_string As String = "")

        Dim c As DropDownList
        Dim l As ListBox
        Dim t As TextBox
        Dim p As System.Web.UI.WebControls.Image
        Dim h As CheckBox
        Dim ht As HtmlInputText
        Dim hta As HtmlTextArea
        Dim html_select As HtmlSelect
        Dim g As GridView

        If TypeOf c_ Is GridView Then
            g = c_
            Try
                g.DataSource = Nothing
            Catch
            End Try
        End If

        If TypeOf c_ Is CheckBox Then
            h = c_
            h.Checked = False
        End If
        If TypeOf c_ Is DropDownList Then
            c = c_
            'Try
            '    c.DataSource = Nothing
            'Catch ex As Exception
            'End Try
            'c.Items.Clear()
            c.Text = default_string
        End If
        If TypeOf c_ Is ListBox Then
            l = c_
            Try
                l.DataSource = Nothing
            Catch ex As Exception
            End Try
            l.Items.Clear()
        End If
        If TypeOf c_ Is TextBox Then
            t = c_
            t.Text = default_string
        End If
        If TypeOf c_ Is System.Web.UI.WebControls.Image Then
            p = c_
            Try
                p.ImageUrl = Nothing
            Catch ex As Exception
            End Try
        End If
        If TypeOf c_ Is HtmlInputText Then
            ht = c_
            ht.Value = default_string
        End If
        If TypeOf c_ Is HtmlTextArea Then
            hta = c_
            hta.Value = default_string
        End If
        If TypeOf c_ Is HtmlSelect Then
            html_select = c_
            Try
                html_select.Items.Clear()
                html_select.Value = default_string
            Catch ex As Exception
            End Try
        End If
    End Sub
    Public Shared Sub Empty(c_ As Array, Optional default_string As String = "")
        For i As Integer = 0 To c_.Length - 1
            Empty(c_(i), default_string)
        Next
    End Sub
    Public Shared Sub Empty(c_ As Object, Optional default_string As String = "")

        Dim c As DropDownList
        Dim l As ListBox

        If TypeOf c_ Is DropDownList Then
            c = c_
            c.DataSource = Nothing
            c.Items.Clear()
            c.Text = default_string
        ElseIf TypeOf c_ Is ListBox Then
            l = c_
            Try
                l.DataSource = Nothing
            Catch ex As Exception
            End Try
            l.Items.Clear()
        End If
    End Sub
    ''' <summary>
    ''' Places text inside TextBox, ComboBox, Button, HTML DIV/SPAN or Label, converts date to short (and also for web if DIV/SPAN) alongside.
    ''' </summary>
    ''' <param name="str_"></param>
    ''' <param name="control_"></param>
    ''' <param name="convert_date_to_short"></param>
    ''' <returns></returns>
    Public Shared Shadows Function WriteText(str_ As Object, control_ As Object, Optional convert_date_to_short As Boolean = True) As String
        If TypeOf (str_) Is Date Or TypeOf (str_) Is DateTime Then

        End If

        Dim r As String = DateToShort(str_, convert_date_to_short)
        Dim t As TextBox, d As DropDownList, l As Label, b As Button, div_ As System.Web.UI.HtmlControls.HtmlGenericControl
        Try
            If TypeOf control_ Is TextBox Then
                t = control_
                t.Text = r
            ElseIf TypeOf control_ Is DropDownList Then
                d = control_
                d.Text = r
            ElseIf TypeOf control_ Is Label Then
                l = control_
                l.Text = r
            ElseIf TypeOf control_ Is Button Then
                b = control_
                b.Text = r
            ElseIf TypeOf control_ Is System.Web.UI.HtmlControls.HtmlGenericControl Then
                div_ = control_
                WriteContent(PrepareForIO(r), div_)
            End If
        Catch ex As Exception

        End Try
        Return r
    End Function


    Public Shared Function Content(f As FileUpload) As Object
        Return CType(f, FileUpload).FileBytes
    End Function

    'Public Shared Property casing__ As TextCase = TextCase.Capitalize
    ''' <summary>
    ''' Gets the text of TextBox, HtmlInputText, HtmlTextArea, HtmlSelect; checkstate of Checkbox, or selected item of ListBox or DropDownList. May be necessary to combine with If Not Page.IsPostBack...
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="casing"></param>
    ''' <returns></returns>
    Public Shared Function Content(control_ As Object, Optional casing As TextCase = TextCase.None) As String

        Try
            If TypeOf control_ Is TextBox Then
                Return TransformText(CType(control_, TextBox).Text, casing)
            ElseIf TypeOf control_ Is HtmlInputText Then
                Return TransformText(CType(control_, HtmlInputText).Value, casing)
            ElseIf TypeOf control_ Is HtmlTextArea Then
                Return CType(control_, HtmlTextArea).Value
            ElseIf TypeOf control_ Is DropDownList Then
                Return TransformText(CType(control_, DropDownList).SelectedItem.Text, casing)
            ElseIf TypeOf control_ Is HtmlSelect Then
                Return TransformText(CType(control_, HtmlSelect).Items.Item(control_.SelectedIndex).ToString, casing)
            ElseIf TypeOf control_ Is CheckBox Then
                Return CType(control_, CheckBox).Checked
            ElseIf TypeOf control_ Is ListBox Then
                Return TransformText(CType(control_, ListBox).SelectedItem.Text, casing)
            End If
        Catch
        End Try

    End Function

    ''' <summary>
    ''' Indicates if supported web control has content
    ''' </summary>
    ''' <param name="g_OR_d_OR_l">The web control (GridView or DropDownList or ListBox</param>
    ''' <returns>True if control has content, False if not.</returns>
    Public Shared Function HasData(g_OR_d_OR_l) As Boolean

        Dim g_ As GridView
        Dim d_ As DropDownList
        Dim l_ As ListBox

        If g_OR_d_OR_l IsNot Nothing And TypeOf (g_OR_d_OR_l) Is GridView Then
            g_ = g_OR_d_OR_l
            Return g_.Rows.Count > 0
        ElseIf g_OR_d_OR_l IsNot Nothing And TypeOf (g_OR_d_OR_l) Is DropDownList Then
            d_ = g_OR_d_OR_l
            Return d_.Items.Count > 0
        ElseIf g_OR_d_OR_l IsNot Nothing And TypeOf (g_OR_d_OR_l) Is ListBox Then
            l_ = g_OR_d_OR_l
            Return l_.Items.Count > 0
        End If
    End Function

    Public Shared Sub TitleDrop(c_ As DropDownList, Optional InitialSelectedIndexIsNegativeOne As Boolean = True)
        If c_.Items.Count > 0 Then Exit Sub
        Dim titles As List(Of String) = GetEnum(New Title)
        Clear(c_)
        With c_
            With .Items
                For i = 0 To titles.Count - 1
                    .Add(titles(i))
                Next
            End With
            c_.SelectedIndex = If(InitialSelectedIndexIsNegativeOne, -1, 0)
            .Text = ""
        End With

    End Sub

    Public Shared Sub GenderDrop(c_ As DropDownList, Optional InitialSelectedIndexIsNegativeOne As Boolean = True, Optional ReplaceUnderscoresWithSpace As Boolean = True)
        If c_.Items.Count > 0 Then Exit Sub
        Dim titles As List(Of String) = GetEnum(New Gender)
        Clear(c_)
        With c_
            With .Items
                For i = 0 To titles.Count - 1
                    .Add(If(ReplaceUnderscoresWithSpace, titles(i).Replace("_", " "), titles(i)))
                Next
            End With
            c_.SelectedIndex = If(InitialSelectedIndexIsNegativeOne, -1, 0)
            .Text = ""
        End With
    End Sub

    Enum ControlsToCheck
        All
        Any
    End Enum

    ''' <summary>
    ''' Determines if all controls have text.
    ''' </summary>
    ''' <param name="controls_">Controls</param>
    ''' <returns>True or False</returns>

    Private Shared Function IsEmpty(controls_ As Array, Optional controls_to_check As ControlsToCheck = ControlsToCheck.Any) As Boolean
        Dim counter_ As Integer = 0
        With controls_
            For i As Integer = 0 To .Length - 1
                If IsEmpty(controls_(i)) Then
                    counter_ += 1
                End If
            Next
        End With
        '		Return Val(counter_) = controls_.Length
        If controls_to_check = ControlsToCheck.All Then
            Return Val(counter_) = controls_.Length
        Else
            Return Val(counter_) > 0
        End If
    End Function

    ''' <summary>
    ''' Determines if control has text or list has items.
    ''' </summary>
    ''' <param name="c_"></param>
    ''' <param name="use_trim">Should content be trimmed before check?</param>
    ''' <returns></returns>
    Public Shared Function IsEmpty(c_ As Object, Optional use_trim As Boolean = True) As Boolean
        Dim d__ As DropDownList
        Dim t__ As TextBox
        Dim html_inputText As HtmlInputText
        Dim html_textAra As HtmlTextArea
        Dim html_select As HtmlSelect
        Dim a As Array
        Dim l_string As List(Of String)
        Dim l_object As List(Of Object)
        Dim l_integer As List(Of Integer)
        Dim l_double As List(Of Double)
        Dim c As Collection

        If TypeOf c_ Is Array Then
            a = c_
            Return a.Length < 1
        End If

        If TypeOf c_ Is List(Of String) Then
            l_string = c_
            Return l_string.Count < 1
        End If

        If TypeOf c_ Is List(Of Object) Then
            l_object = c_
            Return l_object.Count < 1
        End If

        If TypeOf c_ Is List(Of Integer) Then
            l_integer = c_
            Return l_integer.Count < 1
        End If

        If TypeOf c_ Is List(Of Double) Then
            l_double = c_
            Return l_double.Count < 1
        End If

        If TypeOf c_ Is Collection Then
            c = c_
            Return c.Count < 1
        End If

        If TypeOf c_ Is DropDownList Then
            d__ = c_
            If use_trim = True Then
                Return d__.Text.Trim.Length < 1
            Else
                Return d__.Text.Length < 1
            End If
        End If

        If TypeOf c_ Is TextBox Then
            t__ = c_
            If use_trim = True Then
                Return t__.Text.Trim.Length < 1
            Else
                Return t__.Text.Length < 1
            End If
        End If

        If TypeOf c_ Is HtmlInputText Then
            html_inputText = c_
            If use_trim = True Then
                Return html_inputText.Value.Trim.Length < 1
            Else
                Return html_inputText.Value.Length < 1
            End If
        End If

        If TypeOf c_ Is HtmlTextArea Then
            html_textAra = c_
            If use_trim = True Then
                Return html_textAra.Value.Trim.Length < 1
            Else
                Return html_textAra.Value.Length < 1
            End If
        End If

        If TypeOf c_ Is HtmlSelect Then
            html_select = c_
            Return html_select.Items.Item(html_select.SelectedIndex).ToString.Length < 1
        End If

        If TypeOf (c_) Is FileUpload Then
            If CType(c_, FileUpload).HasFile = False Or CType(c_, FileUpload).HasFiles = False Then
                Return True
            Else
                Return False
            End If
        End If

        If TypeOf c_ Is CheckBox Then Return CType(c_, CheckBox).Checked = False

    End Function

    ''' <summary>
    ''' Populates drop with drop-text version of boolean, depending on the pattern.
    ''' </summary>
    ''' <param name="d_">DropDownList</param>
    ''' <param name="firstItemIsEmpty">Should DropDownList's first item be empty?</param>
    ''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
    Private Shared Sub BooleanDrop(d_ As DropDownList, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = False)
        '		Dim format_window_web As New NFunctions
        With d_
            If .Items.Count > 0 Then Exit Sub
            With .Items
                If firstItemIsEmpty = True Then .Add("")
                .Add(BooleanToDropText(True, pattern_))
                .Add(BooleanToDropText(False, pattern_))
            End With
        End With
    End Sub

    Public Shared Sub BooleanDrop(d_ As DropDownList, Optional pattern_ As DropTextPattern = DropTextPattern.AlwaysNever, Optional firstItemIsEmpty As Boolean = False)
        With d_
            If .Items.Count > 0 Then Exit Sub
            With .Items
                If firstItemIsEmpty = True Then .Add("")
                .Add(BooleanToDropText(True, pattern_))
                .Add(BooleanToDropText(False, pattern_))
            End With
            Try
                .Text = ""
                .SelectedIndex = -1
            Catch ex As Exception

            End Try
        End With

    End Sub

    ''' <summary>
    ''' Sets the title of the page.
    ''' </summary>
    ''' <param name="regular_str">The title, if condition_ exists</param>
    ''' <param name="alternate_str">The title, if condition_ does not exist</param>
    ''' <param name="condition_">Determines if alternate_ should be used in place of regular_</param>
    ''' <returns></returns>
    Public Shared Function PageTitle(regular_str As String, alternate_str As String, condition_ As String) As String
        If condition_ IsNot Nothing Then
            Return regular_str
        Else
            Return alternate_str
        End If
    End Function

    Public Shared Function AddCells(grid_with_quantity_price_IF_POSSIBLE As GridView, Optional isCart As Boolean = True, Optional quantity_i As Integer = Nothing, Optional price_i As Integer = Nothing)
        Dim g As GridView = grid_with_quantity_price_IF_POSSIBLE
        Dim counter = 0
        Dim q = 0, p = 0, q_i = 0, p_i = 1
        If quantity_i <> Nothing Then q_i = quantity_i
        If price_i <> Nothing Then p_i = price_i
        With g
            If .Rows.Count < 1 Then Return 0
            For i = 0 To .Rows.Count - 1
                q = Val(.Rows(i).Cells(q_i).Text)
                p = Val(.Rows(i).Cells(p_i).Text)
                If isCart Then
                    counter += (q * p)
                Else
                    counter += q
                End If
            Next
        End With
        Return counter
    End Function

    Public Shared Function EnumDrop(drop As ListBox, enum__ As Object, Optional as_drop_not_list As Boolean = True) As Object
        Dim e = GetEnum(enum__)
        Dim l As New List(Of String)
        With e
            For i = 0 To .Count - 1
                l.Add(EnumToDrop(e(i)))
            Next
        End With

        l.Sort()

        With drop
            .DataSource = Nothing
            .Items.Clear()
            For i = 0 To l.Count - 1
                .Items.Add(l(i))
            Next
        End With

        If as_drop_not_list Then
            Return drop
        Else
            Return l
        End If
    End Function


#End Region

#Region "Bindings"


    ''' <summary>
    ''' Attaches List as DataSource to DropDownList or ListBox.
    ''' </summary>
    ''' <param name="TheControl">DropDownList or ListBox</param>
    ''' <param name="TheList">List</param>
    ''' <returns></returns>
    Public Shared Function BindProperty(TheControl As WebControl, TheList As List(Of String)) As WebControl
        If TypeOf TheControl Is DropDownList Then
            CType(TheControl, DropDownList).DataSource = TheList
            CType(TheControl, DropDownList).DataBind()
        ElseIf TypeOf TheControl Is ListBox Then
            CType(TheControl, ListBox).DataSource = TheList
            CType(TheControl, ListBox).DataBind()
        End If
        Return TheControl
    End Function

    ''' <summary>
    ''' Populates DropDownList with items of List. Won't do anything if DropDownList already has items.
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="list_"></param>
    ''' <param name="FirstItemIsEmpty"></param>
    ''' <returns></returns>
    Public Shared Function BindProperty(control_ As DropDownList, list_ As List(Of String), FirstItemIsEmpty As Boolean, Optional Top_Items_Are As Array = Nothing, Optional clear_before_fill As Boolean = True) As WebControl
        If control_.Items.Count > 0 Then Return control_

        If FirstItemIsEmpty Then control_.Items.Add(EmptyString)

        Dim l As New List(Of String)
        If Top_Items_Are IsNot Nothing Then
            If Top_Items_Are.Length > 0 Then
                With Top_Items_Are
                    For i = 0 To .Length - 1
                        l.Add(Top_Items_Are(i))
                    Next
                End With
            End If
        End If
        For i = 0 To CType(list_, List(Of String)).Count - 1
            l.Add(list_(i))
        Next
        Try
            If (clear_before_fill) And (Top_Items_Are Is Nothing) Then
                CType(control_, DropDownList).Items.Clear()
                CType(control_, DropDownList).DataSource = Nothing
            End If
        Catch ex As Exception

        End Try
        With l
            Try
                For i = 0 To .Count - 1
                    CType(control_, DropDownList).Items.Add(l(i))
                Next
            Catch ex As Exception

            End Try
        End With
        Return control_
    End Function

    ''' <summary>
    ''' Attaches Array as DataSource to DropDownList. Won't do anything if DropDownList already has items.
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="list_"></param>
    ''' <param name="Top_Items_Are">Optional list of items to prepend the list with</param>
    ''' <param name="FirstItemIsEmpty"></param>
    ''' <returns></returns>
    'Public Shared Function BindProperty(control_ As DropDownList, list_ As Array, Optional Top_Items_Are As Array = Nothing, Optional FirstItemIsEmpty As Boolean = True, Optional clear_before_fill As Boolean = True) As WebControl
    '    If control_.Items.Count > 0 Then Return control_

    '    If FirstItemIsEmpty Then control_.Items.Add(EmptyString)

    '    Dim l As New List(Of String)
    '    'If InitialSelectedIndexIsNegativeOne Then l.Add("")
    '    If Top_Items_Are IsNot Nothing Then
    '        If Top_Items_Are.Length > 0 Then
    '            With Top_Items_Are
    '                For i = 0 To .Length - 1
    '                    l.Add(Top_Items_Are(i))
    '                Next
    '            End With
    '        End If
    '    End If
    '    If TypeOf (list_) Is Array Then
    '        For i = 0 To CType(list_, Array).Length - 1
    '            l.Add(list_(i))
    '        Next
    '    End If

    '    Try
    '        If (clear_before_fill) And (Top_Items_Are Is Nothing) Then
    '            CType(control_, DropDownList).Items.Clear()
    '            CType(control_, DropDownList).DataSource = Nothing
    '        End If
    '    Catch ex As Exception

    '    End Try
    '    With l
    '        Try
    '            For i = 0 To .Count - 1
    '                CType(control_, DropDownList).Items.Add(l(i))
    '            Next
    '        Catch ex As Exception

    '        End Try
    '    End With
    '    Return control_
    'End Function



    ''' <summary>
    ''' Populates ListBox with items of List. Won't do anything if ListBox already has items.
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="list_"></param>
    ''' <returns></returns>
    Public Shared Function BindProperty(control_ As ListBox, list_ As List(Of String), Optional Top_Items_Are As Array = Nothing, Optional clear_before_fill As Boolean = False) As WebControl
        If control_.Items.Count > 0 Then Return control_
        Dim l As New List(Of String)
        If Top_Items_Are IsNot Nothing Then
            If Top_Items_Are.Length > 0 Then
                With Top_Items_Are
                    For i = 0 To .Length - 1
                        l.Add(Top_Items_Are(i))
                    Next
                End With
            End If
        End If
        For i = 0 To CType(list_, List(Of String)).Count - 1
            l.Add(list_(i))
        Next
        Try
            If clear_before_fill Then
                CType(control_, ListBox).Items.Clear()
                CType(control_, ListBox).DataSource = Nothing
            End If
        Catch ex As Exception
        End Try
        With l
            Try
                For i = 0 To .Count - 1
                    CType(control_, ListBox).Items.Add(l(i))
                Next
            Catch ex As Exception

            End Try
        End With
        Return control_
    End Function

    ''' <summary>
    ''' Attaches Array as DataSource to ListBox. Won't do anything if ListBox already has items.
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="list_"></param>
    ''' <returns></returns>
    'Public Shared Function BindProperty(control_ As ListBox, list_ As Array, Optional Top_Items_Are As Array = Nothing, Optional clear_before_fill As Boolean = True) As WebControl
    '    If control_.Items.Count > 0 Then Return control_
    '    Dim l As New List(Of String)
    '    If Top_Items_Are IsNot Nothing Then
    '        If Top_Items_Are.Length > 0 Then
    '            With Top_Items_Are
    '                For i = 0 To .Length - 1
    '                    l.Add(Top_Items_Are(i))
    '                Next
    '            End With
    '        End If
    '    End If
    '    If TypeOf (list_) Is Array Then
    '        For i = 0 To CType(list_, Array).Length - 1
    '            l.Add(list_(i))
    '        Next
    '    End If
    '    Try
    '        If clear_before_fill Then
    '            CType(control_, ListBox).Items.Clear()
    '            CType(control_, ListBox).DataSource = Nothing
    '        End If
    '    Catch ex As Exception
    '    End Try
    '    With l
    '        Try
    '            For i = 0 To .Count - 1
    '                CType(control_, ListBox).Items.Add(l(i))
    '            Next
    '        Catch ex As Exception

    '        End Try
    '    End With
    '    Return control_
    'End Function



    '''' <summary>
    '''' Attaches List(Of String) as DataSource to DropDownList
    '''' </summary>
    '''' <param name="control_"></param>
    '''' <param name="list_"></param>
    '''' <param name="First_Item_Is_Empty"></param>
    '''' <returns></returns>
    'Public Shared Function BindProperty(control_ As DropDownList, list_ As List(Of String), Optional First_Item_Is As String = "", Optional First_Item_Is_Empty As Boolean = True) As WebControl
    '	Dim l As New List(Of String)
    '	If First_Item_Is_Empty Then l.Add("")
    '	If First_Item_Is.Trim.Length > 1 Then l.Add(First_Item_Is)
    '	For i = 0 To CType(list_, List(Of String)).Count - 1
    '		l.Add(list_(i))
    '	Next
    '	Try
    '		CType(control_, DropDownList).Items.Clear()
    '		CType(control_, DropDownList).DataSource = Nothing
    '	Catch ex As Exception

    '	End Try
    '	With l
    '		Try
    '			For i = 0 To .Count - 1
    '				CType(control_, DropDownList).Items.Add(l(i))
    '			Next
    '		Catch ex As Exception

    '		End Try
    '	End With
    '	Return control_
    'End Function

    '''' <summary>
    '''' Attaches Array as DataSource to DropDownList
    '''' </summary>
    '''' <param name="control_"></param>
    '''' <param name="list_"></param>
    '''' <param name="First_Item_Is_Empty"></param>
    '''' <returns></returns>
    'Public Shared Function BindProperty(control_ As DropDownList, list_ As Array, Optional First_Item_Is As String = "", Optional First_Item_Is_Empty As Boolean = True) As WebControl
    '	Dim l As New List(Of String)
    '	If First_Item_Is_Empty Then l.Add("")
    '	If First_Item_Is.Trim.Length > 1 Then l.Add(First_Item_Is)
    '	If TypeOf (list_) Is Array Then
    '		For i = 0 To CType(list_, Array).Length - 1
    '			l.Add(list_(i))
    '		Next
    '	End If
    '	Try
    '		CType(control_, DropDownList).Items.Clear()
    '		CType(control_, DropDownList).DataSource = Nothing
    '	Catch ex As Exception

    '	End Try
    '	With l
    '		Try
    '			For i = 0 To .Count - 1
    '				CType(control_, DropDownList).Items.Add(l(i))
    '			Next
    '		Catch ex As Exception

    '		End Try
    '	End With
    '	Return control_
    'End Function



    '''' <summary>
    '''' Attaches List(Of String) as DataSource to ListBox
    '''' </summary>
    '''' <param name="control_"></param>
    '''' <param name="list_"></param>
    '''' <param name="First_Item_Is_Empty"></param>
    '''' <returns></returns>
    'Public Shared Function BindProperty(control_ As ListBox, list_ As List(Of String), Optional First_Item_Is As String = "", Optional First_Item_Is_Empty As Boolean = True) As WebControl
    '	Dim l As New List(Of String)
    '	If First_Item_Is_Empty Then l.Add("")
    '	If First_Item_Is.Trim.Length > 1 Then l.Add(First_Item_Is)
    '	For i = 0 To CType(list_, List(Of String)).Count - 1
    '		l.Add(list_(i))
    '	Next
    '	Try
    '		CType(control_, ListBox).Items.Clear()
    '		CType(control_, ListBox).DataSource = Nothing
    '	Catch ex As Exception
    '	End Try
    '	With l
    '		Try
    '			For i = 0 To .Count - 1
    '				CType(control_, ListBox).Items.Add(l(i))
    '			Next
    '		Catch ex As Exception

    '		End Try
    '	End With
    '	Return control_
    'End Function

    '''' <summary>
    '''' Attaches Array as DataSource to ListBox
    '''' </summary>
    '''' <param name="control_"></param>
    '''' <param name="list_"></param>
    '''' <param name="First_Item_Is_Empty"></param>
    '''' <returns></returns>
    'Public Shared Function BindProperty(control_ As ListBox, list_ As Array, Optional First_Item_Is As String = "", Optional First_Item_Is_Empty As Boolean = True) As WebControl
    '	Dim l As New List(Of String)
    '	If First_Item_Is_Empty Then l.Add("")
    '	If First_Item_Is.Trim.Length > 1 Then l.Add(First_Item_Is)
    '	If TypeOf (list_) Is Array Then
    '		For i = 0 To CType(list_, Array).Length - 1
    '			l.Add(list_(i))
    '		Next
    '	End If
    '	Try
    '		CType(control_, ListBox).Items.Clear()
    '		CType(control_, ListBox).DataSource = Nothing
    '	Catch ex As Exception
    '	End Try
    '	With l
    '		Try
    '			For i = 0 To .Count - 1
    '				CType(control_, ListBox).Items.Add(l(i))
    '			Next
    '		Catch ex As Exception

    '		End Try
    '	End With
    '	Return control_
    'End Function


    ''' <summary>
    ''' Binds database column/field to DropDownList.
    ''' </summary>
    ''' <param name="control_">Control</param>
    ''' <param name="query_">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters (Nothing, if not needed)</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <returns>control_</returns>
    Public Shared Function BindProperty(control_ As DropDownList, query_ As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional data_text_field As String = "", Optional ignore_if_control_has_items As Boolean = True) As WebControl
        Return DData(control_, query_, connection_string, data_text_field, data_text_field, select_parameter_keys_values_, ignore_if_control_has_items)
    End Function

    ''' <summary>
    ''' Binds database column/field to ListBox.
    ''' </summary>
    ''' <param name="control_">Control</param>
    ''' <param name="query_">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters (Nothing, if not needed)</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <returns>control_</returns>
    Public Shared Function BindProperty(control_ As ListBox, query_ As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional data_text_field As String = "", Optional ignore_if_control_has_items As Boolean = True) As WebControl
        Return LData(control_, query_, connection_string, data_text_field, data_text_field, select_parameter_keys_values_, ignore_if_control_has_items)
    End Function

#End Region


#Region "Statistics"



#End Region

#Region "Nationality"
    Public Sub CountriesDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
        If d_.Items.Count > 0 Then Exit Sub

        Clear(d_)
        Dim l As New List(Of String)
        If FirstIsEmpty Then l.Add("")
        l.AddRange(CountriesList())
        d_.DataSource = l
        d_.DataBind()
    End Sub

    Public Sub StatesOfNigeriaDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
        If d_.Items.Count > 0 Then Exit Sub

        Clear(d_)

        Dim l As New List(Of String)
        If FirstIsEmpty Then l.Add("")
        l.AddRange(StatesOfNigeriaList())
        d_.DataSource = l
        d_.DataBind()
    End Sub
    Public Sub LGAsDrop(con_string__ As String, d_ As DropDownList, Optional state_ As String = Nothing, Optional country_ As String = "Nigeria")
        If state_ IsNot Nothing Then
            BindProperty(d_, BuildSelectString_DISTINCT("MyAdminDropdownState", {"LGAsOfNigeria"}, {"Countries", "StatesOfNigeria"}), con_string__, {"Countries", country_, "StatesOfNigeria", state_}, "LGAsOfNigeria")
            Exit Sub
        End If

        BindProperty(d_, BuildSelectString_DISTINCT("MyAdminDropdownState", {"LGAsOfNigeria"}, Nothing), con_string__, Nothing, "LGAsOfNigeria")
    End Sub

#End Region


#Region "JavaScript"

    ''' <summary>
    ''' Gives feedback to user with popup. Use with Materialize CSS. result of path_to_js_file must be present. script tag referencing path_to_js_file must be present. Name of function in the file must be 'materializeToast'. UpdatePanel must be on page.
    ''' </summary>
    ''' <param name="str_">Message</param>
    ''' <param name="UpdatePanel">The UpdatePanel.</param>
    ''' <param name="script_manager_object">Instance of ScriptManager.</param>
    Public Shared Sub Toast(str_ As String, UpdatePanel As Object, script_manager_object As Object, Optional path_to_js_file As String = "scripts/custom.js")
        ''		script_manager_object.RegisterStartupScript(sender_, page_or_updatePanel.GetType(), "text", "fToast('" & str_.Replace("'", "") & "')", True)
        script_manager_object.RegisterStartupScript(UpdatePanel, UpdatePanel.GetType(), "Toast", "materializeToast('" & str_.Replace("'", "") & "')", True)
    End Sub

    'Public Shared Sub Toast(str_ As String, UpdatePanel As Object, script_manager_object As Object, Optional path_to_js_file As String = "scripts/custom.js")
    ''' <summary>
    ''' Runs a JS function.
    ''' </summary>
    ''' <param name="function_"></param>
    ''' <param name="arg_"></param>
    ''' <param name="page_or_updatePanel"></param>
    ''' <param name="script_manager_object"></param>
    ''' <param name="path_to_js_file"></param>
    Public Shared Sub JavaScript(function_ As String, arg_ As Object, page_or_updatePanel As Object, script_manager_object As Object, Optional path_to_js_file As String = "scripts/custom.js")
        Dim param_
        If TypeOf (arg_) Is String Then
            param_ = arg_.ToString.Replace("'", "")
        Else
            param_ = arg_
        End If
        script_manager_object.RegisterStartupScript(page_or_updatePanel, page_or_updatePanel.GetType(), "JavaScript", function_ & "('" & param_ & "')", True)
    End Sub

    'Public Shared Sub JavaScript(function_ As String, arg_ As Object, sender_ As Object, page_or_updatePanel As Object, instance_of_script_manager As Object, Optional placeholder_for_ref_to_js_file As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing, Optional relative_path_to_js_file As String = Nothing)
    '	Dim param_
    '	If TypeOf (arg_) Is String Then
    '		param_ = arg_.ToString.Replace("'", "")
    '	Else
    '		param_ = arg_
    '	End If
    '	If placeholder_for_ref_to_js_file IsNot Nothing And relative_path_to_js_file IsNot Nothing Then
    '		Dim placeholder_text As String = "<script type=""text/javascript"" src=" & relative_path_to_js_file & "></script>"
    '		placeholder_for_ref_to_js_file.InnerHtml = placeholder_text
    '		placeholder_for_ref_to_js_file.Visible = True
    '	End If
    '	instance_of_script_manager.RegisterStartupScript(sender_, page_or_updatePanel.GetType(), "text", function_ & "('" & param_ & "')", True)
    'End Sub

#End Region


#Region "Charts"

#Region "Pie"

    Public Shared Function PieChart(l_values As List(Of String), Optional l_colors As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Return Chart_Pie_String(l_values, l_colors, width, height)
    End Function
    Public Shared Function PieChart(l_values As List(Of String), l_labels As List(Of String), div As HtmlGenericControl, Optional legend As HtmlGenericControl = Nothing, Optional l_colors As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Dim s As String = Chart_Pie_String(l_values, l_colors, width, height)
        If div IsNot Nothing Then div.InnerHtml = s
        Return s
    End Function

    Private Shared Function Chart_Pie_String(l_values As List(Of String), Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Dim id = NewGUID()
        Dim s As String = "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """ style=""width: " & width & "px; height: " & height & "px""></canvas>
			<script>
				var pieData = ["
        s &= Chart_Pie_Variable(l_values, COLORS)
        s &= "];
					new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Pie(pieData);
			  </script>"
        Return s
    End Function

    Private Shared Function Chart_Pie_Variable(l_values As List(Of String), COLORS As List(Of String))
        Dim l As List(Of String) = l_values
        Dim s As String = "", color As String

        For i = 0 To l.Count - 1
            If COLORS IsNot Nothing Then
                color = COLORS(i)
            Else
                color = RandomColor(New List(Of Integer))
            End If
            s &= "{
						value: " & l(i) & ",
						color: """ & color & """
				  }"
            If i <> l.Count - 1 Then s &= ","
        Next
        Return s
    End Function
#End Region

#Region "Doughnut"

    Public Shared Function DoughnutChart(l_values As List(Of String), Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Return Chart_Doughnut_String(l_values, COLORS, width, height)
    End Function
    Public Shared Function DoughnutChart(l_values As List(Of String), div As HtmlGenericControl, Optional legend As HtmlGenericControl = Nothing, Optional l_labels As List(Of String) = Nothing, Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Dim s As String = Chart_Doughnut_String(l_values, COLORS, width, height)
        If div IsNot Nothing Then div.InnerHtml = s
        Return s
    End Function
    'Private Shared Function DoughnutChart(grid_with_values As GridView, Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
    '	Dim g As GridView = grid_with_values
    '	Dim l As New List(Of String)
    '	With g
    '		If .Rows.Count < 1 Then Return ""
    '		For i = 0 To .Rows.Count - 1
    '			l.Add(g.Rows(i).Cells(0).Text)
    '		Next
    '	End With
    '	Dim s As String = Chart_Doughnut_String(l, COLORS, width, height)
    '	Return s
    'End Function
    'Private Shared Function DoughnutChart(grid_with_values As GridView, div As HtmlGenericControl, Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
    '	Dim g As GridView = grid_with_values
    '	Dim l As New List(Of String)
    '	With g
    '		For i = 0 To .Rows.Count - 1
    '			l.Add(g.Rows(i).Cells(0).Text)
    '		Next
    '	End With
    '	Dim s As String = Chart_Doughnut_String(l, COLORS, width, height)
    '	If div IsNot Nothing Then div.InnerHtml = s
    '	Return s
    'End Function

    Private Shared Function Chart_Doughnut_String(l_values As List(Of String), Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Dim id = NewGUID()
        Dim s As String = "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """ style=""width: " & width & "px; height: " & height & "px""></canvas>
			<script>
				var doughnutData = ["
        s &= Chart_Doughnut_Variable(l_values, COLORS)
        s &= "];
					new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Doughnut(doughnutData);
			  </script>"
        Return s
    End Function
    Private Shared Function Chart_Doughnut_Variable(l_values As List(Of String), COLORS As List(Of String))
        Dim l As List(Of String) = l_values
        Dim s As String = "", color As String

        For i = 0 To l.Count - 1
            If COLORS IsNot Nothing Then
                color = COLORS(i)
            Else
                color = RandomColor(New List(Of Integer))
            End If
            s &= "{
						value: " & l(i) & ",
						color: """ & color & """
				  }"
            If i <> l.Count - 1 Then s &= ","
        Next
        Return s
    End Function
#End Region

#Region "Line"

    Public Shared Function LineChart(l_values As List(Of String), l_labels As List(Of String), Optional width As Integer = 400, Optional height As Integer = 300) As String
        Return Chart_Line_String(l_values, l_labels, width, height)
    End Function
    Public Shared Function LineChart(l_values As List(Of String), l_labels As List(Of String), div As HtmlGenericControl, Optional width As Integer = 400, Optional height As Integer = 300) As String
        Dim s As String = Chart_Line_String(l_values, l_labels, width, height)
        div.InnerHtml = s
        Return s
    End Function

    'Public Shared Function Line(grid_with_labels_values As GridView, Optional width As Integer = 400, Optional height As Integer = 300) As String
    '    Dim l_values As New List(Of String), l_labels As New List(Of String)
    '    Dim g As GridView = grid_with_labels_values
    '    With g
    '        For i = 0 To .Rows.Count - 1
    '            l_values.Add(.Rows(i).Cells(1).Text)
    '            l_labels.Add(.Rows(i).Cells(0).Text)
    '        Next
    '    End With
    '    Return Chart_Line_String(l_values, l_labels, width, height)
    'End Function
    'Public Shared Function Line(grid_with_labels_values As GridView, div As HtmlGenericControl, Optional width As Integer = 400, Optional height As Integer = 300) As String
    '    Dim l_values As New List(Of String), l_labels As New List(Of String)
    '    Dim g As GridView = grid_with_labels_values
    '    With g
    '        For i = 0 To .Rows.Count - 1
    '            l_values.Add(.Rows(i).Cells(1).Text)
    '            l_labels.Add(.Rows(i).Cells(0).Text)
    '        Next
    '    End With
    '    Dim s = Chart_Line_String(l_values, l_labels, width, height)
    '    If div IsNot Nothing Then div.InnerHtml = s
    '    Return s
    'End Function

    Private Shared Function Chart_Line_String(l_values As List(Of String), l_labels As List(Of String), Optional width As Integer = 400, Optional height As Integer = 300) As String
        Dim id = NewGUID()
        Dim s As String = "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """ style=""width: " & width & "px; height: " & height & "px""></canvas>
			<script>
				var lineChartData = {"
        s &= Chart_Line_Variable(l_values, l_labels)
        s &= "new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Line(lineChartData);
			  </script>"
        Return s
    End Function
    Private Shared Function Chart_Line_Variable(l_values As List(Of String), l_labels As List(Of String))
        Dim l_string As String = "labels : ["
        For i = 0 To l_labels.Count - 1
            l_string &= """" & l_labels(i) & """"
            If i <> l_labels.Count - 1 Then l_string &= ","
        Next
        Dim fc = RandomColor(New List(Of Integer))
        Dim c = RandomColor(New List(Of Integer))
        l_string &= "],"

        Dim v_string As String = "datasets: [
										{
											fillColor: """ & fc & """,
											strokeColor: """ & c & """,
											pointColor: """ & c & """,
											pointStrokeColor: ""#fff"",
											data: ["
        For j = 0 To l_values.Count - 1
            v_string &= l_values(j)
            If j <> l_values.Count - 1 Then v_string &= ","
        Next
        v_string &= "]}]};"
        Return l_string & v_string
    End Function

#End Region

#Region "Bar"

    ''' <summary>
    ''' l_dataset.Length must be equal to x_labels.Length.
    ''' </summary>
    ''' <param name="x_labels"></param>
    ''' <param name="l_dataset"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <returns></returns>
    Public Shared Function BarChart(x_labels As List(Of String), l_dataset As List(Of BarChartDataSet), Optional width As Integer = 400, Optional height As Integer = 300)
        Return Chart_Bar_String(x_labels, l_dataset, width, height)
    End Function
    Public Shared Function BarChart(x_labels As List(Of String), l_dataset As List(Of BarChartDataSet), div_for_chart As HtmlGenericControl, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square, Optional width As Integer = 400, Optional height As Integer = 300)
        Dim s = Chart_Bar_String(x_labels, l_dataset, width, height)

        Dim colors As New List(Of String)
        For i = 0 To l_dataset.Count - 1
            colors.Add(l_dataset(i).color_)
        Next

        If div_for_legend IsNot Nothing Then
            Dim legend_values As New List(Of String)
            For i = 0 To l_dataset.Count - 1
                legend_values.Add(l_dataset(i).legend_value)
            Next

            Legend(legend_values, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing And chart_title IsNot Nothing Then
            div_for_title.InnerText = chart_title
        End If

        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s
        Return s
    End Function

    Private Shared Function Chart_Bar_String(x_labels As List(Of String), l_dataset As List(Of BarChartDataSet), Optional width As Integer = 400, Optional height As Integer = 300) As String
        Dim id = NewGUID()
        Dim s As String = "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """ style=""width: " & width & "px; height: " & height & "px""></canvas>
        <script>
        	var barChartData = {"
        s &= Chart_Bar_Variable(x_labels, l_dataset)
        s &= "};
        		new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Bar(barChartData);
          </script>"

        Return s
    End Function
    Private Shared Function Chart_Bar_Variable(x_labels As List(Of String), datasets As List(Of BarChartDataSet))
        Dim s As String = "labels: ["
        For i = 0 To x_labels.Count - 1
            s &= """" & x_labels(i) & """"
            If i <> x_labels.Count - 1 Then
                s &= ","
            End If
        Next
        s &= "],
                            datasets: ["

        For i = 0 To datasets.Count - 1
            s &= "{fillColor: """ & datasets(i).color_ & """, data: ["
            For j = 0 To datasets(i).x_values_for_each_x_label.Count - 1
                s &= datasets(i).x_values_for_each_x_label(j)
                If j <> datasets(i).x_values_for_each_x_label.Count - 1 Then
                    s &= ","
                End If
            Next
            s &= "]}"
            If i <> datasets.Count - 1 Then
                s &= ","
            End If
        Next
        s &= "]"
        Return s
    End Function


#End Region

#Region "Progress Bars"

    '-------------------
    'Bars
    '-------------------
    'Private Enum OperationType
    '    Count
    '    MIN
    '    MAX
    '    Sum
    '    AVG

    'End Enum

    'Public Shared Function ProgressBar(label As String, value As Object, div As System.Web.UI.HtmlControls.HtmlGenericControl, Optional title_ As String = Nothing) As String
    '    Dim title__ As String = ""
    '    If title_ IsNot Nothing Then title__ = title_

    '    Dim r As String = "<div class=""home-progres-main"">
    '					<h3>" & title__ & "</h3>
    '				</div>
    '				<div class='bar_group'>"

    '    r &= "<div class='bar_group__bar thin' label='" & label & "' show_values='true' tooltip='true' value='" & value & "'></div>"
    '    r &= "</div>"
    '    WriteContent(r, div)
    '    Return r
    'End Function

    'Public Shared Function ProgressBar(labels_values As List(Of Object), div As System.Web.UI.HtmlControls.HtmlGenericControl, Optional title_ As String = Nothing) As String
    '    Dim title__ As String = ""
    '    If title_ IsNot Nothing Then title__ = title_

    '    Dim r As String = "<div class=""home-progres-main"">
    '					<h3>" & title__ & "</h3>
    '				</div>
    '				<div class='bar_group'>"

    '    With labels_values
    '        If .Count > 0 Then
    '            For i As Integer = 0 To .Count - 1 Step 2
    '                r &= "<div class='bar_group__bar thin' label='" & labels_values(i) & "' show_values='true' tooltip='true' value='" & labels_values(i + 1) & "'></div>"
    '            Next
    '        End If
    '    End With
    '    r &= "</div>"
    '    WriteContent(r, div)
    '    Return r
    'End Function

    'Private Shared Function BarsO(g As GridView, div As System.Web.UI.HtmlControls.HtmlGenericControl, table_ As String, connection_string As String, field_to_group As String, field_to_apply_function_on As String, function_ As OperationType, where_keys As Array, where_keys_values As Array, Optional title_ As String = Nothing) As String
    '    Dim title__ As String = ""
    '    If title_ IsNot Nothing Then title__ = title_

    '    Dim r As String = "<div class=""home-progres-main"">
    '					<h3>" & title__ & "</h3>
    '				</div>
    '				<div class='bar_group'>"

    '    Dim q As String
    '    Select Case function_
    '        Case OperationType.AVG
    '            q = BuildAVGString_GROUPED(table_, field_to_group, field_to_apply_function_on, where_keys)
    '        Case OperationType.Count
    '            q = BuildCountString_GROUPED(table_, field_to_apply_function_on, field_to_group, where_keys)
    '        Case OperationType.Sum
    '            q = BuildSumString_GROUPED(table_, field_to_group, field_to_apply_function_on, where_keys)
    '    End Select
    '    Display(g, q, connection_string, where_keys_values)
    '    With g
    '        If .Rows.Count > 0 Then
    '            .Visible = True
    '            For i As Integer = 0 To .Rows.Count - 1
    '                r &= "<div class='bar_group__bar thin' label='" & .Rows(i).Cells(0).Text & "' show_values='true' tooltip='true' value='" & ToCurrency(.Rows(i).Cells(1).Text) & "'></div>"
    '            Next
    '        End If
    '    End With
    '    r &= "</div>"
    '    WriteContent(r, div)
    '    Return r
    'End Function
#End Region

#Region "Pie From Queries"
    Public Shared Function PieCount(g As GridView, table As String, field_to_count As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildCountString_GROUPED(table, field_to_count, field_to_group, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Counts Of " & field_to_count
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieCountOverTime(g As GridView, table As String, field_to_count As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildCountString_GROUPED_BETWEEN(table, field_to_group, field_to_count, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Counts Of " & field_to_count & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieSum(g As GridView, table As String, field_to_group As String, field_to_sum As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildSumString_GROUPED(table, field_to_group, field_to_sum, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Sums Of " & field_to_sum
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieSumOverTime(g As GridView, table As String, field_to_sum As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query As String = BuildSumString_GROUPED_BETWEEN(table, field_to_group, field_to_sum, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Sums Of " & field_to_sum & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieAverage(g As GridView, table As String, field_to_apply_avg_on As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildAVGString_GROUPED(table, field_to_group, field_to_apply_avg_on, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Averages Of " & field_to_apply_avg_on
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieAverageOverTime(g As GridView, table As String, field_to_apply_avg_on As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query As String = BuildAVGString_GROUPED_BETWEEN(table, field_to_group, field_to_apply_avg_on, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Averages Of " & field_to_apply_avg_on & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Private Shared Function PieTop(g As GridView, table As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional rows_to_select As Long = 10, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.DESC, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildTopString_GROUPED(table, field_to_group, keys_, keys_values, rows_to_select, OrderByField, order_by)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Top " & rows_to_select & " " & field_to_group
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
#End Region

#Region "Doughnut From Queries"
    Public Shared Function DoughnutCount(g As GridView, table As String, field_to_count As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildCountString_GROUPED(table, field_to_count, field_to_group, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Counts Of " & field_to_count
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function
    Public Shared Function DoughnutCountOverTime(g As GridView, table As String, field_to_count As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildCountString_GROUPED_BETWEEN(table, field_to_group, field_to_count, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Counts Of " & field_to_count & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    Public Shared Function DoughnutSum(g As GridView, table As String, field_to_group As String, field_to_sum As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildSumString_GROUPED(table, field_to_group, field_to_sum, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Sums Of " & field_to_sum
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function
    Public Shared Function DoughnutSumOverTime(g As GridView, table As String, field_to_sum As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query As String = BuildSumString_GROUPED_BETWEEN(table, field_to_group, field_to_sum, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Sums Of " & field_to_sum & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function
    Public Shared Function DoughnutAverage(g As GridView, table As String, field_to_apply_avg_on As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildAVGString_GROUPED(table, field_to_group, field_to_apply_avg_on, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Averages Of " & field_to_apply_avg_on
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function
    Public Shared Function DoughnutAverageOverTime(g As GridView, table As String, field_to_apply_avg_on As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query As String = BuildAVGString_GROUPED_BETWEEN(table, field_to_group, field_to_apply_avg_on, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                div_for_title.InnerText = "Averages Of " & field_to_apply_avg_on & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

#End Region

#Region "Line From Queries"
    Public Shared Function Line(g As GridView, table As String, x_label_field As String, y_value_field As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing)
        g.Visible = False
        Dim query As String = BuildSelectString(table, {x_label_field, y_value_field}, keys_)
        Display(g, query, connection_string, keys_values)

        Dim l_values As New List(Of String), l_labels As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                l_labels.Add(.Rows(i).Cells(0).Text)
                l_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With
        Dim s = Chart_Line_String(l_values, l_labels)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s

        If div_for_title IsNot Nothing And chart_title IsNot Nothing Then
            div_for_title.InnerText = chart_title
        End If

        Return s
    End Function

    Public Shared Function Line(g As GridView, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing)
        g.Visible = False
        Dim l_values As New List(Of String), l_labels As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                l_labels.Add(.Rows(i).Cells(0).Text)
                l_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With
        Dim s = Chart_Line_String(l_values, l_labels)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s
        Return s
    End Function
#End Region

#Region "Members"
    Public Enum LegendMarkerStyle
        Circle
        Square
    End Enum
    Public Structure BarChartDataSet
        Public x_values_for_each_x_label As List(Of String)
        Public color_ As String
        Public legend_value As String
    End Structure

#End Region

#Region "Support Functions"
    Private Shared Function Legend(Labels As List(Of String), legend_colors As List(Of String), LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim legend_ As String = ""
        With legend_colors
            For i As Integer = 0 To .Count - 1
                If LegendMarkerStyle_ = LegendMarkerStyle.Circle Then
                    legend_ &= "<span style=""border-radius:50%; background-color:" & legend_colors(i).Replace("""", "") & """>&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;&nbsp;" & Labels(i)
                ElseIf LegendMarkerStyle_ = LegendMarkerStyle.Square Then
                    legend_ &= "<span style=""border-radius:20%; background-color:" & legend_colors(i).Replace("""", "") & """>&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;&nbsp;" & Labels(i)
                End If
                legend_ &= "<br />"
            Next
        End With
        legend_ &= "</div>"
        If LegendControl IsNot Nothing Then WriteContent(legend_, LegendControl)
        Return legend_
    End Function

#End Region


#End Region


#Region "IO - Bootstrap 3"
    Public Shared Sub AddControl(controls_ As Array, div As Object)
        With controls_
            If .Length > 0 Then
                For i As Integer = 0 To .Length - 1
                    AddControl(controls_(i), div)
                Next
            End If
        End With
    End Sub

    Public Shared Sub AddControl(control_ As Object, div As Object)
        Dim p As System.Web.UI.WebControls.PlaceHolder, d As System.Web.UI.HtmlControls.HtmlGenericControl
        If TypeOf div Is System.Web.UI.WebControls.PlaceHolder Then
            p = CType(div, System.Web.UI.WebControls.PlaceHolder)
            p.Controls.Add(control_)
        End If
        If TypeOf div Is System.Web.UI.HtmlControls.HtmlGenericControl Then
            d = CType(div, HtmlControls.HtmlGenericControl)
            d.Controls.Add(control_)
        End If

    End Sub

    ''' <summary>
    ''' Creates empty space.
    ''' </summary>
    ''' <param name="length_">lt 0 for 2 Line breaks.</param>
    ''' <returns></returns>
    Public Shared Function fix_str(Optional length_ As Integer = 5, Optional fix_ As Fix = Fix.LineBreak) As String
        If length_ < 0 Then Return "<br /><br />"

        Dim line_break As String = "<br /><br />"
        Dim paragraph_ As String = "<p>&nbsp;<br />&nbsp;</p>"

        Dim r_ As String
        Select Case fix_
            Case Fix.LineBreak
                r_ = line_break
            Case Fix.LineBreakWithParagraph
                r_ = paragraph_
        End Select

        Dim r As String
        For i As Integer = 1 To length_
            r &= r_
        Next
        Return r
    End Function

    'Public Shared Sub WriteContent(div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional length_ As Integer = 1, Optional content_type As Fix = Fix.LineBreak)
    '    WriteContent(fix_str(length_), div_, content_type)
    'End Sub

    Public Shared Function WriteContent(div_ As System.Web.UI.HtmlControls.HtmlGenericControl, str_ As String, Optional write_content As WriteContentType = WriteContentType.None, Optional WithClose As Boolean = True, Optional alert_is As AlertIs = AlertIs.warning, Optional format_for_web As Boolean = True)
        Return WriteContent(str_, div_, write_content, WithClose, alert_is, format_for_web)
    End Function
    Public Shared Function WriteContent(str_ As String, div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional write_content As WriteContentType = WriteContentType.None, Optional WithClose As Boolean = True, Optional alert_is As AlertIs = AlertIs.warning, Optional format_for_web As Boolean = True)
        If write_content = WriteContentType.Jumbotron Then Return Jumbotron(str_, div_)
        Dim fix_ As String = fix_str()
        Dim str__ As String = str_
        If format_for_web Then str_ = ToIO(str_)

        Try
            If write_content = WriteContentType.Bootstrap Then
                Alert(str__, div_, WithClose, alert_is)
            ElseIf write_content = WriteContentType.None Then
                div_.InnerHtml = str__
                div_.Visible = True
            ElseIf write_content = WriteContentType.Fix Then
                div_.InnerHtml = fix_
                div_.Visible = True
            End If
        Catch ex As Exception
        End Try
    End Function
    Public Shared Function Jumbotron(str__ As String, div__ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional size__ As Byte = 3) As String
        Dim size As Byte = 3
        If size__ < 1 Or size__ > 7 Then size = 7
        Dim str_ As String =
            "<div class=""jumbotron"">
                <h" & size & ">" & str__ & "</h" & size & ">
            </div>"
        div__.InnerHtml = str_
        Return str_
    End Function
    ''' <summary>
    ''' Gives feedback to user. Same as Feedback.
    ''' </summary>
    ''' <param name="str_"></param>
    ''' <param name="WithClose"></param>
    ''' <param name="alert_is"></param>
    ''' <returns></returns>
    Public Shared Function Alert(str_ As String, div As System.Web.UI.HtmlControls.HtmlGenericControl, Optional WithClose As Boolean = True, Optional alert_is As AlertIs = AlertIs.warning, Optional ShowElementAfterAlert As ShowElementAfterward = ShowElementAfterward.No) As String
        Try
            Dim r As String = New Web().Feedback(str_, WithClose, alert_is.ToString)
            div.InnerHtml = r
            If ShowElementAfterAlert = ShowElementAfterward.Yes Then
                div.Visible = True
            Else
                div.Visible = False
            End If
            Return r
        Catch
        End Try
    End Function

    ''' <summary>
    ''' Gives feedback to user. Constructed as an alert DIV.
    ''' </summary>
    ''' <param name="str_"></param>
    ''' <param name="WithClose"></param>
    ''' <param name="alert_OR_danger_OR_success_OR_warning"></param>
    ''' <returns></returns>
    Public Function Feedback(str_ As String, Optional WithClose As Boolean = True, Optional alert_OR_danger_OR_success_OR_warning As String = "warning") As String
        Select Case WithClose
            Case True
                Return FeedbackWithClose(str_, alert_OR_danger_OR_success_OR_warning)
            Case False
                Return FeedbackWithoutClose(str_, alert_OR_danger_OR_success_OR_warning)
        End Select



        '<span runat = "server" id="x"></span>
        'x.InnerHtml = f.Feedback(f.InvalidCredentialFeedback, True, "alert")
        'x.Visible = True

    End Function

    Public Function FeedbackWithoutClose(str_ As String, alert_OR_danger_OR_success_OR_warning As String) As String
        Dim header_ As String = "", footer_ As String = "</div>"
        Select Case alert_OR_danger_OR_success_OR_warning.ToLower
            Case "alert"
                header_ = "<div class=""alert alert-primary text-center"" role=""alert"">"
            Case "danger"
                header_ = "<div class=""alert alert-danger text-center"" role=""alert"">"
            Case "success"
                header_ = "<div class=""alert alert-success text-center"" role=""alert"">"
            Case "warning"
                header_ = "<div class=""alert alert-warning text-center"" role=""alert"">"
        End Select
        Return header_ & str_ & footer_
    End Function

    Public Function FeedbackWithClose(str_ As String, alert_OR_danger_OR_success_OR_warning As String) As String
        Dim header_ As String = "", footer_ As String = "</div>"
        Dim close_ As String = "<button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close""><span aria-hidden=""true"">&times;</span></button>"

        Select Case alert_OR_danger_OR_success_OR_warning.ToLower
            Case "alert"
                header_ = "<div class=""alert alert-primary alert-dismissible fade show"" role=""alert"">"
            Case "danger"
                header_ = "<div class=""alert alert-danger alert-dismissible fade show"" role=""alert"">"
            Case "success"
                header_ = "<div class=""alert alert-success alert-dismissible fade show"" role=""alert"">"
            Case "warning"
                header_ = "<div class=""alert alert-warning alert-dismissible fade show"" role=""alert"">"
        End Select

        Return header_ & str_ & close_ & footer_

    End Function

#End Region


#Region "HTTP"
    ''' <summary>
    ''' Quick http call.
    ''' </summary>
    ''' <param name="path_to_remote_resource"></param>
    ''' <returns>String.</returns>
    Public Shared Function Peek(path_to_remote_resource As String) As String
        Dim client As New WebClient()
        Dim stream As Stream = client.OpenRead(path_to_remote_resource)
        Dim reader As New StreamReader(stream)
        Dim content As String = reader.ReadToEnd()
        Return content
    End Function

#End Region

End Class
