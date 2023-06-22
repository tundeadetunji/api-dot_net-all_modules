Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports iNovation.Code.Machine
Imports Microsoft.Win32
Imports iNovation.Code.Media
Imports System.Text.RegularExpressions


Public Class Form1

    Dim ftp_server = "ftp.radianthope.org.ng"
    Dim username = "lbtwieye"
    Dim password = "h568]ImZFq7;zG"
    Dim remoteFolderPathEscapedString = "dev.radianthope.org.ng\wwwroot\images\"
    Dim dest_folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\iNovation Digital Works"

    Dim table As String = "pie"
    Dim field_to_count As String = "quater"
    Dim x_label_field As String = field_to_count
    Dim field_to_group As String = field_to_count
    Dim field_for_interval As String = "id"
    Dim interval_from = 2
    Dim interval_to = 6
    Dim field_to_sum = "drink"
    Dim field_to_average = field_to_sum
    Dim y_value_field = field_to_sum
    Dim where_keys As Array = Nothing
    Dim keys_ As Array = where_keys

    Private s As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'PieCount
        '' field_to_count = quater: first(2), second(2), third(3)

        Dim query = BuildCountString_GROUPED(table, field_to_count, field_to_group, keys_)
        t.Text = query
        Clipboard.SetText(query)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'PieCountAcrossInterval
        '' field_to_count = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(1), second(2), third(2)

        Dim query = BuildCountString_GROUPED_BETWEEN(table, field_to_group, field_to_count, {field_for_interval})

        t.Text = query
        Clipboard.SetText(query)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'PieSum
        '' field_to_group = quater, field_to_sum = drink: first(75), second(155), third(95)
        Dim query = BuildSumString_GROUPED(table, field_to_group, field_to_sum, keys_)

        t.Text = query
        Clipboard.SetText(query)

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'PieSumAcrossInterval
        '' field_to_sum = drink, field_to_group = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(25), second(155), third(65)
        Dim query As String = BuildSumString_GROUPED_BETWEEN(table, field_to_group, field_to_sum, {field_for_interval})
        t.Text = query
        Clipboard.SetText(query)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'PieAverage
        '' field_to_group = quater, field_to_average = drink: first(37), second(77), third(31)
        Dim query = BuildAVGString_GROUPED(table, field_to_group, field_to_average, keys_)
        t.Text = query
        Clipboard.SetText(query)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'PieAverageAcrossInterval
        '' field_to_average = drink, field_to_group = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(25), second(77), third(32)
        Dim query As String = BuildAVGString_GROUPED_BETWEEN(table, field_to_group, field_to_average, {field_for_interval})
        t.Text = query
        Clipboard.SetText(query)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Line
        '' x_label_field: quater, y_value_field: drink, x_label_field_caption: what should be denoted as the 'topic/subject' in the title, e.g. the column name of y_value_field (in this case, "drink")
        Dim query As String = BuildSelectString(table, {x_label_field, y_value_field}, keys_)
        t.Text = query
        Clipboard.SetText(query)

    End Sub
End Class
