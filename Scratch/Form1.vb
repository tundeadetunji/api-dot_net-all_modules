Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports iNovation.Code.Machine
Imports Microsoft.Win32
Imports iNovation.Code.Media
Imports System.Text.RegularExpressions
Imports iNovation.Code ''.CheckedDifference
Imports iNovation.Code.SearchEngineQueryString
Imports System.Collections.ObjectModel

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
        Dim java As New SearchEngineQueryString.TermWithVariations("java")
        java.AddVariation("jar")
        java.AddVariation("war")
        Dim python As New SearchEngineQueryString.TermWithVariations("python", SearchStringOperator.AND_)
        python.AddVariation("perl")
        python.AddVariation("ruby")
        Dim terms As New List(Of TermWithVariations)
        terms.Add(java)
        terms.Add(python)
        Dim sites As New List(Of String)
        sites.Add("https://site.com")
        sites.Add("https://new.com")
        t.Text = constructQueryString(sites, terms)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
    End Sub
End Class
