Imports iNovation.Code.Web
Imports iNovation.Code.General
Public Class _Default
    Inherits Page

    Private Enum Professions
        Astronaut
        Pilot
        Engineer
    End Enum

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Dim arr As IEnumerable(Of String) = GetEnum(New Professions)
        If Not Page.IsPostBack Then
            BindProperty(GenderListBox, arr)
            BindProperty(GenderDropDown, arr)
        End If
    End Sub

    Protected Sub SubmitButton_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub _Default_Init(sender As Object, e As EventArgs) Handles Me.Init
    End Sub

    Protected Sub Unnamed_Click(sender As Object, e As EventArgs)
    End Sub
End Class