Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports iNovation.Code.Machine
Imports Microsoft.Win32
Imports iNovation.Code.Media
Imports System.Text.RegularExpressions
Imports iNovation.Code ''.CheckedDifference
Imports iNovation.Code.SearchEngineQueryString
Imports System.Collections.ObjectModel
Public Class Form2

    Private Enum Profession
        Doctor
        Pilot
        Engineer
        Writer
        Hairdresser
    End Enum

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TitleDrop(TitleDropDown)
        GenderDrop(GenderDropDown, False)
        BindProperty(ProfessionDropDown, GetEnum(New Profession), False)
    End Sub

    Private Sub ResetButton_Click(sender As Object, e As EventArgs) Handles ResetButton.Click
        Clear({TitleDropDown, NameTextBox, GenderDropDown, ProfessionDropDown, YearsOfExperienceTextBox})
    End Sub

    Private Sub SubmitButton_Click(sender As Object, e As EventArgs) Handles SubmitButton.Click
        SummaryTextBox.Text = Content(TitleDropDown) & vbCrLf & Content(NameTextBox) & vbCrLf & Content(GenderDropDown) & vbCrLf & Content(ProfessionDropDown) & vbCrLf & Content(YearsOfExperienceTextBox)
    End Sub

    Private Sub CopyToListBoxButton_Click(sender As Object, e As EventArgs) Handles CopyToListBoxButton.Click
        BindProperty(SummaryListBox, StringToList(Content(SummaryTextBox)))
    End Sub

    Private Sub UpButton_Click(sender As Object, e As EventArgs) Handles UpButton.Click
        mFeedback("Working", "this appears to be working")
    End Sub

    Private Sub RemoveButton_Click(sender As Object, e As EventArgs) Handles RemoveButton.Click

    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

    End Sub
End Class