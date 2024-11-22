Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

''' <summary>
''' This class contains extension methods based on methods from iNovation.Code.Desktop.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: November 2024
''' </remarks>
Public Module DesktopExtensions
    <Extension()>
    Public Function ToList(ByVal l As ListBox) As List(Of String)
        Return Desktop.ListsItemsToList(l)
    End Function
    <Extension()>
    Public Function BindItems(ByVal l As ListBox, items As List(Of String)) As ListBox
        Return Desktop.BindProperty(l, items, False, False)
    End Function
    Public Sub ToTitleCase(ByVal t As TextBox)
        Desktop.ToTitleCase(t)
    End Sub
    <Extension()>
    Public Sub ToTitleCase(ByVal c As ComboBox)
        Desktop.ToTitleCase(c)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal b As Button)
        Desktop.EnableControl(b)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal c As CheckBox)
        Desktop.EnableControl(c)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal c As ComboBox)
        Desktop.EnableControl(c)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal l As ListBox)
        Desktop.EnableControl(l)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal n As NumericUpDown)
        Desktop.EnableControl(n)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal r As RadioButton)
        Desktop.EnableControl(r)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal r As RichTextBox)
        Desktop.EnableControl(r)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal t As TextBox)
        Desktop.EnableControl(t)
    End Sub
    <Extension()>
    Public Sub Enable(ByVal t As Timer)
        Desktop.EnableControl(t)
    End Sub
    Public Sub Disable(ByVal b As Button)
        Desktop.DisableControl(b)
    End Sub
    <Extension()>
    Public Sub Disable(ByVal c As CheckBox)
        Desktop.DisableControl(c)
    End Sub
    <Extension()>
    Public Sub Disable(ByVal c As ComboBox)
        Desktop.DisableControl(c)
    End Sub
    <Extension()>
    Public Sub Disable(ByVal l As ListBox)
        Desktop.DisableControl(l)
    End Sub
    <Extension()>
    Public Sub Disable(ByVal n As NumericUpDown)
        Desktop.DisableControl(n)
    End Sub
    <Extension()>
    Public Sub Disable(ByVal r As RadioButton)
        Desktop.DisableControl(r)
    End Sub
    <Extension()>
    Public Sub Disable(ByVal r As RichTextBox)
        Desktop.DisableControl(r)
    End Sub
    <Extension()>
    Public Sub Disable(ByVal t As TextBox)
        Desktop.DisableControl(t)
    End Sub
    <Extension()>
    Public Sub Disable(ByVal t As Timer)
        t.Enabled = False
    End Sub




    <Extension()>
    Public Function ToList(ByVal c As ComboBox) As List(Of String)
        Return Desktop.ListsItemsToList(c)
    End Function
    <Extension()>
    Public Function IsEmpty(ByVal c As ComboBox) As Boolean
        Return c.Items.Count < 1
    End Function
    <Extension()>
    Public Function IsEmpty(ByVal l As ListBox) As Boolean
        Return l.Items.Count < 1
    End Function
    <Extension()>
    Public Function IsEmpty(ByVal t As RichTextBox) As Boolean
        Return String.IsNullOrEmpty(t.Text)
    End Function
    <Extension()>
    Public Function IsEmpty(ByVal t As TextBox) As Boolean
        Return String.IsNullOrEmpty(t.Text)
    End Function
    <Extension()>
    Public Function BindItems(ByVal c As ComboBox, items As List(Of String)) As ListBox
        Return Desktop.BindProperty(c, items, False, False)
    End Function
    <Extension()>
    Public Sub Clear(ByVal c As ComboBox)
        Desktop.Clear(c)
    End Sub
    <Extension()>
    Public Sub Clear(ByVal l As ListBox)
        Desktop.Clear(l)
    End Sub
    <Extension()>
    Public Sub Clear(ByVal n As NumericUpDown)
        Desktop.Clear(n)
    End Sub
    <Extension()>
    Public Sub Clear(ByVal t As RichTextBox)
        Desktop.Clear(t)
    End Sub
    <Extension()>
    Public Sub Clear(ByVal t As TextBox)
        Desktop.Clear(t)
    End Sub
    <Extension()>
    Public Sub Clear(ByVal d As DataGridView)
        Desktop.Clear(d)
    End Sub
    <Extension()>
    Public Sub ToClipboard(ByVal s As String)
        If Not String.IsNullOrEmpty(s) Then Clipboard.SetText(s)
    End Sub
    <Extension()>
    Public Sub IncludeItem(ByVal l As ListBox, ByVal r As ListBox, Optional ByVal retain As Boolean = False)
        Desktop.ListsIncludeItem(l, r, retain)
    End Sub
    <Extension()>
    Public Sub IncludeItems(ByVal l As ListBox, ByVal r As ListBox)
        Desktop.ListsIncludeAllItems(l, r)
    End Sub
    <Extension()>
    Public Sub IncludeItem(ByVal l As ComboBox, ByVal r As ComboBox, Optional ByVal retain As Boolean = False)
        Desktop.ListsIncludeItem(l, r, retain)
    End Sub
    <Extension()>
    Public Sub IncludeItems(ByVal l As ComboBox, ByVal r As ComboBox)
        Desktop.ListsIncludeAllItems(l, r)
    End Sub


    <Extension()>
    Public Sub ExcludeItem(ByVal l As ListBox, ByVal r As ListBox)
        Desktop.ListsExcludeItem(l, r)
    End Sub
    <Extension()>
    Public Sub ExcludeItems(ByVal l As ListBox, ByVal r As ListBox)
        Desktop.ListsExcludeAllItems(l, r)
    End Sub
    <Extension()>
    Public Sub ExcludeItem(ByVal l As ComboBox, ByVal r As ComboBox)
        Desktop.ListsExcludeItem(l, r)
    End Sub
    <Extension()>
    Public Sub ExcludeItems(ByVal l As ComboBox, ByVal r As ComboBox)
        Desktop.ListsExcludeAllItems(l, r)
    End Sub
    <Extension()>
    Public Sub AddItem(ByVal c As ComboBox, ByVal i As Object, Optional ByVal allow_duplicate As Boolean = False)
        Desktop.ListsAddItem(c, i, allow_duplicate)
    End Sub
    <Extension()>
    Public Sub AddItem(ByVal l As ListBox, ByVal i As Object, Optional ByVal allow_duplicate As Boolean = False)
        Desktop.ListsAddItem(l, i, allow_duplicate)
    End Sub
    <Extension()>
    Public Sub RemoveItem(ByVal c As ComboBox)
        Desktop.ListsRemoveItem(c)
    End Sub
    <Extension()>
    Public Sub RemovItem(ByVal l As ListBox)
        Desktop.ListsRemoveItem(l)
    End Sub




End Module
