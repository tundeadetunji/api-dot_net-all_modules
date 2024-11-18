Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Module DesktopExtensions
#Region "ListBox"
    <Extension()>
    Public Function ToList(ByVal l As ListBox) As List(Of String)
        Return Desktop.ListsItemsToList(l)
    End Function
    <Extension()>
    Public Function BindItems(ByVal l As ListBox, items As List(Of String)) As ListBox
        Return Desktop.BindProperty(l, items, False, False)
    End Function

#End Region

#Region "ComboBox"
    <Extension()>
    Public Function ToList(ByVal l As ComboBox) As List(Of String)
        Return Desktop.ListsItemsToList(l)
    End Function
    <Extension()>
    Public Function BindItems(ByVal l As ComboBox, items As List(Of String)) As ListBox
        Return Desktop.BindProperty(l, items, False, False)
    End Function

#End Region

End Module
