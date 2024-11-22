Imports System.Runtime.CompilerServices
Imports System.Text

''' <summary>
''' This class contains extension methods useful for logging common data types.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: November 2024
''' </remarks>
Public Module LoggingExtensions
#Region "exported"
    <Extension()>
    Public Sub Log(ByVal i As Integer)
        Print(i)
    End Sub
    <Extension()>
    Public Sub Log(ByVal d As Double)
        Print(d)
    End Sub
    <Extension()>
    Public Sub Log(ByVal d As Decimal)
        Print(d)
    End Sub
    <Extension()>
    Public Sub Log(ByVal s As Short)
        Print(s)
    End Sub
    <Extension()>
    Public Sub Log(ByVal b As Byte)
        Print(b)
    End Sub
    <Extension()>
    Public Sub Log(ByVal l As Long)
        Print(l)
    End Sub
    <Extension()>
    Public Sub Log(ByVal d As DateTime)
        Print(d)
    End Sub
    <Extension()>
    Public Sub Log(Of E)(ByVal l As List(Of E))
        If l Is Nothing Then Return

        Debug.WriteLine("")
        Debug.WriteLine("Log for " & l.GetType().Name)
        For Each item As E In l
            Debug.WriteLine(item)
        Next
        Debug.WriteLine("")
    End Sub
    <Extension()>
    Public Sub Log(Of E)(ByVal l As IEnumerable(Of E))
        If l Is Nothing Then Return

        Debug.WriteLine("")
        Debug.WriteLine("Log for " & l.GetType().Name)
        For Each item As E In l
            Debug.WriteLine(item)
        Next
        Debug.WriteLine("")
    End Sub
    <Extension()>
    Public Sub Log(Of E)(ByVal s As HashSet(Of E))
        If s Is Nothing Then Return

        Debug.WriteLine("")
        Debug.WriteLine("Log for " & s.GetType().Name)
        For Each item As E In s
            Debug.WriteLine(item)
        Next
        Debug.WriteLine("")
    End Sub


    <Extension()>
    Public Sub Log(Of T, U)(ByVal dict As Dictionary(Of T, U))
        If dict Is Nothing Then Return
        If dict.Count = 0 Then Return
        Debug.WriteLine("Listing for " & dict.GetType().Name)
        For Each key As T In dict.Keys
            Dim builder As New StringBuilder
            Debug.WriteLine(builder.Append(key).Append(": ").Append(dict.Item(key)))
        Next
    End Sub
    <Extension()>
    Public Sub Log(ByVal s As String)
        Print(s)
    End Sub
    <Extension()>
    Public Sub Log(ByVal ex As Exception)
        Print(ex)
    End Sub
#End Region
#Region "private"

    Private Sub Print(Of T)(ByVal str1 As T)
        If str1 Is Nothing Then Return

        Debug.WriteLine("")
        Debug.WriteLine("Log for " & str1.GetType().Name)
        Debug.WriteLine(str1)
        Debug.WriteLine("")

    End Sub
#End Region

End Module
