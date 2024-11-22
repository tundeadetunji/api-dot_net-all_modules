Imports System.Runtime.CompilerServices
''' <summary>
''' This class contains extension methods based on methods from iNovation.Code.Feedback.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: November 2024
''' </remarks>
Public Module FeedbackExtensions
    Private ReadOnly Property f As New Feedback
    <Extension()>
    Public Sub Feedback(ByVal message As String, Optional ByVal async As Boolean = True)
        Inform(message, async)
    End Sub
    <Extension()>
    Public Sub Inform(ByVal message As String, Optional ByVal async As Boolean = True)
        f.Inform(message, async)
    End Sub
    <Extension()>
    Public Sub Stress(ByVal message As String, Optional ByVal how_many_times As Byte = 3, Optional ByVal async As Boolean = True)
        f.Stress(message, how_many_times, async)
    End Sub
    <Extension()>
    Public Sub Inform(ByVal messages As List(Of String), Optional ByVal async As Boolean = True)
        f.Inform(messages, async)
    End Sub
    <Extension()>
    Public Sub Stress(ByVal messages As List(Of String), Optional ByVal how_many_times As Byte = 3, Optional ByVal async As Boolean = True)
        f.Stress(messages, how_many_times, async)
    End Sub

End Module
