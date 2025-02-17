Imports System.Windows.Forms
Imports iNovation.Code.Desktop
Public Class ProgramRunner

#Region "Properties"
    Private Property SubsequentPrograms As List(Of String) = New List(Of String)
    Private Property InitialProgram As String = String.Empty
    Private Property InitialInterval As Integer = 0
    Private Property SubsequentIntervals As Integer = 0
    Private Property InitialTimer As Timer
    Private Property SubsequentTimer As Timer
    Private SubsequentProgramRunIndex As Integer = 0

#End Region

#Region "Initialization"
    Public Sub New(initialProgram As String, subsequentPrograms As List(Of String), Optional initialInterval As Integer = 20, Optional subsequentIntervals As Integer = 5)
        Me.InitialProgram = initialProgram
        Me.InitialInterval = initialInterval * 1000
        Me.SubsequentPrograms = subsequentPrograms
        Me.SubsequentIntervals = subsequentIntervals * 1000
    End Sub
    Public Sub New(programs As List(Of String), Optional interval As Integer = 5)
        Me.SubsequentPrograms = programs
        Me.SubsequentIntervals = interval * 1000
    End Sub
#End Region


#Region "Support"
    Private Sub RunInitialProgram()
        If Not String.IsNullOrEmpty(InitialProgram) Then
            StartFile(InitialProgram)

            If SubsequentPrograms.Count > 0 Then
                InitialTimer = New Timer
                AddHandler InitialTimer.Tick, AddressOf InitialTimer_Tick
                InitialTimer.Interval = InitialInterval
                InitialTimer.Enabled = True
            End If
        End If

    End Sub

    Private Sub RunSubsequentPrograms()
        If SubsequentPrograms.Count = 0 Then Return

        SubsequentTimer = New Timer
        AddHandler SubsequentTimer.Tick, AddressOf SubsequentTimer_Tick
        SubsequentTimer.Interval = SubsequentIntervals
        SubsequentTimer.Enabled = True

    End Sub

#End Region

#Region "RunTimer Related"
    Private Sub InitialTimer_Tick(sender As Object, e As EventArgs)
        InitialTimer.Enabled = False
        RunSubsequentPrograms()
    End Sub

    Private Sub SubsequentTimer_Tick(sender As Object, e As EventArgs)
        SubsequentTimer.Enabled = False
        If SubsequentProgramRunIndex < SubsequentPrograms.Count Then
            If Not String.IsNullOrEmpty(SubsequentPrograms(SubsequentProgramRunIndex)) Then
                StartFile(SubsequentPrograms(SubsequentProgramRunIndex))
                SubsequentProgramRunIndex += 1
            End If
            SubsequentTimer.Enabled = True
        Else
            Return
        End If
    End Sub


#End Region

#Region "Exported"
    Public Sub Run()
        If Not String.IsNullOrEmpty(InitialProgram) Then
            RunInitialProgram()
        Else
            RunSubsequentPrograms()
        End If
    End Sub
#End Region

End Class
