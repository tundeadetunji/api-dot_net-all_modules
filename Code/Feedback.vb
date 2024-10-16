Imports System.Speech.Synthesis

Public Class Feedback

#Region ""
    Private Property s As SpeechSynthesizer
    Public Sub New(Optional rate As Integer = -1, Optional volume As Byte = 100)
        s = New SpeechSynthesizer

        If rate = Nothing Then
            s.Rate = -1
        Else
            s.Rate = rate
        End If

        If volume = Nothing Or volume < 100 Or volume > 100 Then
            s.Volume = 100
        Else
            s.Volume = volume
        End If
    End Sub


#End Region

#Region ""

    Private _timer As System.Windows.Forms.Timer, _message As String, _how_many_times_ As Byte, _async As Boolean

    ''' <summary>
    ''' Gives audible feedback to the user (hence the naming of the class - Feedback). You can just use it in a Timer_Tick(sender, e) to repeat at the intervals.
    ''' </summary>
    ''' <param name="message_"></param>
    ''' <param name="how_many_times_">how many times to say the message</param>
    Public Sub MessageUser(message_ As String, Optional how_many_times_ As Byte = 3, Optional async As Boolean = True)
        _message = message_
        _how_many_times_ = how_many_times_
        _async = async
        Dim timer_ As New System.Windows.Forms.Timer
        _timer = timer_

        AddHandler timer_.Tick, AddressOf MessageUser
        _timer.Interval = 1
        _timer.Enabled = True
        MessageUser()
    End Sub

    Private Sub MessageUser()
        _timer.Enabled = False
        For i = 1 To _how_many_times_
            Say(_message, _async)
        Next
    End Sub

#End Region

#Region "Speech"
    Public Sub say(message As String, Optional async As Boolean = True)
        Try
            If async = True Then
                s.SpeakAsync(message)
            Else
                s.Speak(message)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Sub pause()
        Try
            If s.State = SynthesizerState.Speaking Then
                s.Pause()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Sub resumeSay()
        Try
            If s.State = SynthesizerState.Paused Then
                s.Resume()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Sub stress(message As String, Optional howManyTimes As Byte = 3, Optional async As Boolean = True)
        For i = 1 To howManyTimes
            Try
                If async = True Then
                    s.SpeakAsync(message)
                Else
                    s.Speak(message)
                End If

            Catch ex As Exception

            End Try
        Next
    End Sub
    Public Sub createReminder(message As String, interval_in_ms As Integer, Optional howManyTimes As Byte = 3, Optional async As Boolean = False)

        message__ = message
        howManyTimes__ = howManyTimes
        async__ = async

        t__ = New System.Windows.Forms.Timer
        t__.Interval = interval_in_ms
        AddHandler t__.Tick, AddressOf remind
        t__.Enabled = True
    End Sub

#End Region

#Region ""
    Private t__ As System.Windows.Forms.Timer
    Private message__ As String
    Private howManyTimes__ As Byte
    Private async__ As Boolean
    Private Sub remind()
        t__.Enabled = False
        stress(message__, howManyTimes__, async__)
        t__.Enabled = True
    End Sub

#End Region


#Region ""
    Private Property DoEvent_PlayAudio_OutTimer As System.Windows.Forms.Timer
    Private Property DoEvent_PlayAudio_file As String
    Private Property DoEvent_PlayAudio_mode As AudioPlayMode = AudioPlayMode.WaitToComplete
    Private Property DoEvent_PlayAudio_Dialog As System.Windows.Forms.Form
    Private thread_ As New System.Threading.Thread(AddressOf DoEvent_PlayAudio)
    Private thread_welcome As New System.Threading.Thread(AddressOf DoEvent_PlayAudio_Start)

    Public Sub welcome(Optional str As String = "Welcome", Optional sound_ As String = Nothing)

        ''Dim f As New Feedback()
        If str.Length > 0 Then say(str)

        If sound_ IsNot Nothing Then
            Try
                DoEvent_PlayAudio_file = sound_
                DoEvent_PlayAudio_mode = AudioPlayMode.Background
                thread_welcome.Start()
            Catch ex As Exception
            End Try
        End If


    End Sub

    ''' <summary>
    ''' Exits the application with sound. Call from Form_Closing or from desired event.
    ''' </summary>
    ''' <param name="dialog">Form.</param>
    ''' <param name="sound_">Path to sound file.</param>
    ''' <param name="str">What to say.</param>
    ''' <param name="dont_use_voice_feedback">Should str be said?</param>
    Public Sub bye(dialog As System.Windows.Forms.Form, str As String, sound_ As String, Optional dont_use_voice_feedback As Boolean = False)
        Dim out_timer As New System.Windows.Forms.Timer

        DoEvent_PlayAudio_Dialog = dialog
        If str.Length > 0 And dont_use_voice_feedback = False Then say(str, True)
        Try
            DoEvent_PlayAudio_OutTimer = out_timer
            AddHandler out_timer.Tick, New EventHandler(AddressOf OutTimer_Tick)
            DoEvent_PlayAudio_OutTimer.Enabled = True
        Catch
        End Try

        Try
            DoEvent_PlayAudio_file = sound_
        Catch ex As Exception
        End Try
        Try
            thread_.Start()
        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Exits the application without sound. Call from Form_Closing or from desired event.
    ''' </summary>
    ''' <param name="dialog">Form.</param>
    ''' <param name="str">What to say.</param>
    ''' <param name="dont_use_voice_feedback">Should str be said?</param>
    Public Sub bye(dialog As System.Windows.Forms.Form, str As String, Optional dont_use_voice_feedback As Boolean = False)
        Dim out_timer As New System.Windows.Forms.Timer
        DoEvent_PlayAudio_Dialog = dialog
        If str.Length > 0 And dont_use_voice_feedback = False Then say(str, False)
        Try
            DoEvent_PlayAudio_OutTimer = out_timer
            AddHandler out_timer.Tick, New EventHandler(AddressOf OutTimer_Exit_Tick)
            DoEvent_PlayAudio_OutTimer.Enabled = True
        Catch
        End Try
    End Sub
    Private Sub DoEvent_PlayAudio()
        Try
            My.Computer.Audio.Play(DoEvent_PlayAudio_file, AudioPlayMode.WaitToComplete)
        Catch
        End Try
        Environment.Exit(0)
    End Sub
    Private Sub DoEvent_PlayAudio_Start()
        Try
            My.Computer.Audio.Play(DoEvent_PlayAudio_file, AudioPlayMode.Background)
        Catch
        End Try

    End Sub
    Private Sub OutTimer_Tick()
        If DoEvent_PlayAudio_Dialog.Opacity <= 0 Then
            DoEvent_PlayAudio_OutTimer.Enabled = False
        End If
        DoEvent_PlayAudio_Dialog.Opacity -= 0.2
    End Sub
    Private Sub OutTimer_Exit_Tick()
        If DoEvent_PlayAudio_Dialog.Opacity <= 0 Then
            DoEvent_PlayAudio_OutTimer.Enabled = False
            Environment.Exit(0)
        End If
        DoEvent_PlayAudio_Dialog.Opacity -= 0.2
    End Sub

#End Region

End Class
