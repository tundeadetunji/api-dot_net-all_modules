Imports System.Speech.Synthesis

''' <summary>
''' This class contains methods for Text-To-Speech, or otherwise, giving feedback. A good accompaniment to the UI directly viewed by the end-user.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public Class Feedback

#Region "init"
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

#Region "alt - Message"

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
            say(_message, _async)
        Next
    End Sub

#End Region

#Region "alt = Inform"
    ''' <summary>
    ''' Text to speech (same as say())
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="async"></param>
    Public Sub Inform(message As String, Optional async As Boolean = True)
        say(message, async)
    End Sub

    ''' <summary>
    ''' Text to speech + MessageBox
    ''' </summary>
    ''' <param name="message">Message in MessageBox</param>
    ''' <param name="voice_message">Message in TTS</param>
    ''' <param name="title"></param>
    ''' <param name="style"></param>
    ''' <param name="async"></param>
    ''' <returns></returns>
    Public Function Inform(message As String, voice_message As String, Optional title As String = "", Optional style As MsgBoxStyle = MsgBoxStyle.YesNo + MsgBoxStyle.Question, Optional async As Boolean = True) As MsgBoxResult
        Inform(voice_message, async)
        Dim msg = MsgBox(message, style, title)
        Return msg
    End Function
    ''' <summary>
    ''' Text to speech + MessageBox
    ''' </summary>
    ''' <param name="message">Message in MessageBox</param>
    ''' <param name="voice_messages">Message in TTS - picks one at random</param>
    ''' <param name="title"></param>
    ''' <param name="style"></param>
    ''' <param name="async"></param>
    ''' <returns></returns>
    Public Function Inform(message As String, voice_messages As List(Of String), Optional title As String = "", Optional style As MsgBoxStyle = MsgBoxStyle.YesNo + MsgBoxStyle.Question, Optional async As Boolean = True) As MsgBoxResult
        Inform(voice_messages(General.Random_(0, voice_messages.Count)), async)
        Dim msg = MsgBox(message, style, title)
        Return msg
    End Function
    ''' <summary>
    ''' Text to speech
    ''' </summary>
    ''' <param name="messages"></param>
    ''' <param name="async"></param>
    Public Sub Inform(messages As List(Of String), Optional async As Boolean = True)
        Inform(messages(General.Random_(0, messages.Count)), async)
    End Sub
    ''' <summary>
    ''' Text to speech, repeated until number_of_times is reached.
    ''' </summary>
    ''' <param name="messages"></param>
    ''' <param name="number_of_times"></param>
    ''' <param name="async"></param>
    Public Sub Stress(messages As List(Of String), number_of_times As Byte, Optional async As Boolean = True)
        Stress(messages(General.Random_(0, messages.Count)), number_of_times, async)
    End Sub

#End Region

#Region "Greeting"

#Region "Props"
    Private ReadOnly Property MorningGreeting As String = "Good morning"
    Private ReadOnly Property AfternoonGreeting As String = "Good afternoon"
    Private ReadOnly Property EveningGreeting As String = "Good evening"

    Private ReadOnly Property TimeFormat As String = "h:mm tt"
    Private ReadOnly Property MorningDateString As String = "12:00 am"
    Private ReadOnly Property AfternoonDateString As String = "12:00 pm"
    Private ReadOnly Property EveningDateString As String = "6:00 pm"
    Private ReadOnly Property MidnightDateString As String = "12:00 am"
    Private ReadOnly Property Morning As DateTime = DateTime.ParseExact(MorningDateString, TimeFormat, Nothing)
    Private ReadOnly Property Afternoon As DateTime = DateTime.ParseExact(AfternoonDateString, TimeFormat, Nothing)
    Private ReadOnly Property Evening As DateTime = DateTime.ParseExact(EveningDateString, TimeFormat, Nothing)
    Private ReadOnly Property Midnight As DateTime = DateTime.ParseExact(MidnightDateString, TimeFormat, Nothing)

#End Region

#Region "Structures"

    Private ReadOnly Property Greetings As New Dictionary(Of TimeOfDay, String) From {
        {TimeOfDay.Morning, MorningGreeting},
        {TimeOfDay.Afternoon, AfternoonGreeting},
        {TimeOfDay.Evening, EveningGreeting}
    }

    Enum TimeOfDay
        Morning = 1
        Afternoon = 2
        Evening = 3
    End Enum

#End Region

#Region "Routines"
    ''' <summary>
    ''' Text to speech as Greeting (English).
    ''' Morning is between midnight and noon.
    ''' Afternoon is between noon and 6pm.
    ''' Evening is between 6pm and 11.59pm.
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="whatElse"></param>
    Public Sub Greeting(Optional name As String = Nothing, Optional whatElse As String = Nothing)
        Dim f As Feedback = New Feedback
        f.say(Greetings.Item(TheTimeOfDay) & If(name IsNot Nothing, " " & name, "") & If(whatElse IsNot Nothing, vbCrLf & whatElse, ""))
    End Sub
    ''' <summary>
    ''' Text to speech as Greeting (Supply your own word/phrase)
    ''' </summary>
    ''' <param name="_morning">Greeting if it's morning (midnight to noon)</param>
    ''' <param name="_afternoon">Greeting if it's afternoon (noon to 6pm)</param>
    ''' <param name="_evening">Greeting if it's evening (6pm to 11:59pm)</param>
    ''' <param name="name"></param>
    ''' <param name="whatElse"></param>
    Public Sub GreetingNative(_morning As String, _afternoon As String, _evening As String, Optional name As String = Nothing, Optional whatElse As String = Nothing)
        Dim f As Feedback = New Feedback
        f.say(If(TheTimeOfDay() = TimeOfDay.Morning, _morning, If(TheTimeOfDay() = TimeOfDay.Afternoon, _afternoon, _evening)) & If(name IsNot Nothing, " " & name, "") & If(whatElse IsNot Nothing, vbCrLf & whatElse, ""))
    End Sub

#End Region

#Region "Utility"

    Private Function TheTimeOfDay() As TimeOfDay
        Dim now As DateTime = DateTime.ParseExact(DateTime.Now.ToShortTimeString, TimeFormat, Nothing)

        If now >= Morning And now < Afternoon Then
            Return TimeOfDay.Morning
        ElseIf now >= Afternoon And now < Evening Then
            Return TimeOfDay.Afternoon
        ElseIf now >= Evening And now < Midnight Then
            Return TimeOfDay.Evening
            'ElseIf now >= Midnight And now < Afternoon Then
            '    Return TimeOfDay.Morning
        Else
            Return TimeOfDay.Morning
        End If
    End Function

#End Region

#End Region


#Region "Core"
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

    Public Sub Stress(message As String, Optional howManyTimes As Byte = 3, Optional async As Boolean = True)
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

#Region "threading related"
    Private t__ As System.Windows.Forms.Timer
    Private message__ As String
    Private howManyTimes__ As Byte
    Private async__ As Boolean
    Private Sub remind()
        t__.Enabled = False
        Stress(message__, howManyTimes__, async__)
        t__.Enabled = True
    End Sub

#End Region

#Region "threading related"
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
