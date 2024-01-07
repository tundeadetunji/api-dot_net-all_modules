Imports AForge.Video
Imports AForge.Video.DirectShow
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Drawing.Imaging

Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports iNovation.Code.Media

Public Class CameraAndMic

#Region "Fields"
    Private Const default_audio_extension As String = ".wav"
    Public Property isCapturing As Boolean = False

    'Private Shared Property camWidth As Short = 640
    'Private Shared Property camHeight As Short = 360
    'Private Shared ReadOnly Property camSize As Size = New Size(640, 360)

    Private ext__ As String = ".jpg"
    Private webcam As New FilterInfoCollection(Guid.NewGuid)
    Private cam As New VideoCaptureDevice
    Private devices_ As New List(Of String)
    Private selected_device As String
    Private timer_ As Timer
    Private image_control_ As PictureBox
    Private returned As New List(Of Image)
    Private counter As Long = 0

    Private capture__ As DeviceCapture = DeviceCapture.SingleImage

    Private _number_of_seconds_to_capture As Integer
    Private timer_to_stop_capture As Timer


    Public Property directory_ As String

    Private lingerImage As Boolean = False

    Private _ImagesCaptured As List(Of Image)
    Private _AudioCapturedFile As String

    Private Declare Function record Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

#End Region

#Region "Properties"

    ''' <summary>
    ''' Returns the images captured (DeviceCapture.Video)
    ''' 
    ''' Gba àwọn àwòrán.
    ''' </summary>
    ''' <returns></returns>

    Public ReadOnly Property CapturedImages As List(Of Image)
        Get
            Return returned
        End Get
    End Property

    ''' <summary>
    ''' Returns the image captured (DeviceCapture.SingleImage) 
    ''' 
    ''' Gba àwòrán.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property CapturedImage As Image
        Get
            If returned IsNot Nothing And returned.Count > 0 Then
                Return returned(0)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property CapturedAudio As String
        Get
            Return _AudioCapturedFile
        End Get
    End Property


#End Region

#Region "Initialization"

    ''' <summary>
    ''' Creates an instance. Ensure that directory__ contains nothing, to avoid erroneous results.
    ''' For optimum results, use different instances (and different image_control__ - aka PictureBox - as appropriate) for capturing, playback and view.
    ''' Recommended: use one and only one instance for capturing, but one or more for playback and view.
    ''' For optimum results, prefer playback/viewing from folder path (or directory__) - not from List(Of Images) or CapturedImages or CapturedImage.
    ''' 
    ''' Ṣe ètò àti ya àwòrán pèlú ẹ̀rọ. Tí ǹkan bá wà ní inúu directory__, ǹkan náà lè lòdì sí ètò yìí, tàbí kó paradà.
    ''' </summary>
    ''' <param name="prefer">Capture single image, video (series of images alongside audio), screen single image, screen video, or audio. 
    ''' If set to video, screen video or audio, then you must call EndCapture() manually except if number_of_seconds_to_capture is more than 0.
    ''' </param>
    ''' <param name="number_of_seconds_to_capture">Specify number of seconds to capture if prefer is set to Video.</param>
    ''' <param name="image_control__">To preview captured images.</param>
    ''' <param name="directory__">Where to save images in specified format (i.e. ext).</param>
    ''' <param name="ext">Extension to save image files in - e.g, ".jpg". For audio, it defaults to ".wav".</param>
    ''' <param name="linger_image__">Should image_control__ go blank after capturing.</param>
    ''' <example>
    ''' Private camera As CameraAndMic
    ''' camera = New CameraAndMic(DeviceCapture.SingleImage, "target_folder_path")
    ''' camera.StartCapture() ... camera.StopCapture()
    ''' to get caputred_images (DeviceCapture.Video), get the property CapturedImages
    ''' to get caputred_image (DeviceCapture.SingleImage), get the property CapturedImage
    ''' </example>

    Public Sub New(Optional prefer As DeviceCapture = DeviceCapture.SingleImage, Optional directory__ As String = "", Optional image_control__ As Object = Nothing, Optional number_of_seconds_to_capture As Integer = 0, Optional ext As String = ".jpg", Optional linger_image__ As Boolean = False, Optional directory_should_no_longer_be_visible As HideFileOrFolderShould = HideFileOrFolderShould.LeaveAsIs)
        If Exists(directory__) Then
            HideFileOrFolder(directory__, directory_should_no_longer_be_visible)
        End If

        If prefer = DeviceCapture.ScreenMotionPicture Or capture__ = DeviceCapture.ScreenVideo Or prefer = DeviceCapture.ScreenSingleImage Then
            InitializeForScreen(prefer, number_of_seconds_to_capture, image_control__, directory__, ext, linger_image__)
            Return
        End If

        If prefer = DeviceCapture.Audio Then
            capture__ = DeviceCapture.Audio
            InitializeForAudio(number_of_seconds_to_capture, directory__)
            Return
        End If

        Dim image_control As New PictureBox
        If image_control__ IsNot Nothing Then image_control = image_control__

        Try
            image_control.BackgroundImage = Nothing
        Catch ex As Exception

        End Try
        Try
            image_control.Image = Nothing
        Catch ex As Exception

        End Try

        Dim timer__ As Timer = New Timer

        capture__ = prefer
        directory_ = directory__
        image_control_ = image_control
        timer_ = timer__
        lingerImage = linger_image__
        _number_of_seconds_to_capture = number_of_seconds_to_capture

        ext__ = ext

        If number_of_seconds_to_capture > 0 Then
            timer_to_stop_capture = New Timer
            With timer_to_stop_capture
                Try
                    .Interval = number_of_seconds_to_capture * 1000
                Catch ex As Exception

                End Try
                AddHandler .Tick, New EventHandler(AddressOf timer_to_stop_video_or_screen_capture_Tick)
            End With

        End If

        If capture__ = DeviceCapture.SingleImage Or capture__ = DeviceCapture.MotionPicture Or capture__ = DeviceCapture.Video Then
            InitializeForEverythingNotScreenRelated()
            'ElseIf capture__ = DeviceCapture.Video Then
            '    InitializeForVideo()
        End If

    End Sub

    Private Sub InitializeForScreen(Optional prefer As DeviceCapture = DeviceCapture.SingleImage, Optional number_of_seconds_to_capture As Integer = 0, Optional image_control__ As Object = Nothing, Optional directory__ As String = "", Optional ext As String = ".jpg", Optional linger_image__ As Boolean = False)
        Dim image_control As New PictureBox
        If image_control__ IsNot Nothing Then image_control = image_control__

        Try
            image_control.BackgroundImage = Nothing
        Catch ex As Exception

        End Try
        Try
            image_control.Image = Nothing
        Catch ex As Exception

        End Try

        Dim timer__ As Timer = New Timer

        capture__ = prefer
        directory_ = directory__
        image_control_ = image_control
        timer_ = timer__
        lingerImage = linger_image__
        _number_of_seconds_to_capture = number_of_seconds_to_capture

        ext__ = ext

        If number_of_seconds_to_capture > 0 Then
            timer_to_stop_capture = New Timer
            With timer_to_stop_capture
                Try
                    .Interval = number_of_seconds_to_capture * 1000
                Catch ex As Exception

                End Try
                AddHandler .Tick, New EventHandler(AddressOf timer_to_stop_video_or_screen_capture_Tick)
            End With

        End If


        With timer_
            .Interval = 20
            AddHandler .Tick, New EventHandler(AddressOf ScreenTimer_Tick)
        End With


    End Sub

    Private Sub InitializeForEverythingNotScreenRelated()
        webcam = New FilterInfoCollection(FilterCategory.VideoInputDevice)
        For Each VideoCaptureDevice As FilterInfo In webcam
            devices_.Add(VideoCaptureDevice.Name)
        Next
        If devices_.Count < 1 Then Return
        selected_device = devices_(0)

        With timer_
            .Interval = 20
            AddHandler .Tick, New EventHandler(AddressOf ImageAndMotionPictureAndVideoTimer_Tick)
        End With

        cam = New VideoCaptureDevice(webcam.Item(0).MonikerString)
        AddHandler cam.NewFrame, New NewFrameEventHandler(AddressOf cam_NewFrame)
        cam.Start()

    End Sub

    Private Sub InitializeForAudio(Optional number_of_seconds_to_capture As Integer = 0, Optional directory__ As String = "")

        directory_ = directory__
        _number_of_seconds_to_capture = number_of_seconds_to_capture

    End Sub
#End Region

#Region "Audio"
    Private Sub EndAudioCapture()
        Dim filename As String = directory_ & "\" & String.Format("{0:00}", CStr(Now.Month)) & "_" & String.Format("{0:00}", CStr(Now.Day)) & "_" & String.Format("{0:0000}", CStr(Now.Year)) & "_T_" & String.Format("{0:00}", CStr(Now.Hour)) & "_" & String.Format("{0:00}", CStr(Now.Minute)) & "_" & String.Format("{0:00}", CStr(Now.Millisecond)) & default_audio_extension ''.Replace("*", "")
        'Dim filename As String = directory_ & "\" & String.Format("{0:00}", MonthString(CStr(Now.Month))) & "_" & String.Format("{0:00}", CStr(Now.Day)) & "_" & String.Format("{0:0000}", CStr(Now.Year)) & "_at_" & LeadingZero(String.Format("{0:00}", CStr(Now.Hour))) & "_" & LeadingZero(String.Format("{0:00}", CStr(Now.Minute))) & "_" & String.Format("{0:00}", LeadingZero(CStr(Now.Millisecond))) & ext__.Replace("*", "")

        record("save recsound " & filename, "", 0, 0)
        record("close recsound", "", 0, 0)

        _AudioCapturedFile = filename

    End Sub
#End Region

#Region "Screen"
    Private Sub ScreenTimer_Tick()
        Try
            Dim captureBitmap As Bitmap = New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppArgb)
            Dim captureRectangle As Rectangle = Screen.AllScreens(0).Bounds
            Dim captureGraphics As Graphics = Graphics.FromImage(captureBitmap)
            captureGraphics.CopyFromScreen(captureRectangle.Left, captureRectangle.Top, 0, 0, captureRectangle.Size)

            returned.Add(captureBitmap)

            If image_control_ IsNot Nothing Then
                CType(image_control_, PictureBox).Image = captureBitmap
                'CType(image_control_, PictureBox).BackgroundImage = captureBitmap
            End If


            If directory_.Length > 0 Then
                Try
                    MkDir(directory_)
                Catch ex As Exception

                End Try
            End If
            counter += 1
            If directory_.Length > 0 Then
                Try
                    captureBitmap.Save(directory_ & "\" & LeadingZero(counter) & ext__, ImageFormat.Jpeg)
                Catch ex As Exception

                End Try

            End If


        Catch ex As Exception

        End Try


        If capture__ = DeviceCapture.ScreenSingleImage Then
            timer_.Enabled = False
            StopCapture()
        End If
    End Sub

#End Region

#Region "Shared"
    Private Sub timer_to_stop_video_or_screen_capture_Tick()
        Try
            If timer_to_stop_capture IsNot Nothing Then
                timer_to_stop_capture.Enabled = False
            End If
        Catch ex As Exception

        End Try

        StopCapture()

        Try
            If cam.IsRunning Then cam.Stop()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Stops capturing.
    ''' </summary>
    Public Sub StopCapture()
        isCapturing = False

        Try
            If timer_to_stop_capture IsNot Nothing Then
                timer_to_stop_capture.Enabled = False
            End If
        Catch ex As Exception

        End Try

        If capture__ = DeviceCapture.Audio Then
            EndAudioCapture()
            Return
        End If

        If capture__ = DeviceCapture.Video Or capture__ = DeviceCapture.ScreenVideo Then
            EndAudioCapture()
        End If

        Try
            timer_.Enabled = False
        Catch
        End Try

        Try
            If cam.IsRunning Then
                cam.Stop()
            End If
        Catch
        End Try

        Try
            If lingerImage = False Then image_control_.Image = Nothing
        Catch
        End Try

    End Sub

    Private Sub timer_to_stop_audio_capture_Tick()
        timer_to_stop_capture.Enabled = False
        EndAudioCapture()
    End Sub

    ''' <summary>
    ''' Starts capturing.
    ''' </summary>
    Public Sub StartCapture()

        If isCapturing = True Then Return

        isCapturing = True

        If capture__ = DeviceCapture.MotionPicture Or capture__ = DeviceCapture.ScreenMotionPicture Or capture__ = DeviceCapture.Video Or capture__ = DeviceCapture.ScreenVideo Then
            StartVideoOrScreenCapture()
        ElseIf capture__ = DeviceCapture.SingleImage Or capture__ = DeviceCapture.ScreenSingleImage Then
            StartImageCapture()
        ElseIf capture__ = DeviceCapture.Audio Then
            StartAudioCapture()
        End If

    End Sub

    Private Sub StartImageCapture()
        timer_.Enabled = True

    End Sub

    Private Sub StartVideoOrScreenCapture()
        If capture__ = DeviceCapture.Video Or capture__ = DeviceCapture.ScreenVideo Then

            record("open new Type waveaudio Alias recsound", "", 0, 0)
            record("record recsound", "", 0, 0)

        End If

        timer_.Enabled = True
        If timer_to_stop_capture IsNot Nothing Then
            timer_to_stop_capture.Enabled = True
        End If

    End Sub

    Private Sub StartAudioCapture()

        record("open new Type waveaudio Alias recsound", "", 0, 0)
        record("record recsound", "", 0, 0)

        If _number_of_seconds_to_capture > 0 Then
            timer_to_stop_capture = New Timer
            With timer_to_stop_capture
                .Interval = _number_of_seconds_to_capture * 1000
                AddHandler .Tick, New EventHandler(AddressOf timer_to_stop_audio_capture_Tick)
                .Enabled = True
            End With

        End If

    End Sub

#End Region

#Region "Callback"

    ''' <summary>
    ''' Sends captured file to remote server via ftp. Works with DeviceCapture.SingleImage and DeviceCapture.ScreenSingleImage.
    ''' </summary>
    ''' <param name="directory_"></param>
    ''' <param name="ftp_server"></param>
    ''' <param name="username"></param>
    ''' <param name="password"></param>
    ''' <param name="remoteFolderPathEscapedString"></param>
    ''' <param name="target_filename_without_extension"></param>
    Public Sub SendToWeb(directory_ As String, ftp_server As String, username As String, password As String, remoteFolderPathEscapedString As String, Optional target_filename_without_extension As String = "01")
        SendImageToWeb(directory_, ftp_server, username, password, remoteFolderPathEscapedString, target_filename_without_extension)

    End Sub

    Private Sub SendImageToWeb(directory_ As String, ftp_server As String, username As String, password As String, remoteFolderPathEscapedString As String, Optional target_filename_without_extension As String = "01")
        Dim possible_image As String
        If Exists(directory_) Then
            Dim possible_images As New List(Of String)
            possible_images = GetFiles(directory_, "*" & ext__, FileIO.SearchOption.SearchTopLevelOnly)
            If possible_images.Count > 0 Then
                possible_image = possible_images(0)
            End If
        End If
        If possible_image IsNot Nothing Then
            'w Dim host As New Uploader(ftp_server, username, password)
            Dim host As New ServerSide.Ftp(ftp_server, username, password)
            host.uploadBytes(target_filename_without_extension & ext__, PictureFromStream(possible_image), remoteFolderPathEscapedString)
        End If

    End Sub

    'Private Sub SendAudioToWeb(directory_ As String, ftp_server As String, username As String, password As String, remoteFolderPathEscapedString As String, Optional target_filename_without_extension As String = "01")
    '    Dim possible_audio As String
    '    If Exists(directory_) Then
    '        Dim possible_audios As New List(Of String)
    '        possible_audios = GetFiles(directory_, "*" & default_audio_extension, FileIO.SearchOption.SearchTopLevelOnly)
    '        If possible_audios.Count > 0 Then
    '            possible_audio = possible_audios(0)
    '        End If
    '    End If
    '    If possible_audio IsNot Nothing Then
    '        Dim serverside As New Uploader(ftp_server, username, password)
    '        serverside.uploadBytes(target_filename_without_extension & default_audio_extension, PictureFromStream(possible_audio), remoteFolderPathEscapedString)
    '    End If
    'End Sub

#End Region

#Region "Playback"

    Private Property playback_timer As Timer
    Private Property playback_directory As String
    Private Property playback_image_control As PictureBox
    Private Property playback_images_from_directory As List(Of String)
    Private Property playback_images_from_list As List(Of Image)
    Private Property playback_audio_file As String

    Public Sub StartPlayback(image_control As PictureBox, directory_ As String, Optional timer As Timer = Nothing, Optional timer_interval_in_ms As Integer = 20, Optional file_type As String = "*.jpg", Optional audio_file As String = Nothing)
        If capture__ = DeviceCapture.MotionPicture Or capture__ = DeviceCapture.ScreenMotionPicture Or capture__ = DeviceCapture.Video Or capture__ = DeviceCapture.ScreenVideo Then
            StartMotionPicturePlayback(image_control, directory_, timer, timer_interval_in_ms, file_type, audio_file)
        End If
    End Sub

    Public Sub StartPlayback(directory_ As String)
        If capture__ = DeviceCapture.Audio Then
            StartAudioPlayback(directory_)
        End If
    End Sub

    Public Sub ViewCaputuredImage(image_control As PictureBox, directory_ As String)
        If capture__ = DeviceCapture.SingleImage Or capture__ = DeviceCapture.ScreenSingleImage Then
            ViewImage(image_control, directory_)
        End If
    End Sub

    Public Sub StopPlayback()
        Try
            StopAudioPlayback()
        Catch ex As Exception

        End Try
        Try
            StopMotionPicturePlayback()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub StartAudioPlayback(directory_ As String)

        If Exists(directory_) Then
            Dim possible_audio_files As New List(Of String)
            Try
                possible_audio_files = GetFiles(directory_, "*" & default_audio_extension, FileIO.SearchOption.SearchTopLevelOnly)
            Catch ex As Exception

            End Try
            If possible_audio_files.Count > 0 Then
                playback_audio_file = possible_audio_files(0)
            Else
                Return
            End If
        Else
            Return
        End If


        PlayAudioFile(playback_audio_file, AudioPlayMode.Background)

    End Sub

    'Private Sub StartAudioPlayback(Optional audio_file As String = Nothing, Optional directory_ As String = Nothing)

    '    'If audio_file IsNot Nothing Then
    '    '    playback_audio_file = audio_file
    '    'ElseIf AudioCaptured IsNot Nothing Then
    '    '    If AudioCaptured.Length > 0 Then
    '    '        playback_audio_file = AudioCaptured
    '    '    End If
    '    'End If

    '    If audio_file IsNot Nothing Then
    '        playback_audio_file = audio_file

    '    ElseIf directory_ IsNot Nothing Then
    '        If Exists(directory_) Then
    '            Dim possible_audio_files As New List(Of String)
    '            Try
    '                possible_audio_files = GetFiles(directory_, "*" & default_audio_extension, FileIO.SearchOption.SearchTopLevelOnly)
    '            Catch ex As Exception

    '            End Try
    '            If possible_audio_files.Count > 0 Then
    '                playback_audio_file = possible_audio_files(0)
    '            End If
    '        End If
    '    Else
    '        Return
    '    End If


    '    PlayAudioFile(audio_file, AudioPlayMode.Background)

    'End Sub

    Private Sub StopAudioPlayback()
        If playback_audio_file Is Nothing Then Return

        Try
            StopPlayingAudioFile()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub StartMotionPicturePlayback(image_control As PictureBox, directory_ As String, Optional timer As Timer = Nothing, Optional timer_interval_in_ms As Integer = 20, Optional file_type As String = "*.jpg", Optional audio_file As String = Nothing)
        If directory_ Is Nothing Then Return

        Dim images_ As New List(Of String)
        Try
            images_ = GetFiles(directory_, file_type, FileIO.SearchOption.SearchTopLevelOnly)
        Catch ex As Exception
            Return
        End Try

        If images_.Count < 1 Then Return

        If timer Is Nothing Then
            playback_timer = New Timer
            playback_timer.Interval = 20
        Else
            playback_timer = timer
        End If

        playback_timer.Interval = timer_interval_in_ms

        playback_directory = directory_
        playback_image_control = image_control
        playback_images_from_directory = images_

        If audio_file IsNot Nothing Then
            playback_audio_file = audio_file
        Else
            Dim possible_audio_files As New List(Of String)
            Try
                possible_audio_files = GetFiles(directory_, "*" & default_audio_extension, FileIO.SearchOption.SearchTopLevelOnly)
            Catch ex As Exception

            End Try
            If possible_audio_files.Count > 0 Then
                playback_audio_file = possible_audio_files(0)
            End If
        End If

        With playback_timer
            AddHandler .Tick, New EventHandler(AddressOf PlaybackFromDirectory)
            If playback_audio_file IsNot Nothing Then
                Try
                    PlayAudioFile(playback_audio_file, AudioPlayMode.Background)
                Catch ex As Exception

                End Try
            End If
            .Enabled = True
        End With
        _counter = 0
    End Sub

    'Private Sub StartMotionPicturePlayback(image_control As PictureBox, Optional timer As Timer = Nothing, Optional timer_interval_in_ms As Integer = 20, Optional images As List(Of Image) = Nothing, Optional audio_file As String = Nothing)

    '    If images IsNot Nothing Then
    '        playback_images_from_list = images
    '    ElseIf _ImagesCaptured IsNot Nothing Then
    '        If _ImagesCaptured.Count > 0 Then
    '            playback_images_from_list = _ImagesCaptured
    '        Else
    '            Return
    '        End If
    '    End If

    '    If timer Is Nothing Then
    '        playback_timer = New Timer
    '        playback_timer.Interval = 20
    '    Else
    '        playback_timer = timer
    '    End If

    '    playback_timer.Interval = timer_interval_in_ms

    '    playback_image_control = image_control

    '    If audio_file IsNot Nothing Then
    '        playback_audio_file = audio_file
    '    ElseIf AudioCaptured IsNot Nothing Then
    '        If AudioCaptured.Length > 0 Then
    '            playback_audio_file = AudioCaptured
    '        End If
    '    End If

    '    With playback_timer
    '        AddHandler .Tick, New EventHandler(AddressOf PlaybackFromImagesList)
    '        If playback_audio_file IsNot Nothing Then
    '            Try
    '                PlayAudioFile(playback_audio_file, AudioPlayMode.Background)
    '            Catch ex As Exception

    '            End Try
    '        End If
    '        .Enabled = True
    '    End With
    '    _counter = 0
    'End Sub

    Private Sub StopMotionPicturePlayback()
        If playback_timer Is Nothing Then Return
        Try
            playback_timer.Enabled = False
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ViewImage(image_control As PictureBox, directory_ As String)
        If Exists(directory_) Then
            Dim possible_images As New List(Of String)
            possible_images = GetFiles(directory_, "*" & ext__, FileIO.SearchOption.SearchTopLevelOnly)
            If possible_images.Count > 0 Then
                image_control.BackgroundImage = Image.FromFile(possible_images(0))
            End If
        End If
    End Sub

    Dim _counter

    Private Sub PlaybackFromDirectory()
        If _counter >= playback_images_from_directory.Count - 1 Then
            playback_timer.Enabled = False
            Return
        End If
        Try
            playback_image_control.BackgroundImage = Image.FromFile(playback_images_from_directory(counter))
        Catch ex As Exception

        End Try
        If _counter < playback_images_from_directory.Count - 1 Then
            counter += 1
        End If
    End Sub

    'Private Sub PlaybackFromImagesList()
    '    If _counter >= playback_images_from_list.Count - 1 Then
    '        playback_timer.Enabled = False
    '        Return
    '    End If
    '    Try
    '        playback_image_control.BackgroundImage = playback_images_from_list(counter)
    '    Catch ex As Exception

    '    End Try
    '    If _counter < playback_images_from_list.Count - 1 Then
    '        counter += 1
    '    End If
    'End Sub

#End Region

#Region "SingleImage, Video"

    Private Sub cam_NewFrame(sender As Object, eventArgs As NewFrameEventArgs)

        Dim bit As Bitmap = eventArgs.Frame.Clone()
        returned.Add(bit)

        image_control_.Image = bit
    End Sub

    Private Sub ImageAndMotionPictureAndVideoTimer_Tick()
        If image_control_.Image IsNot Nothing Then

            'If directory_ IsNot Nothing Then

            If directory_.Length > 0 Then
                Try
                    MkDir(directory_)
                Catch ex As Exception

                End Try
            End If
            counter += 1
            If directory_.Length > 0 Then
                Try
                    image_control_.Image.Save(directory_ & "\" & LeadingZero(counter) & ext__)
                Catch ex As Exception

                End Try

            End If
            'End If

            If capture__ = DeviceCapture.SingleImage Then
                timer_.Enabled = False
                StopCapture()
            End If
        End If
    End Sub
#End Region

End Class

