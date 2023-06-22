Imports AForge.Video
Imports AForge.Video.DirectShow
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Drawing.Imaging

Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports iNovation.Code.Web
Imports iNovation.Code.Sequel
Imports iNovation.Code.Media

Public Class CameraAndMic

#Region "Fields"
    Private Shared Property camWidth As Short = 640
    Private Shared Property camHeight As Short = 360
    Private Shared ReadOnly Property camSize As Size = New Size(640, 360)

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


    Private Shared Property directory_ As String

    Private Shared lingerImage As Boolean = False

    Private _ImagesCaptured As List(Of Image)
    Private _AudioCapturedFile As String

    Private Declare Function record Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

#End Region

#Region "Properties"

    ''' <summary>
    ''' Returns the images captured (DeviceCapture.Video) 
    ''' </summary>
    ''' <returns></returns>

    Public ReadOnly Property ImagesCaptured As List(Of Image)
        Get
            Return returned
        End Get
    End Property

    ''' <summary>
    ''' Returns the images captured (DeviceCapture.SingleImage) 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ImageCaptured As Image
        Get
            If returned IsNot Nothing And returned.Count > 0 Then
                Return returned(0)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property AudioCaptured As String
        Get
            Return _AudioCapturedFile
        End Get
    End Property


#End Region

#Region "Initialization"

    ''' <summary>
    ''' Starts Capturing immediately.
    ''' </summary>
    ''' <param name="prefer">Capture single image, video, screen single image, screen video, or audio. 
    ''' If set to video, screen video or audio, then you must call EndCapture() manually except if number_of_seconds_to_capture is more than 0.
    ''' </param>
    ''' <param name="number_of_seconds_to_capture">Specify number of seconds to capture if prefer is set to Video.</param>
    ''' <param name="image_control__">To preview captured images.</param>
    ''' <param name="directory__">Where to save images in specified format (i.e. ext).</param>
    ''' <param name="ext">Extension to save image files in.</param>
    ''' <param name="linger_image__">Should image_control__ go blank after capturing.</param>
    ''' <example>
    ''' Private camera As SecurityAdapter
    ''' Private caputred_images As List(Of Image)
    ''' Private captured_image As Image
    ''' camera = New SecurityAdapter(DeviceCapture.Video, 5, PictureBox1, some_directory_path)
    ''' to get caputred_images (DeviceCapture.Video), get the property ImagesCaptured
    ''' to get caputred_image (DeviceCapture.SingleImage), get the property ImageCaptured
    ''' </example>

    Public Sub New(Optional prefer As DeviceCapture = DeviceCapture.SingleImage, Optional number_of_seconds_to_capture As Integer = 0, Optional directory__ As String = "", Optional ext As String = ".jpg", Optional image_control__ As Object = Nothing, Optional linger_image__ As Boolean = False)
        If prefer = DeviceCapture.ScreenVideo Or prefer = DeviceCapture.ScreenSingleImage Then
            InitializeForScreen(prefer, number_of_seconds_to_capture, image_control__, directory__, ext, linger_image__)
            Return
        End If

        If prefer = DeviceCapture.Audio Then
            capture__ = DeviceCapture.Audio
            InitializeForAudio(number_of_seconds_to_capture, directory__, ext)
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

        Dim timer__ As New Timer

        capture__ = prefer
        directory_ = directory__
        image_control_ = image_control
        timer_ = timer__
        lingerImage = linger_image__
        _number_of_seconds_to_capture = number_of_seconds_to_capture

        ext__ = ext

        If prefer = DeviceCapture.Video Or number_of_seconds_to_capture > 0 Then
            timer_to_stop_capture = New Timer
            With timer_to_stop_capture
                .Interval = number_of_seconds_to_capture * 1000
                AddHandler .Tick, New EventHandler(AddressOf timer_to_stop_video_or_screen_capture_Tick)
            End With

        End If

        InitializeForSingleImageAndVideo()

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

        Dim timer__ As New Timer

        capture__ = prefer
        directory_ = directory__
        image_control_ = image_control
        timer_ = timer__
        lingerImage = linger_image__
        _number_of_seconds_to_capture = number_of_seconds_to_capture

        ext__ = ext

        If prefer = DeviceCapture.ScreenVideo Or number_of_seconds_to_capture > 0 Then
            timer_to_stop_capture = New Timer
            With timer_to_stop_capture
                .Interval = number_of_seconds_to_capture * 1000
                AddHandler .Tick, New EventHandler(AddressOf timer_to_stop_video_or_screen_capture_Tick)
            End With

        End If

        With timer_
            .Interval = 20
            AddHandler .Tick, New EventHandler(AddressOf ScreenTimer_Tick)
        End With

        timer_.Enabled = True
        If timer_to_stop_capture IsNot Nothing Then
            timer_to_stop_capture.Enabled = True
        End If

    End Sub

    Private Sub InitializeForSingleImageAndVideo()
        webcam = New FilterInfoCollection(FilterCategory.VideoInputDevice)
        For Each VideoCaptureDevice As FilterInfo In webcam
            devices_.Add(VideoCaptureDevice.Name)
        Next
        If devices_.Count < 1 Then Return
        selected_device = devices_(0)
        With timer_
            .Interval = 20
            AddHandler .Tick, New EventHandler(AddressOf ImageAndVideoTimer_Tick)
        End With

        cam = New VideoCaptureDevice(webcam.Item(0).MonikerString)
        AddHandler cam.NewFrame, New NewFrameEventHandler(AddressOf cam_NewFrame)
        cam.Start()

        timer_.Enabled = True
        If timer_to_stop_capture IsNot Nothing Then
            timer_to_stop_capture.Enabled = True
        End If

    End Sub

    Private Sub InitializeForAudio(Optional number_of_seconds_to_capture As Integer = 0, Optional directory__ As String = "", Optional ext As String = ".jpg")

        directory_ = directory__
        _number_of_seconds_to_capture = number_of_seconds_to_capture

        ext__ = ext


        record("open new Type waveaudio Alias recsound", "", 0, 0)
        record("record recsound", "", 0, 0)

        If number_of_seconds_to_capture > 0 Then
            timer_to_stop_capture = New Timer
            With timer_to_stop_capture
                .Interval = number_of_seconds_to_capture * 1000
                AddHandler .Tick, New EventHandler(AddressOf timer_to_stop_audio_capture_Tick)
                .Enabled = True
            End With

        End If

    End Sub
#End Region

#Region "Audio"
    Private Sub EndAudioCapture()
        Dim filename As String = directory_ & "\" & String.Format("{0:00}", CStr(Now.Month)) & "_" & String.Format("{0:00}", CStr(Now.Day)) & "_" & String.Format("{0:0000}", CStr(Now.Year)) & "_T_" & String.Format("{0:00}", CStr(Now.Hour)) & "_" & String.Format("{0:00}", CStr(Now.Minute)) & "_" & String.Format("{0:00}", CStr(Now.Millisecond)) & ext__.Replace("*", "")
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
            EndCapture()
        End If
    End Sub

#End Region

#Region "Shared"
    Private Sub timer_to_stop_video_or_screen_capture_Tick()
        timer_to_stop_capture.Enabled = False
        EndCapture()
    End Sub

    ''' <summary>
    ''' Stops capturing. Must be called manually if prefer is Device.Video.
    ''' </summary>
    Public Sub EndCapture()

        If capture__ = DeviceCapture.Audio Then
            EndAudioCapture()
            Return
        End If

        Try
            timer_.Enabled = False
        Catch
        End Try

        Try
            If cam.IsRunning Then cam.Stop()
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

#End Region

#Region "ToDo"

#Region "Playback"

    Private Property playback_timer As Timer
    Private Property playback_directory As String
    Private Property playback_image_control As PictureBox
    Private Property images As List(Of Image)

    Private Sub StartPlayback(image_control As PictureBox, directory_ As String, Optional file_type As String = "*.jpg")
        If directory_ Is Nothing Then Return

        'Dim images_ As List(Of String) = GetFiles(directory_, )

        playback_directory = directory_
        playback_image_control = image_control
        playback_timer = New Timer



        With playback_timer
            .Interval = 20
            AddHandler .Tick, New EventHandler(AddressOf PlaybackFromDirectory)
            .Enabled = True
        End With
    End Sub

    Private Sub PlaybackFromDirectory()

    End Sub

#End Region

#End Region


#Region "SingleImage, Video"

    Private Sub cam_NewFrame(sender As Object, eventArgs As NewFrameEventArgs)
        Dim bit As Bitmap = eventArgs.Frame.Clone()
        returned.Add(bit)

        image_control_.Image = bit
    End Sub

    Private Sub ImageAndVideoTimer_Tick()
        If image_control_.Image IsNot Nothing Then

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

            If capture__ = DeviceCapture.SingleImage Then
                timer_.Enabled = False
                EndCapture()
            End If
        End If
    End Sub
#End Region



End Class

