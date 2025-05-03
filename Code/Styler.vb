Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D

''' <summary>
''' This class contains methods for desktop development, particularly, styling the Form.
''' Reference to System.Windows.Forms may be required.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public Class Styler

#Region "init"

    Private Shared Property Instance As Styler
    Private Property borderColor As Color
    Private Property dialogWidth As Integer
    Private Property dialogHeight As Integer
    Private Sub New(borderColor As Color, dialogWidth As Integer, dialogHeight As Integer)
        Me.borderColor = borderColor
        Me.dialogHeight = dialogHeight
        Me.dialogWidth = dialogWidth
    End Sub

#End Region

#Region "OS API related"
    Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
    Private Declare Sub ReleaseCapture Lib "User32" ()
    Const WM_NCLBUTTONDOWN As Short = &HA1S
    Const HTCAPTION As Short = 2

    <DllImport("Gdi32.dll", EntryPoint:="CreateRoundRectRgn")>
    Private Shared Function CreateRoundRectRgn(ByVal nLeftRect As Integer, ByVal nTopRect As Integer, ByVal nRightRect As Integer, ByVal nBottomRect As Integer, ByVal nWidthEllipse As Integer, ByVal nHeightEllipse As Integer) As IntPtr
    End Function

    Private Shared Sub DialogMouseMove(sender As Object, e As MouseEventArgs)
        Dim Button As Short = e.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = (e.X)
        Dim Y As Single = (e.Y)
        Dim lngReturnValue As Integer
        If Button = 1 Then
            Call ReleaseCapture()
            lngReturnValue = SendMessage(sender.Handle.ToInt32, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If

    End Sub

#End Region

#Region "Font Variables"
    Public Font_Sans_12 As New Font("Microsoft Sans Serif", 12, GraphicsUnit.Point)
    Public Font_Sans_10 As New Font("Microsoft Sans Serif", 10, GraphicsUnit.Point)
    Public Font_Sans_14 As New Font("Microsoft Sans Serif", 14, GraphicsUnit.Point)
    Public Font_Bauhaus_12 As New Font("Bauhaus 93", 12, FontStyle.Bold, GraphicsUnit.Point)
    Public Font_Bauhaus_10 As New Font("Bauhaus 93", 10, FontStyle.Bold, GraphicsUnit.Point)
    Public Font_Bauhaus_14 As New Font("Bauhaus 93", 14, FontStyle.Bold, GraphicsUnit.Point)
    Public Shared Font_Verdana_12 As New Font("Verdana", 12, GraphicsUnit.Point)
    Public Shared Font_Verdana_10 As Font = New Font("Verdana", 10, GraphicsUnit.Point)
    Public Font_Verdana_14 As Font = New Font("Verdana", 14, GraphicsUnit.Point)
    Public Shared Font_JetBrains_12 As New Font("JetBrains Mono", 12, GraphicsUnit.Point)
    Public Font_JetBrains_10 As Font = New Font("JetBrains Mono", 10, GraphicsUnit.Point)
    Public Font_JetBrains_14 As Font = New Font("JetBrains Mono", 14, GraphicsUnit.Point)
    Public Font_Cascadia_Mono_Light_12 As New Font("Cascadia Mono Light", 12, GraphicsUnit.Point)
    Public Font_Cascadia_Mono_Light_10 As Font = New Font("Cascadia Mono Light", 10, GraphicsUnit.Point)
    Public Font_Cascadia_Mono_Light_14 As Font = New Font("Cascadia Mono Light", 14, GraphicsUnit.Point)
    Public Font_DejaVu_Sans_Light_12 As New Font("DejaVu Sans Light", 12, GraphicsUnit.Point)
    Public Font_DejaVu_Sans_Light_10 As Font = New Font("DejaVu Sans Light", 10, GraphicsUnit.Point)
    Public Font_DejaVu_Sans_Light_14 As Font = New Font("DejaVu Sans Light", 14, GraphicsUnit.Point)
    Public Font_DejaVu_Sans_12 As New Font("DejaVu Sans", 12, GraphicsUnit.Point)
    Public Font_DejaVu_Sans_10 As Font = New Font("DejaVu Sans", 10, GraphicsUnit.Point)
    Public Font_DejaVu_Sans_14 As Font = New Font("DejaVu Sans", 14, GraphicsUnit.Point)

#End Region

#Region "Theme Variables"

    'Yellow
    'Public ReadOnly Property yellow_border_color As Color = Color.Black
    Public Shared ReadOnly Property yellow_background_color As Color = Color.Wheat
    Public Shared ReadOnly Property yellow_foreground_color As Color = Color.Brown
    Public Shared ReadOnly Property yellow_alt_foreground_color As Color = Color.Brown

    'Green
    'Public ReadOnly Property net_border_background_color_green As Color = Color.Black
    Public Shared ReadOnly Property green_background_color As Color = Color.DarkGreen
    Public Shared ReadOnly Property green_foreground_color As Color = Color.Lime
    Public Shared ReadOnly Property green_alt_foreground_color As Color = Color.Turquoise

    'Turqoise
    'Public ReadOnly Property net_border_background_color_turqoise As Color = Color.Black
    Public Shared ReadOnly Property turqoise_background_color As Color = Color.Turquoise
    Public Shared ReadOnly Property turqoise_foreground_color As Color = Color.FromArgb(0, 0, 102)
    Public Shared ReadOnly Property turqoise_alt_foreground_color As Color = Color.FromArgb(0, 0, 102)

    'Purple
    'Public ReadOnly Property net_border_background_color_purple As Color = Color.Black
    Public Shared ReadOnly Property purple_background_color As Color = Color.FromArgb(0, 0, 102)
    Public Shared ReadOnly Property purple_foreground_color As Color = Color.Magenta
    Public Shared ReadOnly Property purple_alt_foreground_color As Color = Color.DarkMagenta

    'White
    'Public ReadOnly Property net_border_background_color_white As Color = Color.FromArgb(128, 128, 255)
    Public Shared ReadOnly Property white_background_color As Color = Color.FromArgb(255, 255, 255)
    Public Shared ReadOnly Property white_foreground_color As Color = Color.Navy
    Public Shared ReadOnly Property white_alt_foreground_color As Color = Color.Navy

    'Velvet
    'Public ReadOnly Property net_border_background_color_velvet As Color = Color.FromArgb(177, 67, 67)
    Public Shared ReadOnly Property velvet_background_color As Color = Color.FromArgb(34, 34, 34)
    Public Shared ReadOnly Property velvet_foreground_color As Color = Color.FromArgb(72, 61, 139)
    Public Shared ReadOnly Property velvet_alt_foreground_color As Color = Color.FromArgb(177, 67, 67)

    'Brown
    'Public ReadOnly Property net_border_background_color_brown As Color = Color.Black
    Public Shared ReadOnly Property brown_background_color As Color = Color.FromArgb(34, 34, 34)
    Public Shared ReadOnly Property brown_foreground_color As Color = Color.Wheat
    Public Shared ReadOnly Property brown_alt_foreground_color As Color = Color.Wheat
#End Region

#Region "GetColors"
    Private Shared Function BackgroundColorFromBaseTheme(current_theme As BaseTheme) As Color
        Select Case current_theme
            Case BaseTheme.Green
                BackgroundColorFromBaseTheme = green_background_color
            Case BaseTheme.Turqoise
                BackgroundColorFromBaseTheme = turqoise_background_color
            Case BaseTheme.Velvet
                BackgroundColorFromBaseTheme = velvet_background_color
            Case BaseTheme.Purple
                BackgroundColorFromBaseTheme = purple_background_color
            Case BaseTheme.White
                BackgroundColorFromBaseTheme = white_background_color
            Case BaseTheme.Brown
                BackgroundColorFromBaseTheme = brown_background_color
            Case BaseTheme.Yellow
                BackgroundColorFromBaseTheme = yellow_background_color
        End Select
    End Function

    Private Shared Function AltForeGroundColorFromBaseTheme(current_theme As BaseTheme) As Color
        Select Case current_theme
            Case BaseTheme.Green
                AltForeGroundColorFromBaseTheme = green_alt_foreground_color
            Case BaseTheme.Turqoise
                AltForeGroundColorFromBaseTheme = turqoise_alt_foreground_color
            Case BaseTheme.Velvet
                AltForeGroundColorFromBaseTheme = velvet_alt_foreground_color
            Case BaseTheme.Purple
                AltForeGroundColorFromBaseTheme = purple_alt_foreground_color
            Case BaseTheme.White
                AltForeGroundColorFromBaseTheme = white_alt_foreground_color
            Case BaseTheme.Brown
                AltForeGroundColorFromBaseTheme = brown_alt_foreground_color
            Case BaseTheme.Yellow
                AltForeGroundColorFromBaseTheme = yellow_alt_foreground_color
        End Select
    End Function

    Private Shared Function ForeGroundColorFromBaseTheme(current_theme As BaseTheme) As Color
        Select Case current_theme
            Case BaseTheme.Green
                ForeGroundColorFromBaseTheme = green_foreground_color
            Case BaseTheme.Turqoise
                ForeGroundColorFromBaseTheme = turqoise_foreground_color
            Case BaseTheme.Velvet
                ForeGroundColorFromBaseTheme = velvet_foreground_color
            Case BaseTheme.Purple
                ForeGroundColorFromBaseTheme = purple_foreground_color
            Case BaseTheme.White
                ForeGroundColorFromBaseTheme = white_foreground_color
            Case BaseTheme.Brown
                ForeGroundColorFromBaseTheme = brown_foreground_color
            Case BaseTheme.Yellow
                ForeGroundColorFromBaseTheme = yellow_foreground_color
        End Select
    End Function
    'Private Function net_border_background_color(current_theme As BaseTheme) As Color
    '    Select Case current_theme
    '        Case BaseTheme.Green
    '            net_border_background_color = net_border_background_color_green
    '        Case BaseTheme.Turqoise
    '            net_border_background_color = net_border_background_color_turqoise
    '        Case BaseTheme.Velvet
    '            net_border_background_color = net_border_background_color_velvet
    '        Case BaseTheme.Purple
    '            net_border_background_color = net_border_background_color_purple
    '        Case BaseTheme.White
    '            net_border_background_color = net_border_background_color_white
    '        Case BaseTheme.Brown
    '            net_border_background_color = net_border_background_color_brown
    '        Case BaseTheme.Yellow
    '            net_border_background_color = yellow_border_color
    '    End Select
    'End Function

#End Region

#Region "Fields"

    Private Shared time_label__ As Label
    'Private app_name__ As String

    Private Shared dialog__ As Form
    Private Shared LeftBorder__ As PictureBox
    Private Shared RightBorder__ As PictureBox
    Private Shared TopBorder__ As PictureBox
    Private Shared BottomBorder__ As PictureBox
    Private Shared Title__ As Label
    Private AcceptButton__ As Button
    Private Shared MinimizeButton__ As Label
    Private Shared AdjustWindowSizeAutomatically As Boolean
    Public Shared Property DialogCloseButton As Label
    'Private MenuStrip__ As MenuStrip
    Private Shared SystemCloseButton__ As Button
    Private Shared ShowMimize__ As Boolean
    'Private UseMaximize__ As Boolean
    'Private UseMenustrip__ As Boolean
    'Private UseClose__ As Boolean
    Private Shared TimeLabel__ As Label
    Private Shared set_format_params As FormatParameters
    Private Shared set_dialog_params As DialogParameters
    Private Shared set_base_theme As BaseTheme

    'TitleBar_ = TitleBar : FooterBar_ = FooterBar
    'Private MaximizeButton__ As Label
    Private Shared ShowTime__ As Boolean
    Private Shared DialogTitle__ As Label
    Public Shared Property FadeOutTimer As Timer
    Private Shared TimeTimer__ As New Timer


    Private Shared d As Form
#End Region

#Region "Enums and Structures"

    Public Enum WhatCloseButtonDoes
        ExitsProgram = 0
        ClosesDialog = 1
        MinimizesDialog = 2
        DoesNothing = 3
    End Enum

    Public Enum BaseTheme
        Brown = 0
        Velvet = 1
        Yellow = 2
        Green = 3
        Turqoise = 4
        Purple = 5
        White = 6
    End Enum

    Public Structure FormatParameters
        Public title_font As Font
        Public title_color As Color
        Public background_color As Color
        Public dialog_font As Font

        Public border_color As Color

    End Structure

    Public Structure DialogParameters
        Public sort_dropdowns_and_listboxes As Boolean
        Public show_minimize_button As Boolean
        Public use_default_cancel_button As Boolean
        'Public show_time As Boolean
        Public show_title As Boolean
        Public title As String
        Public close_button_should As WhatCloseButtonDoes
        Public read_this_loud_on_exit As String
    End Structure
    Public Structure ButtonInfo
        Public theLabel As Label
        Public showIt As Boolean
        Public inThisColor As Color
        Public withThisCallback As EventHandler
    End Structure

    Public Structure DialogInfo
        Public titleLabel As Label
        Public title As String
        Public titleColor As Color
        Public dialog As Form
        Public backgroundImage As Image
        Public backgroundImageLayout As ImageLayout
        Public backgroundColor As Color
        Public borderColor As Color
    End Structure
#End Region

#Region "Support"

    Private Shared Sub Dialog_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        If Instance Is Nothing Then Return
        Using pen As New Pen(Instance.borderColor, 2)
            e.Graphics.DrawPath(pen, GetRoundedRectPath(New Rectangle(0, 0, Instance.dialogWidth - 1, Instance.dialogHeight - 1), 20))
        End Using
    End Sub
    Private Shared Function GetRoundedRectPath(ByVal rect As Rectangle, ByVal radius As Integer) As GraphicsPath
        Dim path As New GraphicsPath()
        path.AddArc(rect.X, rect.Y, radius, radius, 180, 90) ' Top-left corner
        path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 90) ' Top-right corner
        path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90) ' Bottom-right corner
        path.AddArc(rect.X, rect.Bottom - radius, radius, radius, 90, 90) ' Bottom-left corner
        path.CloseFigure() ' Close the path
        Return path
    End Function

#End Region

#Region "Exported"
    Public Shared Sub Style(dialog As Form, LabelHasThisForeColor As Color, Optional ListBoxHasBothScrollBars As Boolean = True)
        AddHandler dialog.MouseMove, New MouseEventHandler(AddressOf MouseMove)

        For Each l As Control In dialog.Controls
            If TypeOf (l) Is Label Then
                CType(l, Label).BackColor = Color.Transparent
                CType(l, Label).ForeColor = If(LabelHasThisForeColor <> Nothing, LabelHasThisForeColor, Color.Black)
            End If
            If TypeOf (l) Is PictureBox Then
                CType(l, PictureBox).BackColor = Color.Transparent
            End If
            If TypeOf (l) Is ListBox Then
                CType(l, ListBox).ScrollAlwaysVisible = True
                CType(l, ListBox).HorizontalScrollbar = ListBoxHasBothScrollBars
            End If
        Next

        'If LabelHasThisForeColor <> Nothing Then
        '    For Each l As Control In dialog.Controls
        '        If TypeOf (l) Is Label Then
        '            CType(l, Label).ForeColor = LabelHasThisForeColor
        '        End If
        '    Next
        'End If

        'If ListBoxHasBothScrollBars Then
        '    For Each l As Control In dialog.Controls
        '        If TypeOf (l) Is ListBox Then
        '            CType(l, ListBox).ScrollAlwaysVisible = True
        '            CType(l, ListBox).HorizontalScrollbar = True
        '        End If
        '    Next
        'End If



    End Sub
    Public Shared Sub Style(dialog As Form)
        AddHandler dialog.MouseMove, New MouseEventHandler(AddressOf MouseMove)


        'labels
        For Each l As Control In dialog.Controls
            If TypeOf (l) Is Label Then
                CType(l, Label).BackColor = Color.Transparent
            End If
        Next

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dialog">The current System.Windows.Forms.Form</param>
    ''' <param name="adjust_window_size_automatically">Adjust window size to accomodate for new placement of controls. When this is set to false, top of the topmost control should be set to 63 or more.</param>
    Public Shared Sub Style(dialog As Form, adjust_window_size_automatically As Boolean)
        Dim set_dialog_params As New DialogParameters With {
            .sort_dropdowns_and_listboxes = True,
            .show_minimize_button = False,
            .show_title = True,
            .use_default_cancel_button = False,
            .close_button_should = WhatCloseButtonDoes.ClosesDialog,
            .title = "",
            .read_this_loud_on_exit = ""
            }

        Dim set_format_params As New FormatParameters With {
            .background_color = BackgroundColorFromBaseTheme(set_base_theme),
            .border_color = Color.Black,
            .dialog_font = Font_JetBrains_12,
            .title_color = Color.Black,
            .title_font = Font_Verdana_12
            }
        Style(dialog, set_dialog_params, set_format_params, BaseTheme.Brown, adjust_window_size_automatically)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dialog">The current System.Windows.Forms.Form</param>
    ''' <param name="dialog_params">Set this to DoesNothing if you want to use your own routine for closebutton by calling AddHandler(current_instance_here.DialogCloseButton.Click ...</param>
    ''' <param name="format_params"></param>
    ''' <param name="base_theme"></param>
    ''' <param name="adjust_window_size_automatically">Adjust window size to accomodate for new placement of controls. When this is set to false, top of the topmost control should be set to 63 or more.</param>
    Public Shared Sub Style(dialog As Form, dialog_params As DialogParameters, format_params As FormatParameters, base_theme As BaseTheme, adjust_window_size_automatically As Boolean)
        set_base_theme = If(base_theme <> Nothing, base_theme, BaseTheme.Brown)

        set_dialog_params = New DialogParameters With {
            .sort_dropdowns_and_listboxes = If(dialog_params.sort_dropdowns_and_listboxes <> Nothing, dialog_params.sort_dropdowns_and_listboxes, True),
            .show_minimize_button = If(dialog_params.show_minimize_button <> Nothing, dialog_params.show_minimize_button, False),
            .show_title = If(dialog_params.show_title <> Nothing, dialog_params.show_title, True),
            .use_default_cancel_button = If(dialog_params.use_default_cancel_button <> Nothing, dialog_params.use_default_cancel_button, False),
            .close_button_should = If(dialog_params.close_button_should <> Nothing, dialog_params.close_button_should, WhatCloseButtonDoes.ClosesDialog),
            .title = If(dialog_params.title IsNot Nothing, dialog_params.title, ""),
            .read_this_loud_on_exit = If(dialog_params.read_this_loud_on_exit IsNot Nothing, dialog_params.read_this_loud_on_exit, "")
            }

        set_format_params = New FormatParameters With {
            .background_color = If(format_params.background_color <> Nothing, format_params.background_color, BackgroundColorFromBaseTheme(set_base_theme)),
            .border_color = If(format_params.border_color <> Nothing, format_params.border_color, Color.Black),
            .dialog_font = If(format_params.dialog_font IsNot Nothing, format_params.dialog_font, Font_JetBrains_12),
            .title_color = If(format_params.title_color <> Nothing, format_params.title_color, Color.Black),
            .title_font = If(format_params.title_font IsNot Nothing, format_params.title_font, Font_Verdana_12)
            }


        dialog__ = dialog
        d = dialog

        Dim LeftBorder As New PictureBox
        Dim RightBorder As New PictureBox
        Dim TopBorder As New PictureBox
        Dim BottomBorder As New PictureBox
        Dim DialogTitle As New Label

        FadeOutTimer = New Timer


        Dim CloseButton As New Label
        Dim MinimizeButton As New Label
        'Dim MaximizeButton As New Label
        'Dim AcceptButton As New Button
        Dim CancelButton As New Button ''SystemCloseButton


        Dim TimeLabel As New Label
        Dim EmptyCloseButton As New Button
        'Dim TitleBar As New PictureBox, FooterBar As New PictureBox
        Dim TimeTimer As New Timer

        d.Controls.Add(CloseButton)
        d.Controls.Add(MinimizeButton)
        'd.Controls.Add(MaximizeButton)
        d.Controls.Add(LeftBorder)
        d.Controls.Add(RightBorder)
        d.Controls.Add(TopBorder)
        d.Controls.Add(BottomBorder)
        d.Controls.Add(DialogTitle)
        'd.Controls.Add(AcceptButton)
        d.Controls.Add(CancelButton)
        d.Controls.Add(TimeLabel)


        CloseButton.Text = ChrW(10539)
        MinimizeButton.Text = ChrW(800)
        'MaximizeButton.Text = ChrW(10064)

        MinimizeButton.Visible = set_dialog_params.show_minimize_button

        time_label__ = TimeLabel
        'app_name__ = set_dialog_params.title

        DialogTitle.Text = If(set_dialog_params.show_title, set_dialog_params.title, "")


        DialogCloseButton = CloseButton
        MinimizeButton__ = MinimizeButton
        LeftBorder__ = LeftBorder
        RightBorder__ = RightBorder
        TopBorder__ = TopBorder
        BottomBorder__ = BottomBorder
        DialogTitle__ = DialogTitle
        'AcceptButton__ = AcceptButton
        SystemCloseButton__ = CancelButton
        'UseMaximize__ = UseMaximize


        'MenuStrip__ = MenuStrip
        ShowMimize__ = set_dialog_params.show_minimize_button
        'UseMenustrip__ = UseMenustrip
        'UseClose__ = UseClose
        'TitleBar_ = TitleBar : FooterBar_ = FooterBar

        'ShowTime__ = set_dialog_params.show_time
        'TimeTimer__ = TimeTimer
        'With TimeTimer__
        '    'If ShowTime__ Then
        '    If set_dialog_params.show_time Then
        '        .Interval = 1000
        '        AddHandler .Tick, New EventHandler(AddressOf ShowTimeNow)
        '        .Enabled = True
        '    End If
        'End With

        dialog.FormBorderStyle = FormBorderStyle.None
        dialog.BackColor = set_format_params.background_color

        If set_dialog_params.close_button_should <> WhatCloseButtonDoes.DoesNothing Then
            AddHandler CloseButton.Click, New EventHandler(AddressOf CloseDialog)
        End If

        If set_dialog_params.use_default_cancel_button Then
            dialog.CancelButton = CancelButton
            AddHandler CancelButton.Click, New EventHandler(AddressOf CloseDialog)
            'ElseIf UseClose = False Then
            '    Dialog.CancelButton = EmptyCloseButton
        End If
        CancelButton.Visible = False

        ''removed        AddHandler MaximizeButton.Click, New EventHandler(AddressOf Restore)
        '		If TitleBar.Tag = "" Then AddHandler TitleBar.DoubleClick, New EventHandler(AddressOf Restore)

        AddHandler dialog.MouseMove, New MouseEventHandler(AddressOf MouseMove)
        AddHandler MinimizeButton.Click, New EventHandler(AddressOf MinimizeDialog)

        'Dim control_t As TextBox
        'Dim control_t_counter As Integer = 0
        'For Each control_c As Control In dialog.Controls
        '    If TypeOf control_c Is TextBox Then
        '        control_t = control_c
        '        If control_t.Multiline = True And control_t.ReadOnly = False Then
        '            control_t_counter += 1
        '        End If
        'If control_t.Name.StartsWith("stat", StringComparison.CurrentCultureIgnoreCase) Then control_t.TabStop = False : control_t.ReadOnly = True : control_t.Enabled = True ': control_t.Font = stat_f
        'End If
        'Next
        'If control_t_counter < 1 And AcceptButton IsNot CancelButton And AcceptButton IsNot EmptyCloseButton Then d.AcceptButton = AcceptButton

        FormatButtons(d, set_base_theme)
        FormatTextBoxes(d, set_base_theme)
        FormatDataGridViews(d, set_base_theme)
        FormatComboBoxes(d, set_base_theme, set_dialog_params.sort_dropdowns_and_listboxes)
        FormatCheckBoxes(d, set_base_theme)
        FormatLabels(d, set_base_theme)
        FormatMenuStrips(d, set_base_theme)
        FormatPictureBoxes(d, set_base_theme)
        FormatNumericUpDowns(d, set_base_theme)
        FormatRadios(d, set_base_theme)
        FormatListBoxes(d, set_base_theme, set_dialog_params.sort_dropdowns_and_listboxes)

        ShieldControls(d)

        AdjustWindowSizeAutomatically = adjust_window_size_automatically


        SetStyle()

    End Sub

    ''' <summary>
    ''' If titleLabel, minimizeButton or maximizeButton is present, then the top of the topmost control on the form should be at least 54.
    ''' </summary>
    Public Shared Sub Style(dialogProperties As DialogInfo, minimizeButton As ButtonInfo, closeButton As ButtonInfo)
        Dim dialog As Form = dialogProperties.dialog
        If dialog Is Nothing Then Return
        If Not dialogProperties.Equals(Nothing) Then
            If Not dialogProperties.backgroundImage Is Nothing Then
                dialog.BackgroundImage = dialogProperties.backgroundImage
                dialog.BackgroundImageLayout = dialogProperties.backgroundImageLayout
            End If
            dialog.BackColor = dialogProperties.backgroundColor

            If Not dialogProperties.titleLabel Is Nothing And Not String.IsNullOrEmpty(dialogProperties.title) Then
                dialogProperties.titleLabel.Left = 16
                dialogProperties.titleLabel.Top = 20
                dialogProperties.titleLabel.BackColor = Color.Transparent
                dialogProperties.titleLabel.Text = dialogProperties.title
                dialogProperties.titleLabel.ForeColor = dialogProperties.titleColor
            End If
        End If

        dialog.FormBorderStyle = FormBorderStyle.None
        ' Create a rounded rectangle region
        Dim roundedRegion As IntPtr = CreateRoundRectRgn(0, 0, dialog.Width, dialog.Height, 20, 20)
        ' Set the Region property of the dialog
        dialog.Region = System.Drawing.Region.FromHrgn(roundedRegion)
        dialog.AllowTransparency = True
        dialog.Invalidate()

        AddHandler dialog.MouseMove, New MouseEventHandler(AddressOf DialogMouseMove)
        Instance = New Styler(dialogProperties.borderColor, dialog.Width, dialog.Height)
        AddHandler dialog.Paint, New PaintEventHandler(AddressOf Dialog_Paint)

        If Not minimizeButton.Equals(Nothing) Then
            If minimizeButton.showIt Then
                Dim MinimizeDialoglabel As Label = minimizeButton.theLabel
                MinimizeDialoglabel.Text = ChrW(800)
                MinimizeDialoglabel.Font = New Font("Microsoft Sans Serif", 14, FontStyle.Bold)
                MinimizeDialoglabel.Left = dialog.Width - 68
                MinimizeDialoglabel.Top = 8
                MinimizeDialoglabel.ForeColor = minimizeButton.inThisColor
                MinimizeDialoglabel.Visible = minimizeButton.showIt
                MinimizeDialoglabel.BringToFront()
                MinimizeDialoglabel.BackColor = Color.Transparent
                AddHandler MinimizeDialoglabel.Click, minimizeButton.withThisCallback
            End If
        End If

        If Not closeButton.Equals(Nothing) Then
            If closeButton.showIt Then
                Dim CloseDialoglabel As Label = closeButton.theLabel
                CloseDialoglabel.Text = ChrW(10539)
                CloseDialoglabel.Font = New Font("Microsoft Sans Serif", 14, FontStyle.Bold)
                CloseDialoglabel.Left = dialog.Width - 38
                CloseDialoglabel.Top = 15
                CloseDialoglabel.ForeColor = closeButton.inThisColor
                CloseDialoglabel.Visible = closeButton.showIt
                CloseDialoglabel.BringToFront()
                CloseDialoglabel.BackColor = Color.Transparent
                AddHandler CloseDialoglabel.Click, closeButton.withThisCallback
            End If
        End If

    End Sub
#End Region

#Region "Methods"
    Private Shared Sub SetStyle()
        If AdjustWindowSizeAutomatically Then
            dialog__.Height += If(ShowTime__, 63 + 54, 54)
        End If

        Dim default_margin = 18
        Dim _bottom = 17
        Dim default_width = 16

        'LeftBorder
        With LeftBorder__
            .Left = 0
            .Top = 0
            .Width = 1
            .Height = d.Height
            .BackColor = set_format_params.border_color
        End With

        'RightBorder
        With RightBorder__
            .Left = d.Width - 1
            .Top = 0
            .Width = 1
            .Height = d.Height
            .BackColor = set_format_params.border_color
        End With

        'TopBorder
        With TopBorder__
            .Left = 0
            .Top = 0
            .Width = d.Width
            .Height = 1
            .BackColor = set_format_params.border_color
        End With

        'BottomBorder
        With BottomBorder__
            .Left = 0
            .Top = d.Height - 1
            .Width = d.Width
            .Height = 1
            .BackColor = set_format_params.border_color
        End With

        'DialogTitle
        With DialogTitle__
            .Left = default_margin
            .Top = 14
            .Font = set_format_params.title_font
            .ForeColor = set_format_params.title_color
        End With

        'TimeLabel
        If ShowTime__ Then
            With time_label__
                .Left = default_margin
                .Top = d.Height - _bottom - .Height
                .Visible = True
            End With
        Else
            time_label__.Visible = False
        End If

        For Each l As Control In d.Controls
            If TypeOf (l) Is Label Then
                If l Is DialogCloseButton Then
                    With l
                        '.Font = item_f_m
                        .Font = Font_Verdana_10
                        .Width = default_width
                        .Left = dialog__.Width - default_margin - DialogCloseButton.Width
                        .Top = 14
                        .Cursor = Cursors.Hand
                        If set_base_theme = BaseTheme.Velvet Then
                            .ForeColor = AltForeGroundColorFromBaseTheme(set_base_theme)
                        ElseIf set_base_theme = BaseTheme.Brown Then
                            .ForeColor = Color.Green
                        ElseIf set_base_theme = Nothing Then
                            .ForeColor = Color.Green
                        Else
                            .ForeColor = set_format_params.border_color
                        End If

                    End With
                End If
                If l Is MinimizeButton__ Then
                    If ShowMimize__ Then
                        With l
                            .Font = Font_Verdana_10
                            .Width = default_width
                            .Left = dialog__.Width - default_margin - DialogCloseButton.Width - default_margin - MinimizeButton__.Width
                            .Top = 10
                            .Cursor = Cursors.Hand
                            If set_base_theme = BaseTheme.Velvet Then
                                .ForeColor = AltForeGroundColorFromBaseTheme(set_base_theme)
                            ElseIf set_base_theme = BaseTheme.Brown Then
                                .ForeColor = Color.Green
                            ElseIf set_base_theme = Nothing Then
                                .ForeColor = Color.Green
                            Else
                                .ForeColor = set_format_params.border_color
                            End If
                        End With
                    End If
                End If
            End If
        Next

        If AdjustWindowSizeAutomatically Then
            Try
                For Each c As Control In d.Controls
                    If c IsNot LeftBorder__ And
                            c IsNot RightBorder__ And
                            c IsNot TopBorder__ And
                            c IsNot BottomBorder__ And
                            c IsNot DialogTitle__ And
                            c IsNot time_label__ And
                            c IsNot DialogCloseButton And
                            c IsNot MinimizeButton__ And
                            c IsNot time_label__ And
                            TypeOf c IsNot Panel Then
                        c.Top += 63
                    End If
                Next
            Catch ex As Exception

            End Try
        End If

    End Sub

    'Private Shared Sub ShowTimeNow()
    '    time_label__.Text = "It is now " & Now.ToLongTimeString & ",  " & Now.ToShortDateString
    'End Sub


    Private Shared Sub CloseDialog()
        If set_dialog_params.read_this_loud_on_exit.Length > 0 Then
            Dim f As New iNovation.Code.Feedback
            f.say(set_dialog_params.read_this_loud_on_exit, False)
        End If

        With FadeOutTimer
            .Interval = 100
            AddHandler .Tick, New EventHandler(AddressOf DialogFadeOut)
            .Enabled = True
        End With
    End Sub

    Private Shared Sub DialogFadeOut()
        If dialog__.Opacity <= 0 Then
            FadeOutTimer.Enabled = False
            Select Case set_dialog_params.close_button_should
                Case WhatCloseButtonDoes.ClosesDialog
                    d.Close()
                Case WhatCloseButtonDoes.MinimizesDialog
                    d.WindowState = FormWindowState.Minimized
                'Case WhatCloseButtonShouldDo.DoNothing
                Case WhatCloseButtonDoes.ExitsProgram
                    Environment.Exit(0)
            End Select

            Exit Sub
        End If
        dialog__.Opacity -= 0.2
    End Sub

    Private Shared Sub MouseMove(sender As Object, e As MouseEventArgs)
        Dim Button As Short = e.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = (e.X)
        Dim Y As Single = (e.Y)
        Dim lngReturnValue As Integer
        If Button = 1 Then
            Call ReleaseCapture()
            lngReturnValue = SendMessage(sender.Handle.ToInt32, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If

    End Sub

    Private Shared Sub MinimizeDialog()
        d.WindowState = FormWindowState.Minimized
    End Sub


#End Region

#Region ""
    Private Shared Sub FormatButtons(d As Form, theme As BaseTheme)
        For Each b As Control In d.Controls
            If TypeOf (b) Is Button Then
                FormatButton(b, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatButton(b As Button, theme As BaseTheme)
        With b
            .FlatStyle = FlatStyle.Flat
            .BackColor = BackgroundColorFromBaseTheme(theme)
            If theme = BaseTheme.Brown Then
                .FlatAppearance.BorderColor = Color.Green
            End If
            .ForeColor = ForeGroundColorFromBaseTheme(theme)
            'If InStr(b.Tag, "trans") Or InStr(b.Tag, "transparent") Then b.FlatAppearance.BorderSize = 0
        End With
    End Sub

    Private Shared Sub FormatTextBoxes(d As Form, theme As BaseTheme)
        For Each t As Control In d.Controls
            If TypeOf (t) Is TextBox Then
                FormatTextBox(t, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatTextBox(t As TextBox, theme As BaseTheme)
        With t
            .BackColor = BackgroundColorFromBaseTheme(theme)
            .ForeColor = ForeGroundColorFromBaseTheme(theme)
            .AllowDrop = True
            Dim t_ As TextBox = t
            With t_
                If .Multiline = True Then
                    .ScrollBars = ScrollBars.Both
                    .WordWrap = True
                End If
            End With
        End With
    End Sub
    Private Shared Sub FormatDataGridViews(d As Form, theme As BaseTheme)
        For Each g As Control In d.Controls
            If TypeOf (g) Is DataGridView Then
                FormatDataGridView(g, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatDataGridView(g_ As DataGridView, theme As BaseTheme)
        On Error Resume Next
        With g_
            .BackgroundColor = BackgroundColorFromBaseTheme(theme)
            .DefaultCellStyle.BackColor = BackgroundColorFromBaseTheme(theme)
            .DefaultCellStyle.ForeColor = ForeGroundColorFromBaseTheme(theme)
            .ColumnHeadersHeight = 68
            .BorderStyle = System.Windows.Forms.BorderStyle.None
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False

            'If .Tag = "" Then .TabStop = False
            '.MultiSelect = False
        End With

    End Sub
    Private Shared Sub FormatComboBoxes(d As Form, theme As BaseTheme, sorted As Boolean)
        For Each c As Control In d.Controls
            If TypeOf (c) Is ComboBox Then
                FormatComboBox(c, theme, sorted)
            End If
        Next
    End Sub
    Private Shared Sub FormatComboBox(c As ComboBox, theme As BaseTheme, sorted As Boolean)
        With c
            .BackColor = BackgroundColorFromBaseTheme(theme)
            .ForeColor = ForeGroundColorFromBaseTheme(theme)
        End With
        Try
            c.AutoCompleteMode = AutoCompleteMode.Suggest
            c.AutoCompleteSource = AutoCompleteSource.ListItems
            c.Sorted = sorted
        Catch
        End Try
    End Sub
    Private Shared Sub FormatCheckBoxes(d As Form, theme As BaseTheme)
        For Each h As Control In d.Controls
            If TypeOf (h) Is CheckBox Then
                FormatCheckBox(h, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatCheckBox(h As CheckBox, theme As BaseTheme)
        With h
            .BackColor = Color.Transparent
            .ForeColor = ForeGroundColorFromBaseTheme(theme)
        End With
    End Sub
    Private Shared Sub FormatLabels(d As Form, theme As BaseTheme)
        For Each l As Control In d.Controls
            If TypeOf (l) Is Label Then
                FormatLabel(l, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatLabel(l As Label, theme As BaseTheme)
        With l
            .BackColor = Color.Transparent
            If l IsNot Title__ Then .ForeColor = ForeGroundColorFromBaseTheme(theme)

            'If theme = BaseTheme.Brown Then
            '    If l Is DialogCloseButton Or l Is MinimizeButton__ Then
            '        .ForeColor = Color.Green
            '    End If
            'End If

            'If l Is MaximizeButton__ Or l Is MinimizeButton__ Or l Is CloseButton__ Then .ForeColor = net_foreground_color(theme)
        End With
    End Sub
    Private Shared Sub FormatMenuStrips(d As Form, theme As BaseTheme)
        For Each m As Control In d.Controls
            If TypeOf (m) Is MenuStrip Then
                FormatMenuStrip(m, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatMenuStrip(m As MenuStrip, theme As BaseTheme)
        On Error Resume Next
        With m
            .BackColor = Color.Transparent
            .ForeColor = ForeGroundColorFromBaseTheme(theme)
        End With
    End Sub
    Private Shared Sub FormatPictureBoxes(d As Form, theme As BaseTheme)
        For Each p As Control In d.Controls
            If TypeOf (p) Is PictureBox Then
                FormatPictureBox(p, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatPictureBox(p As PictureBox, theme As BaseTheme)

        With p
            .BackgroundImageLayout = ImageLayout.Zoom
            .AllowDrop = True
            Try
                If p IsNot LeftBorder__ And p IsNot RightBorder__ And p IsNot TopBorder__ And p IsNot BottomBorder__ Then
                    p.BackColor = Color.Transparent
                ElseIf p Is LeftBorder__ Or p Is RightBorder__ Or p Is TopBorder__ Or p Is BottomBorder__ Then
                    p.BackColor = Color.Black
                End If
            Catch
            Finally
            End Try
        End With
    End Sub
    Private Shared Sub FormatNumericUpDowns(d As Form, theme As BaseTheme)
        For Each u_ As Control In d.Controls
            If TypeOf (u_) Is NumericUpDown Then
                FormatNumericUpDown(u_, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatNumericUpDown(n_ As NumericUpDown, theme As BaseTheme)
        With n_
            .BackColor = BackgroundColorFromBaseTheme(theme)
            .ForeColor = ForeGroundColorFromBaseTheme(theme)
        End With
    End Sub
    Private Shared Sub FormatRadios(d As Form, theme As BaseTheme)
        For Each r As Control In d.Controls
            If TypeOf (r) Is RadioButton Then
                FormatRadio(r, theme)
            End If
        Next
    End Sub
    Private Shared Sub FormatRadio(r As RadioButton, theme As BaseTheme)
        With r
            .BackColor = Color.Transparent
            .ForeColor = ForeGroundColorFromBaseTheme(theme)
        End With
    End Sub
    Private Shared Sub FormatListBoxes(d As Form, theme As BaseTheme, sorted As Boolean)
        For Each l_ As Control In d.Controls
            If TypeOf (l_) Is ListBox Then
                FormatListBox(l_, theme, sorted)
            End If
        Next
    End Sub
    Private Shared Sub FormatListBox(l_ As ListBox, theme As BaseTheme, sorted As Boolean)
        With l_
            .BackColor = BackgroundColorFromBaseTheme(theme)
            .ForeColor = ForeGroundColorFromBaseTheme(theme)
            .ScrollAlwaysVisible = True
            .HorizontalScrollbar = True

        End With
        Try
            l_.Sorted = sorted
        Catch x As Exception
        End Try
    End Sub
    Private Shared Sub ShieldControls(d As Form)

        On Error Resume Next
        For Each c_ As Control In d.Controls
            If c_.Left + c_.Width < 0 Or
                    c_.Top + c_.Height < 0 Or
                    c_.Left > d.Width Or
                    c_.Top > d.Height Then

                If TypeOf c_ Is Button Then
                    c_.TabStop = False
                ElseIf TypeOf c_ Is TextBox Then
                    c_.TabStop = False
                    Dim t_ As TextBox = c_
                    t_.ReadOnly = True
                ElseIf TypeOf c_ Is ComboBox Then
                    c_.TabStop = False
                    Dim b_ As ComboBox = c_
                ElseIf TypeOf c_ Is CheckBox Then
                    c_.TabStop = False
                    c_.Enabled = False
                ElseIf TypeOf c_ Is MenuStrip Then
                    c_.TabStop = False
                    c_.Enabled = False
                ElseIf TypeOf c_ Is NumericUpDown Then
                    c_.TabStop = False
                    c_.Enabled = False
                ElseIf TypeOf c_ Is ContextMenuStrip Then
                    c_.Enabled = False
                ElseIf TypeOf c_ Is DateTimePicker Then
                    c_.TabStop = False
                    c_.Enabled = False
                ElseIf TypeOf c_ Is RadioButton Then
                    c_.TabStop = False
                    c_.Enabled = False
                ElseIf TypeOf c_ Is RichTextBox Then
                    c_.TabStop = False
                    Dim rtb_ As TextBox = c_
                    rtb_.ReadOnly = True
                ElseIf TypeOf c_ Is TabControl Then
                    c_.TabStop = False
                ElseIf TypeOf c_ Is DataGridView Then
                    c_.TabStop = False
                ElseIf TypeOf c_ Is Panel Then
                    c_.TabStop = False
                End If
            End If
        Next
    End Sub

#End Region

End Class
