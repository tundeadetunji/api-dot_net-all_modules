'from iNovation
Imports iNovation.Code.General
Imports iNovation.Code.Values


'from Assemblies
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Net
Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Security.AccessControl
Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles

''' <summary>
''' This class contains methods geared mainly towards desktop development, the part the end-user interacts with directly.
''' Reference to System.Windows.Forms may be required.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public Class Desktop

#Region "Desktop"

    Public Shared Function GetRunningPrograms(Optional sought As SideToReturn = SideToReturn.AsCustomApplicationInfo, Optional only_if_main_window_title_is_not_blank As Boolean = False, Optional machine_name As String = Nothing)

        Dim p()
        If machine_name IsNot Nothing Then
            If machine_name.Length > 0 Then p = Process.GetProcesses(machine_name)
        Else
            p = Process.GetProcesses()
        End If

        If p.Length < 1 Then Return Nothing

        Dim temp As New List(Of String)
        Dim strings As New List(Of String)
        Dim infos As New List(Of CustomApplicationInfo)

        For i = 0 To p.Length - 1
            If temp.Contains(p(i).ProcessName) = False Then
                If only_if_main_window_title_is_not_blank Then
                    If p(i).MainWindowTitle.Length > 0 Then
                        Try
                            If sought = SideToReturn.AsCustomApplicationInfo Then
                                Dim info As New CustomApplicationInfo
                                info.Filename = p(i).MainModule.FileName
                                info.DisplayName = p(i).MainWindowTitle
                                info.ProcessName = p(i).ProcessName
                                infos.Add(info)
                            ElseIf sought = SideToReturn.AsCustomApplicationInfoFilename Then
                                strings.Add(p(i).MainModule.FileName)
                            ElseIf sought = SideToReturn.AsCustomApplicationInfoDisplayName Then
                                strings.Add(p(i).MainWindowTitle)
                            ElseIf sought = SideToReturn.AsCustomApplicationInfoProcessName Then
                                strings.Add(p(i).ProcessName)
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Else
                    Try
                        If sought = SideToReturn.AsCustomApplicationInfo Then
                            Dim info As New CustomApplicationInfo
                            info.Filename = p(i).MainModule.FileName
                            info.DisplayName = p(i).MainWindowTitle
                            info.ProcessName = p(i).ProcessName
                            infos.Add(info)
                        ElseIf sought = SideToReturn.AsCustomApplicationInfoFilename Then
                            strings.Add(p(i).MainModule.FileName)
                        ElseIf sought = SideToReturn.AsCustomApplicationInfoDisplayName Then
                            strings.Add(p(i).MainWindowTitle)
                        ElseIf sought = SideToReturn.AsCustomApplicationInfoProcessName Then
                            strings.Add(p(i).ProcessName)
                        End If
                    Catch ex As Exception
                    End Try
                End If
                temp.Add(p(i).ProcessName)

            End If
        Next

        If sought = SideToReturn.AsCustomApplicationInfo Then
            Return infos
        Else
            Return strings
        End If

    End Function

    Public Shared Function GetInstalledPrograms(Optional sought As SideToReturn = SideToReturn.AsCustomApplicationInfo, Optional only_if_main_window_title_is_not_blank As Boolean = False)
        Dim infos As New List(Of CustomApplicationInfo)
        Dim strings As New List(Of String)

        Dim registry_key As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

        Using key As Microsoft.Win32.RegistryKey = Registry.LocalMachine.OpenSubKey(registry_key)

            For Each subkey_name As String In key.GetSubKeyNames()

                Using subkey As RegistryKey = key.OpenSubKey(subkey_name)

                    If only_if_main_window_title_is_not_blank Then
                        If subkey.GetValue("DisplayName") IsNot Nothing Then
                            If subkey.GetValue("DisplayName").ToString.Length > 0 Then
                                If sought = SideToReturn.AsCustomApplicationInfo Then
                                    Dim app As CustomApplicationInfo
                                    app.DisplayName = subkey.GetValue("DisplayName")
                                    app.InstallLocation = subkey.GetValue("InstallLocation")
                                    infos.Add(app)
                                ElseIf sought = SideToReturn.AsCustomApplicationInfoDisplayName Then
                                    strings.Add(subkey.GetValue("DisplayName"))
                                ElseIf sought = SideToReturn.AsCustomApplicationInfoInstallLocation Then
                                    strings.Add(subkey.GetValue("InstallLocation"))
                                End If

                            End If
                        End If
                    Else
                        If sought = SideToReturn.AsCustomApplicationInfo Then
                            Dim app As CustomApplicationInfo
                            app.DisplayName = subkey.GetValue("DisplayName")
                            app.InstallLocation = subkey.GetValue("InstallLocation")
                            infos.Add(app)
                        ElseIf sought = SideToReturn.AsCustomApplicationInfoDisplayName Then
                            Dim display_name As String = subkey.GetValue("DisplayName")
                            If display_name.Length > 1 Then
                                strings.Add(display_name)
                            End If
                        ElseIf sought = SideToReturn.AsCustomApplicationInfoInstallLocation Then
                            strings.Add(subkey.GetValue("InstallLocation"))
                        End If
                    End If


                End Using
            Next
        End Using

        If sought = SideToReturn.AsCustomApplicationInfo Then
            Return infos
        Else
            Return strings
        End If
    End Function

    Public Shared Sub HideFileOrFolder(file_or_folder As String, Optional should As HideFileOrFolderShould = HideFileOrFolderShould.SealTheResource)
        If Exists(file_or_folder) = False Then
            Return
        End If

        Select Case should
            Case HideFileOrFolderShould.HideTheResource
                Try

                    SetAttr(file_or_folder, FileAttribute.Hidden)
                Catch ex As Exception

                End Try
            Case HideFileOrFolderShould.SealTheResource
                Try

                    SetAttr(file_or_folder, FileAttribute.Hidden + FileAttribute.System)
                Catch ex As Exception

                End Try
            Case HideFileOrFolderShould.MakeTheResourceVisible
                Try

                    SetAttr(file_or_folder, FileAttribute.Normal)
                Catch ex As Exception

                End Try
        End Select


    End Sub

    ''' <summary>
    ''' Binds the values of Enum to listbox or combobox, while replacing any underscores in the values with the space character.
    ''' </summary>
    ''' <param name="drop">ListBox or ComboBox</param>
    ''' <param name="enum__">New instance of desired Enum</param>
    ''' <param name="return_control_not_list">Should return the original ListBox/ComboBox or the list generated from the enum</param>
    ''' <param name="return_sorted">Should the list generated from the enum be sorted before bound to the control?</param>
    ''' <returns></returns>
    Public Shared Function EnumDrop(drop As Control, enum__ As Object, Optional return_control_not_list As Boolean = True, Optional return_sorted As Boolean = True, Optional InitialSelectedIndexIsNegativeOne As Boolean = True) As Object
        Dim e = GetEnum(enum__)
        Dim l As New List(Of String)
        With e
            For i = 0 To .Count - 1
                l.Add(EnumToDrop(e(i)))
            Next
        End With

        If return_sorted Then l.Sort()

        Return If(return_control_not_list, BindProperty(drop, l, InitialSelectedIndexIsNegativeOne), l)
    End Function
    Public Shared Function ListsToKeyValueArray(listK As ListBox, listV As ListBox) As List(Of String)
        Dim s As String = ""
        If IsEmpty({listK, listV}, ControlsToCheck.Any) = True Then
            Return Nothing
            Exit Function
        End If
        Dim l As New List(Of String)
        With listK
            For i = 0 To .Items.Count - 1
                l.Add(.Items.Item(i))
                l.Add(listV.Items.Item(i))
            Next
        End With
        Return l
    End Function
    Public Shared Function ListToArray(l As ListBox, return_as As ReturnInfo) As Object
        If l Is Nothing Then Return Nothing
        If l.Items.Count < 1 Then Return Nothing
        Dim l_ As New List(Of String)
        With l.Items
            For i As Integer = 0 To .Count - 1
                l_.Add(.Item(i).ToString)
            Next
        End With
        Select Case return_as
            Case ReturnInfo.AsListOfString
                Return l_
            Case Else
                Return l_.ToArray
        End Select
    End Function
    Public Shared Function ListToArray(l As ComboBox, return_as As ReturnInfo) As Object
        If l Is Nothing Then Return Nothing
        If l.Items.Count < 1 Then Return Nothing
        Dim l_ As New List(Of String)
        With l.Items
            For i As Integer = 0 To .Count - 1
                l_.Add(.Item(i).ToString)
            Next
        End With
        Select Case return_as
            Case ReturnInfo.AsListOfString
                Return l_
            Case Else
                Return l_.ToArray
        End Select
    End Function

    Public suffx
    Public prefx
    Public countr
    Public TextHasSpace As Boolean
    Public strSource


    ''' <summary>
    ''' Changes text to title case. Called from _TextChanged. Multiline is not supported yet.
    ''' </summary>
    ''' <param name="strSource"></param>
    Public Shared Sub ToTitleCase(ByRef strSource As Control)
        '        Dim g As New FormatWindow
        Try
            ConvertTextToTitleCase(strSource)
        Catch
        End Try
    End Sub


    ''' <summary>
    ''' Use ToTitleCase instead.
    ''' </summary>
    ''' <param name="strSource"></param>
    Private Shared Sub ConvertTextToTitleCase(ByRef strSource As Control)
        Dim d As New Desktop

        'convert text to title case
        'called from TextChanged
        If TypeOf strSource Is TextBox Then
            Dim t As TextBox = strSource
            If t.Multiline = True Then Exit Sub
        End If

        On Error Resume Next
        If Len(strSource.Text) = 0 Then
            d.suffx = ""
            d.prefx = ""
            d.countr = Nothing
            Exit Sub
        End If

        If Len(strSource.Text) = 1 Then strSource.Text = UCase(strSource.Text) : System.Windows.Forms.SendKeys.Send("{End}")

        If Mid(strSource.Text, Len(strSource.Text), 1) = " " Or Mid(strSource.Text, Len(strSource.Text), 1) = "." Or Mid(strSource.Text, Len(strSource.Text), 1) = ChrW(13) Then
            d.TextHasSpace = True
            d.countr = Len(strSource.Text) + 1
            d.prefx = strSource.Text
        End If

        If d.prefx <> "" And Len(strSource.Text) = Val(d.countr) Then
            d.suffx = UCase(Mid(strSource.Text, Len(strSource.Text), 1))
            strSource.Text = d.prefx & d.suffx
            System.Windows.Forms.SendKeys.Send("{End}")
            'clear counters
            d.suffx = ""
            d.prefx = ""
            d.countr = Nothing
        End If
    End Sub

    Public Sub SetTextChange(c As Control, placeholder As String, placeholderLabel As Label, Mode As String, IsUsername As Boolean)

        If Trim(c.Text) = "" Then
            c.Text = placeholder
            placeholderLabel.Visible = False
        ElseIf Trim(c.Text) <> "" And c.Text <> placeholder Then
            placeholderLabel.Visible = True
        End If

        Select Case Mode.ToLower
            Case "email"
                If TypeOf c Is TextBox Then
                    Dim t As TextBox = c
                    If Trim(t.Text).Length > 0 And t.Multiline = False And IsUsername = False Then
                        ConvertTextToLowerCase(c)
                    End If
                ElseIf TypeOf c Is ComboBox Then
                    If Trim(c.Text).Length > 0 Then ConvertTextToLowerCase(c)
                End If
            Case Else
                If TypeOf c Is TextBox Then
                    Dim t As TextBox = c
                    If Trim(t.Text).Length > 0 And t.Multiline = False And IsUsername = False Then
                        ConvertTextToTitleCase(c)
                    End If
                ElseIf TypeOf c Is ComboBox Then
                    If Trim(c.Text).Length > 0 Then ConvertTextToTitleCase(c)
                End If
        End Select
    End Sub

    Public Sub SetUsernameTextChange(c As Control, placeholder As String)
        If TypeOf (c) Is TextBox Or TypeOf (c) Is ComboBox Then
            If c.Text = "" Then c.Text = placeholder
        End If
    End Sub
    Public Sub SetPasswordChange(c As TextBox, Optional placeholder As String = "Password")
        If c.Text = "" Then c.Text = placeholder : c.PasswordChar = Nothing : Exit Sub
        If c.Text.ToLower <> placeholder Then c.PasswordChar = "*" : Exit Sub
    End Sub
    Public Sub UsernameTextChange(c As Control, Optional placeholder As String = "Username")
        If TypeOf (c) Is TextBox Or TypeOf (c) Is ComboBox Then
            If c.Text = "" Then c.Text = placeholder
        End If
    End Sub
    Public Sub PasswordTextChange(c As TextBox, Optional placeholder As String = "Password")
        If c.Text = "" Then c.Text = placeholder : c.PasswordChar = Nothing : Exit Sub
        If c.Text.ToLower <> placeholder Then c.PasswordChar = "*" : Exit Sub
    End Sub
    ''' <summary>
    ''' Checks if Enter key was pressed. Called from _KeyPress.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <returns></returns>

    Public Shared Function EnterKeyWasPressed(ByRef e As System.Windows.Forms.KeyPressEventArgs) As Boolean
        Return e.KeyChar = ChrW(13)
    End Function

    ''' <summary>
    ''' Allows numbers and % only. Called from _KeyPress.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <param name="controlText"></param>
    Public Shared Sub AllowPercentage(ByRef e As System.Windows.Forms.KeyPressEventArgs, controlText As String)
        '        On Error Resume Next
        'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127), percent(37) and minus(45) only
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Or e.KeyChar = ChrW(45) Or e.KeyChar = ChrW(37) Then
            Exit Sub
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub

    ''' <summary>
    ''' Allows numbers(48-57), backspace(8), enter(13), full stop(46), delete(127) and minus(45) only. Called from control_KeyPress.
    ''' </summary>
    ''' <param name="e"></param>
    Public Shared Sub AllowNumberOnly(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        '        On Error Resume Next
        'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127) and minus(45) only
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Or e.KeyChar = ChrW(45) Then
            Exit Sub
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub
    Public Shared Sub AllowFileNameCharacterOnly(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        'alphabets, numbers, underscore, backspace
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or (e.KeyChar >= ChrW(65) And e.KeyChar <= ChrW(90)) Or (e.KeyChar >= ChrW(97) And e.KeyChar <= ChrW(122)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(95) Then
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub

    ''' <summary>
    ''' Discontinued. Do not use it.
    ''' </summary>
    ''' <param name="e"></param>
    Private Sub AllowPIN(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        '        On Error Resume Next
        'allow numbers (48-57), backspace(8), enter(13), delete(127) only
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Then
            Exit Sub
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub

    ''' <summary>
    ''' Allows numbers only. Called from _KeyPress. Good for PIN.
    ''' </summary>
    ''' <param name="e"></param>

    Public Shared Sub AllowDigitOnly(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        '        On Error Resume Next
        'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127) only
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Then
            Exit Sub
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub

    ''' <summary>
    ''' Allows nothing.
    ''' </summary>
    ''' <param name="e"></param>
    Public Shared Sub AllowNothing(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        e.KeyChar = ChrW(0)
    End Sub
    ''' <summary>
    ''' Changes text to uppercase. Called from _TextChanged.
    ''' </summary>
    ''' <param name="strSource"></param>
    Public Shared Sub ToUpperCase(ByRef strSource As Control)
        '		On Error Resume Next
        'called from TextChanged
        If Len(strSource.Text) = 0 Then Exit Sub
        Try
            strSource.Text = UCase(strSource.Text)
        Catch
        End Try
    End Sub
    ''' <summary>
    ''' Changes text to lower case. Called from _TextChanged. Useful for email addresses and websites.
    ''' </summary>
    ''' <param name="strSource"></param>
    Public Shared Sub ToLowerCase(ByRef strSource As Control)
        Try
            ConvertTextToLowerCase(strSource)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Use ToLowerCase instead.
    ''' </summary>
    ''' <param name="strSource"></param>
    Private Shared Sub ConvertTextToLowerCase(ByRef strSource As Control)
        On Error Resume Next
        'convert text to lower case (useful for email and websites addresses)
        'called from TextChanged
        If Len(strSource.Text) = 0 Then Exit Sub

        strSource.Text = LCase(strSource.Text)

    End Sub

    Public Sub NoKey(sender As Object, e As KeyPressEventArgs)
        AllowNothing(e)
    End Sub

    Public Shared Function ToCurrency(val_)
        Try
            Return RoundNumber(Val(val_))
        Catch
        End Try
    End Function

    Public Shared Sub SendFocusToControl(theControl As Control)
        Try
            theControl.Focus()
            SendKeys.Send("{End}")
        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' Displays a message box, accompanied by voice feedback.
    ''' </summary>
    ''' <param name="message_box_message"></param>
    ''' <param name="voice_feedback_message"></param>
    ''' <param name="MessageBoxStyle"></param>
    ''' <param name="title"></param>
    ''' <param name="voice_feedback_is_async"></param>
    ''' <returns></returns>
    Public Shared Function mFeedback(message_box_message As String, Optional voice_feedback_message As String = "", Optional MessageBoxStyle As MsgBoxStyle = MsgBoxStyle.Information + MsgBoxStyle.OkOnly, Optional title As String = "", Optional voice_feedback_is_async As Boolean = True) As MsgBoxResult
        If voice_feedback_message.Trim.Length > 0 Then
            Dim feedback As New iNovation.Code.Feedback
            feedback.say(voice_feedback_message, voice_feedback_is_async)
        End If
        Return MsgBox(message_box_message, MessageBoxStyle, title)
    End Function
    ''' <summary>
    ''' Same as mFeedback(String, String, MsgBoxStyle, String, Boolean)
    ''' </summary>
    ''' <param name="message_box_message"></param>
    ''' <param name="voice_feedback_message"></param>
    ''' <param name="MessageBoxStyle"></param>
    ''' <param name="title"></param>
    ''' <param name="voice_feedback_is_async"></param>
    ''' <returns></returns>
    Public Shared Function messageFeedback(message_box_message As String, Optional voice_feedback_message As String = "", Optional MessageBoxStyle As MsgBoxStyle = MsgBoxStyle.Information + MsgBoxStyle.OkOnly, Optional title As String = "", Optional voice_feedback_is_async As Boolean = True) As MsgBoxResult
        Return mFeedback(message_box_message, voice_feedback_message, MessageBoxStyle, title, voice_feedback_is_async)
    End Function
    Public Shared Sub StatusDrop(d_ As ComboBox, Optional IsUpdate As Boolean = False, Optional FirstIsEmpty As Boolean = False, Optional sort_ As Boolean = True)
        If d_.Items.Count > 0 Then Exit Sub

        Clear(d_)

        If FirstIsEmpty Then d_.Items.Add("")

        Dim l As List(Of String) = StatusList(IsUpdate)
        For i As Integer = 0 To l.Count - 1
            d_.Items.Add(l(i).ToString)
        Next
        d_.Sorted = sort_
    End Sub

    Public Shared Sub EnableControls(control_ As Array, Optional state_ As Boolean = True)
        On Error Resume Next
        If control_.Length < 1 Then Exit Sub
        For i As Integer = 0 To control_.Length - 1
            control_(i).Enabled = state_
        Next

    End Sub

    Public Shared Sub EnableControl(control As Object, Optional state_ As Boolean = True)
        On Error Resume Next
        control.Enabled = state_

    End Sub

    Public Shared Sub DisableControl(control As Control)
        EnableControl(control, False)

    End Sub


    ''' <summary>
    ''' Populates ComboBox with numbers.
    ''' </summary>
    ''' <param name="d_">ComboBox to fill</param>
    ''' <param name="end_">Ending Number</param>
    ''' <param name="start_">Beginning Number</param>
    ''' <param name="firstItemIsEmpty"></param>
    Public Shared Sub NumberDrop(d_ As ComboBox, end_ As Integer, Optional start_ As Integer = 0, Optional firstItemIsEmpty As Boolean = False)
        If d_.Items.Count > 0 Then Exit Sub

        With d_.Items
            Try
                If firstItemIsEmpty = True And d_.DataSource Is Nothing Then .Add("")
            Catch
            End Try
            For i As Integer = start_ To end_
                .Add(i)
            Next
        End With
    End Sub

    Public Shared Function DropToList(drop_or_list As Object) As List(Of String)
        Dim l As New List(Of String)
        If drop_or_list.Items.Count = 0 Then
            Return l
            Exit Function
        Else
            With drop_or_list
                .SelectedIndex = 0
                For i = 0 To .Items.Count - 1
                    l.Add(.Items.Item(i).ToString)
                    'If i <> .Items.Count - 1 Then
                    '    .SelectedIndex += 1
                    'End If
                Next
            End With

        End If
        Return l
    End Function


    ''' <summary>
    ''' Determines if all controls and/or file(s) do not have text.
    ''' </summary>
    ''' <param name="controls_">Controls or File paths or both</param>
    ''' <returns>True or False</returns>

    Public Shared Function IsEmpty(controls_ As Array, Optional controls_to_check As ControlsToCheck = ControlsToCheck.Any) As Boolean
        Dim counter_ As Integer = 0
        With controls_
            For i As Integer = 0 To .Length - 1
                If IsEmpty(controls_(i), True, "", "", Nothing) Then
                    counter_ += 1
                End If
            Next
        End With
        Return If(controls_to_check = ControlsToCheck.All, Val(counter_) = controls_.Length, Val(counter_) > 0)
    End Function

    ''' <summary>
    ''' Determines if control or file does not have text (or items, if it's listbox, array, list of string or object or integer or double, collection).
    ''' </summary>
    ''' <param name="c_">File path or ComboBox or TextBox or PictureBox or NumericUpDown</param>
    ''' <param name="use_trim">Should content be trimmed before check?</param>
    ''' <param name="control_to_focus">Focus on this when IsEmpty is True</param>
    ''' <param name="string_feedback">What to say to user</param>
    ''' <returns>True or False</returns>
    Public Shared Function IsEmpty(c_ As Object, Optional use_trim As Boolean = True, Optional string_feedback As String = "", Optional app_ As String = "", Optional control_to_focus As Control = Nothing) As Boolean
        Dim n__ As NumericUpDown
        Dim p__ As PictureBox
        ''        Dim d__ As ComboBox
        Dim t__ As TextBox
        Dim l__ As ListBox

        Dim return_val As Boolean = False

        Dim a As Array
        Dim l_string As List(Of String)
        Dim l_object As List(Of Object)
        Dim l_integer As List(Of Integer)
        Dim l_double As List(Of Double)
        Dim c As Collection

        If TypeOf c_ Is Array Then
            a = c_
            Return a.Length < 1
        End If

        If TypeOf c_ Is List(Of String) Then
            l_string = c_
            Return l_string.Count < 1
        End If

        If TypeOf c_ Is List(Of Object) Then
            l_object = c_
            Return l_object.Count < 1
        End If

        If TypeOf c_ Is List(Of Integer) Then
            l_integer = c_
            Return l_integer.Count < 1
        End If

        If TypeOf c_ Is List(Of Double) Then
            l_double = c_
            Return l_double.Count < 1
        End If

        If TypeOf c_ Is Collection Then
            c = c_
            Return c.Count < 1
        End If

        If TypeOf c_ Is PictureBox Then
            p__ = c_
            If p__.BackgroundImage Is Nothing And p__.Image Is Nothing Then
                return_val = True
            End If
        ElseIf TypeOf c_ Is NumericUpDown Then
            n__ = c_
            If (n__.Value) = 0 Then
                return_val = True
            End If
        ElseIf TypeOf c_ Is ListBox Then
            l__ = c_
            If l__.Items.Count < 1 Then
                return_val = True
            End If
        ElseIf TypeOf c_ Is ComboBox Then
            If use_trim = True Then
                If CType(c_, ComboBox).Text.Trim.Length < 1 Then
                    return_val = True
                End If
            Else
                If CType(c_, ComboBox).Text.Length < 1 Then
                    return_val = True
                End If
            End If
        ElseIf TypeOf c_ Is TextBox Then
            t__ = c_
            If use_trim = True Then
                If t__.Text.Trim.Length < 1 Then
                    return_val = True
                End If
            Else
                If t__.Text.Length < 1 Then
                    return_val = True
                End If
            End If
        ElseIf TypeOf c_ Is CheckBox Then
            Return CType(c_, CheckBox).Checked = False
        ElseIf CType(c_, ComboBox).Text.Trim.Length < 1 Then
            return_val = True
        Else
            Try
                If ReadText(c_).Length < 1 Then
                    return_val = True
                End If
            Catch ex As Exception
            End Try
        End If
        'If return_val = True And string_feedback.Length > 0 And app_.Length > 0 Then
        '    Try
        '        Dim feedback_ As New Feedback()
        '        feedback_.fFeedback(string_feedback)
        '    Catch
        '    End Try
        'End If
        If control_to_focus IsNot Nothing Then
            Try
                control_to_focus.Focus()
            Catch ex As Exception
            End Try
        ElseIf control_to_focus Is Nothing Then
            Try
                c_.Focus()
            Catch ex As Exception
            End Try
        End If
        Return return_val
    End Function

    ''' <summary>
    ''' Attaches the Titles enum to ComboBox or ListBox.
    ''' </summary>
    ''' <param name="c_">ComboBox or ListBox</param>
    ''' <param name="InitialSelectedIndexIsNegativeOne">Should the SelectedIndex be set to -1? Only applies to ComboBox</param>
    ''' <param name="ReplaceUnderscoresWithSpace">If using enum with values having more than one word separted by underscore, this replaces the underscores with space</param>
    ''' <returns></returns>
    Public Shared Function TitleDrop(c_ As Control, Optional InitialSelectedIndexIsNegativeOne As Boolean = True, Optional ReplaceUnderscoresWithSpace As Boolean = True) As Control
        Return BindProperty(c_, GetEnum(New Title), InitialSelectedIndexIsNegativeOne, ReplaceUnderscoresWithSpace)
    End Function

    ''' <summary>
    ''' Attaches the Genders enum to ComboBox or ListBox.
    ''' </summary>
    ''' <param name="c_">ComboBox or ListBox</param>
    ''' <param name="InitialSelectedIndexIsNegativeOne">Should the SelectedIndex be set to -1? Only applies to ComboBox</param>
    ''' <param name="ReplaceUnderscoresWithSpace">If using enum with values having more than one word separted by underscore, this replaces the underscores with space</param>
    ''' <returns></returns>
    Public Shared Function GenderDrop(c_ As Control, Optional InitialSelectedIndexIsNegativeOne As Boolean = True, Optional ReplaceUnderscoresWithSpace As Boolean = True) As Control
        Return BindProperty(c_, GetEnum(New Gender), InitialSelectedIndexIsNegativeOne, ReplaceUnderscoresWithSpace)
    End Function

    ''' <summary>
    ''' Grabs the picture inside PictureBox. Same as passing PictureBox as src to PictureFromStream().
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="file_extension"></param>
    ''' <param name="PropertyIsImage"></param>
    ''' <returns></returns>
    Public Shared Function Content(control_ As PictureBox, Optional file_extension As FileType = FileType.jpg, Optional PropertyIsImage As Boolean = False) As Object
        Return PictureFromStream(control_, file_extension, PropertyIsImage)
        'Try
        '    If use_image Then
        '        Return PictureFromStream(control_, ".jpg", True)
        '    Else
        '        Return PictureFromStream(control_)
        '    End If
        'Catch ex As Exception

        'End Try

    End Function

    'Public Shared Property casing__ As TextCase = TextCase.Capitalize

    ''' <summary>
    ''' Gets the content of the control or file.
    ''' </summary>
    ''' <param name="control_">NumericUpDown, TrackBar, CheckBox, DateTimePicker, ListBox, ComboBox or string representing path to file</param>
    ''' <param name="casing">Should the output's casing be upper, lower, capitalized or as is?</param>
    ''' <param name="timeValue">Which part of the value to pick if control is DateTimePicker (e.g. Day, Hour, Minute etc)</param>
    ''' <returns></returns>
    Public Shared Function Content(control_ As Object, Optional casing As TextCase = TextCase.None, Optional trimOutput As Boolean = True, Optional timeValue As TimeValue = TimeValue.ShortDate) As String

        Dim d As DateTimePicker
        '		Dim html_text As HtmlInputText
        Try
            If TypeOf control_ Is NumericUpDown Then
                Return TransformText(CType(control_, NumericUpDown).Value, casing)
            ElseIf TypeOf control_ Is TrackBar Then
                Return TransformText(CType(control_, NumericUpDown).Value, casing)
            ElseIf TypeOf control_ Is CheckBox Then
                Return TransformText(CType(control_, CheckBox).Checked, casing)
            ElseIf TypeOf control_ Is DateTimePicker Then
                d = control_
                Select Case timeValue
                    Case TimeValue.Day
                        Return TransformText(CType(control_, DateTimePicker).Value.Day, casing)
                    Case TimeValue.Hour
                        Return TransformText(CType(control_, DateTimePicker).Value.Hour, casing)
                    Case TimeValue.LongDate
                        Return TransformText(CType(control_, DateTimePicker).Value.ToLongDateString, casing)
                    Case TimeValue.LongTime
                        Return TransformText(CType(control_, DateTimePicker).Value.ToLongTimeString, casing)
                    Case TimeValue.Millisecond
                        Return TransformText(CType(control_, DateTimePicker).Value.Millisecond, casing)
                    Case TimeValue.Minute
                        Return TransformText(CType(control_, DateTimePicker).Value.Minute, casing)
                    Case TimeValue.Month
                        Return TransformText(CType(control_, DateTimePicker).Value.Month, casing)
                    Case TimeValue.Second
                        Return TransformText(CType(control_, DateTimePicker).Value.Second, casing)
                    Case TimeValue.ShortDate
                        Return TransformText(CType(control_, DateTimePicker).Value.ToShortDateString, casing)
                    Case TimeValue.ShortTime
                        Return TransformText(CType(control_, DateTimePicker).Value.ToShortTimeString, casing)
                    Case TimeValue.Year
                        Return TransformText(CType(control_, DateTimePicker).Value.Year, casing)
                    Case TimeValue.DayOfWeek
                        Return TransformText(CType(control_, DateTimePicker).Value.DayOfWeek, casing)
                    Case TimeValue.DayOfYear
                        Return TransformText(CType(control_, DateTimePicker).Value.DayOfYear, casing)
                End Select
            ElseIf TypeOf control_ Is ComboBox Then
                Return TransformText(If(trimOutput, CType(control_, ComboBox).Text.Trim, CType(control_, ComboBox).Text), casing)
            ElseIf TypeOf control_ Is ListBox Then
                Try
                    Return TransformText(If(CType(control_, ListBox).SelectedIndex >= 0, If(trimOutput, CType(control_, ListBox).SelectedItem.ToString.Trim, CType(control_, ListBox).SelectedItem), ""), casing)
                Catch ex As Exception

                End Try
            ElseIf TypeOf control_ Is String Then
                Return TransformText(ReadText(control_, trimOutput), casing)
            Else
                Try
                    Return TransformText(control_.Text, casing)
                Catch ex As Exception
                End Try
            End If
        Catch
        End Try
    End Function

    Public Shared Sub Clear(c_ As Array, Optional initial_string As String = "")
        If c_.Length > 0 Then
            With c_
                For i = 0 To .Length - 1
                    Clear(c_(i), initial_string)
                Next
            End With
        End If
    End Sub

    Public Shared Sub Clear(c_ As Object, Optional initial_string As String = "")

        Dim c As ComboBox
        Dim l As ListBox
        Dim t As TextBox
        Dim p As PictureBox
        Dim h As CheckBox
        Dim g As DataGridView

        If TypeOf c_ Is DataGridView Then
            g = c_
            Try
                g.DataSource = Nothing
            Catch
            End Try
        ElseIf TypeOf c_ Is CheckBox Then
            h = c_
            h.Checked = False
        ElseIf TypeOf c_ Is ComboBox Then
            c = c_
            c.Text = initial_string
        ElseIf TypeOf c_ Is ListBox Then
            l = c_
            Try
                l.DataSource = Nothing
            Catch ex As Exception
            End Try
            l.Items.Clear()
        ElseIf TypeOf c_ Is TextBox Then
            t = c_
            t.Text = initial_string
        ElseIf TypeOf c_ Is PictureBox Then
            p = c_
            Try
                p.Image = Nothing
            Catch ex As Exception
            End Try
            Try
                p.BackgroundImage = Nothing
            Catch ex As Exception
            End Try
        ElseIf TypeOf c_ Is NumericUpDown Then
            Try
                If IsNumeric(initial_string) Then '/
                    CType(c_, NumericUpDown).Value = initial_string
                Else
                    CType(c_, NumericUpDown).Value = CType(c_, NumericUpDown).Minimum
                End If
            Catch ex As Exception

            End Try
        Else
            WriteText(c_, initial_string)
        End If

        Try
            SendFocusToControl(c_)
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Sub Empty(c_ As Array, Optional initial_string As String = "")
        If c_.Length > 0 Then
            With c_
                For i = 0 To .Length - 1
                    Empty(c_(i), initial_string)
                Next
            End With
        End If
    End Sub

    Public Shared Sub Empty(c_ As Object, Optional initial_string As String = "")

        Dim c As ComboBox
        Dim l As ListBox

        If TypeOf c_ Is ComboBox Then
            c = c_
            c.DataSource = Nothing
            c.Items.Clear()
            c.Text = initial_string
        ElseIf TypeOf c_ Is ListBox Then
            l = c_
            Try
                l.DataSource = Nothing
            Catch ex As Exception
            End Try
            l.Items.Clear()
        End If
    End Sub

    'Public Shared Sub GenderDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = False)

    '    If d_.Items.Count > 0 Then Exit Sub

    '    Clear(d_)

    '    If FirstIsEmpty Then d_.Items.Add("")

    '    Dim l As List(Of String) = GenderList()
    '    For i As Integer = 0 To l.Count - 1
    '        d_.Items.Add(l(i).ToString)
    '    Next
    'End Sub


    ''' <summary>
    ''' Populates drop with drop-text version of boolean, depending on the pattern. Same as PopulateBooleanDrop.
    ''' </summary>
    ''' <param name="d_">ComboBox to fill</param>
    ''' <param name="firstItemIsEmpty">Should 's first item be empty?</param>
    ''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
    'Public Shared Sub BooleanDrop(d_ As ComboBox, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = False)
    '    BooleanDrop(d_, pattern_, firstItemIsEmpty)

    'End Sub

    Public Shared Sub BooleanDrop(d_ As ComboBox, Optional pattern_ As DropTextPattern = DropTextPattern.AlwaysNever, Optional firstItemIsEmpty As Boolean = False)
        With d_
            If .Items.Count > 0 Then Exit Sub
            With .Items
                If firstItemIsEmpty = True Then .Add("")
                .Add(BooleanToDropText(True, pattern_))
                .Add(BooleanToDropText(False, pattern_))
            End With
            Try
                .Text = ""
                .SelectedIndex = -1
            Catch ex As Exception

            End Try
        End With

    End Sub

    ''' <summary>
    ''' Populates drop with drop-text version of boolean, depending on the pattern.
    ''' </summary>
    ''' <param name="d_">ComboBox to fill</param>
    ''' <param name="firstItemIsEmpty">Should 's first item be empty?</param>
    ''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
    Private Sub PopulateBooleanDrop(d_ As ComboBox, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = False)
        With d_
            If .Items.Count > 0 Then Exit Sub
            With .Items
                If firstItemIsEmpty = True Then .Add("")
                .Add(BooleanToDropText(True, pattern_))
                .Add(BooleanToDropText(False, pattern_))
            End With
            Try
                .Text = ""
                .SelectedIndex = -1
            Catch ex As Exception

            End Try
        End With
    End Sub

    Public Sub PopulateBEDrop(d_ As ComboBox, Optional firstItemIsEmpty As Boolean = False)
        If d_.Items.Count > 0 Then Exit Sub

        With d_.Items
            Try
                If firstItemIsEmpty = True And d_.DataSource Is Nothing Then .Add("")
            Catch
            End Try
            .Add("Begin")
            .Add("End")
        End With
    End Sub

    Public Function Toboolean(str_)
        Convert.ToBoolean(Convert.ToInt32(str_))
    End Function

    Public Shared Function AddCells(grid_with_quantity_price_IF_POSSIBLE As DataGridView, Optional isCart As Boolean = False, Optional quantity_i As Integer = Nothing, Optional price_i As Integer = Nothing)
        Dim g As DataGridView = grid_with_quantity_price_IF_POSSIBLE
        Dim counter = 0
        Dim q = 0, p = 0, q_i = 0, p_i = 1
        If quantity_i <> Nothing Then q_i = quantity_i
        If price_i <> Nothing Then p_i = price_i
        With g
            If .Rows.Count < 1 Then Return 0
            For i = 0 To .Rows.Count - 1
                q = Val(.Rows(i).Cells(q_i).Value)
                p = Val(.Rows(i).Cells(p_i).Value)
                If isCart Then
                    counter += (q * p)
                Else
                    counter += q
                End If
            Next
        End With
        Return counter
    End Function

    ''' <summary>
    ''' Gets items in combobox and adds them to a list of string
    ''' </summary>
    ''' <param name="c"></param>
    ''' <param name="returnAs">List(Of String) OR Array</param>
    ''' <returns>List(Of String)</returns>
    Public Shared Function ComboToList(c As ComboBox, Optional returnAs As ReturnInfo = ReturnInfo.AsArray)
        Dim l As New List(Of String)
        With c
            For i = 0 To .Items.Count - 1
                l.Add(c.Items(i))
            Next
        End With
        If returnAs = ReturnInfo.AsListOfString Then
            Return l
        ElseIf returnAs = ReturnInfo.AsArray Then
            Return l.ToArray
        End If
    End Function

    Public Sub ClearSearch(LocateBy_ As ComboBox, LoadThis_ As ComboBox, Optional Stat_ As TextBox = Nothing)
        Clear(LoadThis_)
        LocateBy_.Text = ""
        If Stat_ IsNot Nothing Then StatFromLoad(Stat_)
    End Sub

    Public Sub SetSearch(LocateBy_ As ComboBox, LoadThis_ As ComboBox, Optional grid_ As DataGridView = Nothing, Optional StatFromLoad_ As Boolean = False, Optional Stat As TextBox = Nothing)
        Clear(LoadThis_)
        LocateBy_.Sorted = True
        Clear(LocateBy_)
        If grid_ IsNot Nothing Then
            With grid_
                For i As Integer = 0 To .Columns.Count - 1
                    If .Columns.Item(i).Visible = True Then LocateBy_.Items.Add(.Columns.Item(i).HeaderText)
                Next
            End With
        End If
        LocateBy_.Text = ""
        If StatFromLoad_ = True Then StatFromLoad(Stat)
    End Sub

    Private Sub ClearText(c_ As Control)

        Dim c As ComboBox
        Dim t As TextBox
        Dim p As PictureBox
        Dim h As CheckBox
        If TypeOf c_ Is CheckBox Then
            h = c_
            h.Checked = False
        End If
        If TypeOf c_ Is ComboBox Then
            c = c_
            c.Text = ""
        End If
        If TypeOf c_ Is TextBox Then
            t = c_
            t.Text = ""
        End If
        If TypeOf c_ Is PictureBox Then
            p = c_
            Try
                p.Image = Nothing
            Catch ex As Exception
            End Try
            Try
                p.BackgroundImage = Nothing
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub ClearFields(dialog As Control, Optional clearData As Boolean = False)
        If TypeOf (dialog) IsNot Form And TypeOf (dialog) IsNot Panel Then Exit Sub

        For Each c As Control In dialog.Controls
            If clearData = True Then
                Clear(c)
            Else
                ClearText(c)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Commits record to SQL Server database by default, or to MS Access database if DB_Is_SQL_ is set to false.
    ''' </summary>
    ''' <param name="query">The SQL query.</param>
    ''' <param name="connection_string">The server connection string.</param>
    ''' <param name="parameters_keys_values_">Values to put in table.</param>
    ''' <returns>True if successful, False if not.</returns>
    'Private Function CommitRecord(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional DB_Is_SQL_ As Boolean = True) As Boolean
    '    Dim select_parameter_keys_values() = {}
    '    select_parameter_keys_values = parameters_keys_values_

    '    If DB_Is_SQL_ = True Then
    '        CommitSQLRecord(query, connection_string, select_parameter_keys_values)
    '        Return True
    '        Exit Function
    '    End If

    '    Try
    '        Dim insert_query As String = query
    '        Using insert_conn As New OleDbConnection(connection_string)
    '            Using insert_comm As New OleDbCommand()
    '                With insert_comm
    '                    .Connection = insert_conn
    '                    .CommandText = insert_query
    '                    If select_parameter_keys_values IsNot Nothing Then
    '                        For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
    '                            .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
    '                        Next
    '                    End If
    '                End With
    '                Try
    '                    insert_conn.Open()
    '                    insert_comm.ExecuteNonQuery()
    '                Catch ex As Exception
    '                End Try
    '            End Using
    '        End Using
    '        Return True
    '    Catch ex As Exception
    '    End Try

    '    '		Dim Insert_String As String = "INSERT INTO [iNUsers] (LastName, FirstName, MiddleName, Gender, Title, BusinessEmail, Department, JobTitle, Phone, PrimaryEmail, FacebookID, WhatsAppID, Birthday, Nationality, StateOfOrigin, LGA, ContactAddress, Religion, BloodType, UserGroup, Allergies, Picture, Username, UserPassword, NetSub, iNInfo, NetHelpline, IsEnabled, DateAdded, DateAddedD, SessionID, NetManager, DateLastModified, DateLastModifiedD) VALUES (@LastName, @FirstName, @MiddleName, @Gender, @Title, @BusinessEmail, @Department, @JobTitle, @Phone, @PrimaryEmail, @FacebookID, @WhatsAppID, @Birthday, @Nationality, @StateOfOrigin, @LGA, @ContactAddress, @Religion, @BloodType, @UserGroup, @Allergies, @Picture, @Username, @UserPassword, @NetSub, @iNInfo, @NetHelpline, @IsEnabled, @DateAdded, @DateAddedD, @SessionID, @NetManager, @DateLastModified, @DateLastModifiedD)"
    '    '		Dim parameters_() = {}
    '    '		parameters_ = {"LastName", LastName.Text.Trim, "FirstName", FirstName.Text.Trim, "MiddleName", OtherNames.Text.Trim, "Gender", Gender.Text.Trim, "Title", TitleOfCourtesy.Text.Trim, "BusinessEmail", BusinessEmail.Text.Trim, "Department", Department.Text.Trim, "JobTitle", JobTitle.Text.Trim, "Phone", Phone.Text.Trim, "PrimaryEmail", PersonalEmail.Text.Trim, "FacebookID", Facebook.Text.Trim, "WhatsAppID", WhatsApp.Text.Trim, "Birthday", Birthday.Value.ToShortDateString, "Nationality", Nationality.Text.Trim, "StateOfOrigin", StateOfOrigin.Text.Trim, "LGA", LGA.Text.Trim, "ContactAddress", ContactAddress.Text.Trim, "Religion", Religion.Text.Trim, "BloodType", KinName.Text.Trim, "UserGroup", KinEmail.Text.Trim, "Allergies", KinPhone.Text.Trim, "Picture", d.PictureFromStream(UserPicture__, PicturePath.Text), "Username", Username.Text.Trim, "UserPassword", Password.Text, "NetSub", NetSub.Checked, "iNInfo", LookOut.Checked, "NetHelpline", ChatBox.Checked, "IsEnabled", IsEnabled.Checked, "DateAdded", Now.ToShortDateString, "DateAddedD", Date.Parse(Now.ToShortDateString), "SessionID", session_id, "NetManager", SysPicture.Checked, "DateLastModified", Now.ToShortDateString, "DateLastModifiedD", Date.Parse(Now.ToShortDateString)}
    '    '		d.CommitRecord(Insert_String, con_string, parameters_)


    '    '		Dim Insert_String As String = "UPDATE [] SET LastAppLoginDateD=@LastAppLoginDateD, LastAppLoginDate=@LastAppLoginDate, LastAppLoginTimeD=@LastAppLoginTimeD, LastAppLoginTime=@LastAppLoginTime, LastAppLogin=@LastAppLogin WHERE (AccountID=@AccountID)"
    '    '		Dim parameters_() = {}
    '    '		parameters_ = {"LastAppLoginDateD", Date.Parse(Now.ToShortDateString), "LastAppLoginDate", Now.ToLongDateString, "LastAppLoginTimeD", Date.Parse(Now.ToLongTimeString), "LastAppLoginTime", Now.ToLongTimeString, "LastAppLogin", g.FullDateAsLong, "AccountID", a.id}
    '    '		d.CommitRecord(Insert_String, con_string_, parameters_)


    'End Function


    Private Function CommitSQLRecords(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As Boolean

        Try
            Dim insert_query As String = query
            Using insert_conn As New SqlConnection(connection_string)
                Using insert_comm As New SqlCommand()
                    With insert_comm
                        .Connection = insert_conn
                        .CommandTimeout = 0
                        .CommandType = CommandType.Text
                        .CommandText = insert_query
                        If select_parameter_keys_values IsNot Nothing Then
                            For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                                .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                            Next
                        End If
                    End With
                    Try
                        insert_conn.Open()
                        insert_comm.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End Using
            End Using
            Return True
        Catch ex As Exception
        End Try
    End Function
    Public Shared Function ControlImageIsNull(control_ As PictureBox, Optional UseImage As Boolean = False) As Boolean
        If UseImage = True Then
            Return control_.Image Is Nothing
        Else
            Return control_.BackgroundImage Is Nothing
        End If
    End Function

    Public Shared Function ControlTextIsNull(control_ As Control, Optional UseTrim As Boolean = True) As Boolean
        If UseTrim = True Then
            Return control_.Text.Trim.Length < 1
        Else
            Return control_.Text.Length < 1
        End If
    End Function
    Public Shared Sub PictureToPC(picture_ As Object, Optional UseImage As Boolean = False)
        Dim photo_ As Image
        If TypeOf picture_ Is PictureBox Then
            Select Case UseImage
                Case True
                    photo_ = picture_.Image
                Case False
                    photo_ = picture_.BackgroundImage
            End Select
        ElseIf TypeOf picture_ Is Image Then
            photo_ = picture_
        ElseIf TypeOf picture_ Is String Then
            photo_ = Image.FromFile(picture_)
        End If

        Try
            Clipboard.SetImage(photo_)
        Catch ex As Exception

        End Try

    End Sub
    ''' <summary>
    ''' Gets the content of a picture.
    ''' </summary>
    ''' <param name="src">PictureBox, Image, Bitmap or String representing file path</param>
    ''' <param name="file_extension"></param>
    ''' <param name="PropertyIsImage">If src is PictureBox, setting this to True grabs the Image property, not the BackgroundImage property</param>
    ''' <returns></returns>
    Private Shared Function PictureFromStream(src As Object, Optional file_extension As String = ".jpg", Optional PropertyIsImage As Boolean = False)
        Dim photo_
        Dim stream_ As New IO.MemoryStream

        If TypeOf src Is PictureBox Then
            Select Case PropertyIsImage
                Case True
                    photo_ = src.Image
                Case False
                    photo_ = src.BackgroundImage
            End Select
        ElseIf TypeOf src Is Image Or TypeOf src Is Bitmap Then
            photo_ = src
        Else
            photo_ = Image.FromFile(src)
        End If


        If photo_ IsNot Nothing Then
            Select Case LCase(file_extension)
                Case ".png"
                    photo_.Save(stream_, Imaging.ImageFormat.Png)
                Case ".gif"
                    photo_.Save(stream_, Imaging.ImageFormat.Gif)
                Case ".bmp"
                    photo_.Save(stream_, Imaging.ImageFormat.Bmp)
                Case ".tif"
                    photo_.Save(stream_, Imaging.ImageFormat.Tiff)
                Case ".ico"
                    photo_.Save(stream_, Imaging.ImageFormat.Icon)
                Case Else
                    photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
            End Select
        End If
        Return stream_.GetBuffer
    End Function
    ''' <summary>
    ''' Gets the content of a picture.
    ''' </summary>
    ''' <param name="src">PictureBox, Image, Bitmap or String representing file path</param>
    ''' <param name="file_extension"></param>
    ''' <param name="PropertyIsImage">If src is PictureBox, setting this to True grabs the Image property, not the BackgroundImage property</param>
    ''' <returns></returns>

    Private Shared Function PictureFromStream(src As Object, Optional file_extension As FileType = FileType.jpg, Optional PropertyIsImage As Boolean = False)
        Return PictureFromStream(src, "." & file_extension.ToString, PropertyIsImage)
    End Function


    Private Shared Function PictureFromStream_REVERT_TO_THIS(picture_ As PictureBox, Optional file_extension As String = ".jpg", Optional UseImage As Boolean = False) As Byte()
        Dim photo_ As Image
        Dim stream_ As New IO.MemoryStream
        Select Case UseImage
            Case True
                photo_ = picture_.Image
            Case False
                photo_ = picture_.BackgroundImage
        End Select

        If photo_ IsNot Nothing Then
            Select Case LCase(file_extension)
                Case ".jpg"
                    photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
                Case ".jpeg"
                    photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
                Case ".png"
                    photo_.Save(stream_, Imaging.ImageFormat.Png)
                Case ".gif"
                    photo_.Save(stream_, Imaging.ImageFormat.Gif)
                Case ".bmp"
                    photo_.Save(stream_, Imaging.ImageFormat.Bmp)
                Case ".tif"
                    photo_.Save(stream_, Imaging.ImageFormat.Tiff)
                Case ".ico"
                    photo_.Save(stream_, Imaging.ImageFormat.Icon)
                Case Else
                    photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
            End Select
        End If
        Return stream_.GetBuffer
    End Function

    ''' <summary>
    ''' Prompts the user to select application and argument, then creates batch file and writes it to file optionally.
    ''' </summary>
    ''' <param name="filename"></param>
    ''' <param name="append"></param>
    ''' <returns></returns>
    Public Shared Function BatchFileFromScratch(Optional filename As String = Nothing, Optional no_quotes As Boolean = False, Optional append As Boolean = True) As String
        Dim application = GetFile(FileDialogExpectedFormat.any)
        If application.Length < 1 Then Return ""
        Dim details As New List(Of String)
        details.Add(application)
        Dim argument = GetFile(FileDialogExpectedFormat.any, "Select Argument")
        details.Add(argument)
        Return BatchFile(details(0), details(1), no_quotes, filename, append)
    End Function

    ''' <summary>
    ''' Creates batch file string and writes it to file optionally.
    ''' </summary>
    ''' <param name="application"></param>
    ''' <param name="argument"></param>
    ''' <param name="filename"></param>
    ''' <param name="append"></param>
    ''' <returns></returns>

    Public Shared Function BatchFile(application As String, Optional argument As String = Nothing, Optional no_quotes As Boolean = False, Optional filename As String = Nothing, Optional append As Boolean = True) As String
        If CType(application, String).Trim.Length < 1 Then
            Return ""
            Exit Function
        End If

        Dim r As String = ""

        If no_quotes = False Then
            r = """" & application & """" & vbCr
            If argument IsNot Nothing Then
                If CType(argument, String).Trim.Length > 1 Then r = """" & application & """" & " " & """" & argument & """" & vbCr
            End If
        Else
            r = application & vbCr
            If argument IsNot Nothing Then
                If CType(argument, String).Trim.Length > 1 Then r = application & " " & argument & vbCr
            End If
        End If
        If filename IsNot Nothing Then
            If filename.Length > 0 Then
                My.Computer.FileSystem.WriteAllText(filename, r, append)
            End If
        End If
        Return r
    End Function
    'Public Shared Function BatchFile(application As String, Optional argument As String = Nothing, Optional filename As String = Nothing, Optional append As Boolean = True) As String
    '    If CType(application, String).Trim.Length < 1 Then
    '        Return ""
    '        Exit Function
    '    End If

    '    Dim r As String = """" & application & """" & vbCr
    '    If argument IsNot Nothing Then
    '        If CType(argument, String).Trim.Length > 1 Then r = """" & application & """" & " " & """" & argument & """" & vbCr
    '    End If
    '    If filename IsNot Nothing Then
    '        If filename.Length > 0 Then
    '            My.Computer.FileSystem.WriteAllText(filename, r, append)
    '        End If
    '    End If
    '    Return r
    'End Function


    ''' <summary>
    ''' Search.
    ''' </summary>
    ''' <param name="control_">Textbox to search from.</param>
    ''' <param name="sought_">What to select if found in control_'s content.</param>
    ''' <param name="current_position">Variable to keep track of how many of sought_ has been found (should be overwritten at each next call otherwise it'll only always check once).</param>
    ''' <returns>0 if sought_ isn't found, else next index where it is found.</returns>
    ''' <example>
    ''' 'DECLARATION
    ''' Public search_position As Integer = 1
    ''' Public search_ As String = ""
    ''' 'Ctrl+F
    ''' search_ = GetText("Locate ...")
    ''' search_position = LocateText(tContent, search_, search_position)
    ''' 'F3
    ''' If search_.Length > 0 Then search_position = LocateText(tContent, search_, search_position)
    ''' </example>
    Public Shared Function LocateText(control_ As TextBox, sought_ As String, Optional current_position As Integer = Nothing) As Integer
        If control_.Text.Length < 1 Then Return 0
        Dim search_ As Integer = 1
        If current_position <> Nothing Then search_ = Val(current_position)
        If search_ = 0 Then search_ = 1

        For i% = search_ To control_.Text.Length
            If Mid(control_.Text, i, sought_.Length).ToLower() = sought_.Trim.ToLower() Then
                control_.SelectionStart = i - 1
                control_.SelectionLength = sought_.Length
                control_.ScrollToCaret()
                control_.Focus()
                search_ = i + 1
                Exit For
            ElseIf Mid(control_.Text, i, sought_.Length).ToLower() <> sought_.Trim.ToLower() And control_.Text.Length - i < sought_.Length Then
                search_ = 0
                Exit For
            End If
        Next
        Return search_

        'DECLARATION
        'Public search_position As Integer = 1
        'Public search_ As String = ""
        'Ctrl+F
        'search_ = GetText("Locate ...")
        'search_position = LocateText(tContent, search_, search_position)
        'F3
        'If search_.Length > 0 Then search_position = LocateText(tContent, search_, search_position)


        'INITIALLY
        'u.text = 0
        'RECURRENT, IF u.text > 0
        'u.Text = LocateText(t, "html", u.Text)

    End Function

    'Private Shared Function CommitFile(txt_ As String, Optional default_ext As String = ".txt", Optional add_extension As Boolean = True, Optional filter_ As String = "Document|*.txt", Optional title_ As String = "Save file...") As String
    '    If txt_.Length < 1 Then Return ""
    '    Dim return_ As String = ""
    '    Dim filter__ As String = filter_docs
    '    If filter_.Length > 0 Then filter__ = filter_

    '    Dim default_ext__ As String = ".txt"
    '    If default_ext.Length > 0 Then default_ext__ = default_ext

    '    Dim title__ As String = "Save file..."
    '    If title_.Length > 0 Then title__ = title_

    '    Dim save_file As New SaveFileDialog
    '    With save_file
    '        .Filter = filter__
    '        .DefaultExt = default_ext__
    '        .AddExtension = add_extension
    '        .Title = title__
    '        .ShowDialog()

    '        If .FileName.Length > 0 Then
    '            WriteText(.FileName, txt_)
    '            return_ = .FileName
    '        End If
    '    End With
    '    Return return_
    'End Function

    Public Function SetTextToClipboard(str_ As String) As Boolean
        Try
            Clipboard.SetText(str_)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function GetTextFromClipboard(str_ As String) As String
        Try
            If Clipboard.ContainsText Then
                Return Clipboard.GetText
            End If
        Catch ex As Exception
        End Try
    End Function
    Public Shared Sub PermitFile(file_ As String, Optional temp_file As String = "", Optional user_ As String = "")
        Dim list As New List(Of String)
        list.Add(file_)
        PermitFiles(list, user_)
    End Sub
    Public Shared Sub PermitFiles(file_ As List(Of String), Optional temp_file As String = "", Optional user_ As String = "")
        On Error Resume Next
        Dim t_file As String
        If temp_file.Length < 1 Then temp_file = My.Application.Info.DirectoryPath & "\FilePermissions.bat"
        If user_.Length < 1 Then user_ = Environment.UserName

        Dim str_ As String = ""
        For i As Integer = 0 To file_.Count - 1
            str_ &= "icacls """ & file_(i) & """ /grant " & user_ & ":(F)" & vbCrLf
        Next
        WriteText(temp_file, str_)
        StartFile(temp_file)

    End Sub

    Public Sub PermissionForFile(file_ As String, Optional perm_ As FileSystemRights = FileSystemRights.FullControl, Optional remove_existing As Boolean = False)
        If file_.Length < 1 Then Exit Sub
        If perm_.ToString.Length < 1 Then perm_ = FileSystemRights.FullControl

        Dim user_ As String = Environment.UserName

        Dim FilePath As String = file_
        Dim UserAccount As String = user_
        Dim FileInfo As IO.FileInfo = New IO.FileInfo(FilePath)
        Dim FileAcl As New FileSecurity
        FileAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, AccessControlType.Deny))
    End Sub
    ''' <summary>
    ''' Sets folder acccess. Windows related.
    ''' </summary>
    ''' <param name="folder_">Folder to set permission on</param>
    ''' <param name="perm_">Type of permission</param>
    ''' <param name="remove_existing">Should existing permissions for this user be overwritten if found?</param>
    Public Sub PermissionForFolder(folder_ As String, Optional perm_ As FileSystemRights = FileSystemRights.FullControl, Optional remove_existing As Boolean = False)
        If folder_.Length < 1 Then Exit Sub
        If perm_.ToString.Length < 1 Then perm_ = FileSystemRights.FullControl

        Dim user_ As String = Environment.UserName

        Dim FolderPath As String = folder_
        Dim UserAccount As String = user_

        Dim FolderInfo As IO.DirectoryInfo = New IO.DirectoryInfo(FolderPath)
        Dim FolderAcl As New DirectorySecurity
        FolderAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
        If remove_existing = True Then FolderAcl.SetAccessRuleProtection(True, False)
        FolderInfo.SetAccessControl(FolderAcl)
    End Sub

    ''' <summary>
    ''' Sets folder acccess. Windows related. Same as PermissionForFolder.
    ''' </summary>
    ''' <param name="folder_">Folder to set permission on</param>
    ''' <param name="perm_">Type of permission</param>
    ''' <param name="remove_existing">Should existing permissions for this user be overwritten if found?</param>

    Public Shared Sub PermitFolder(folder_ As String, Optional perm_ As FileSystemRights = FileSystemRights.FullControl, Optional remove_existing As Boolean = False)
        Try
            If folder_.Length < 1 Then Exit Sub
            If perm_.ToString.Length < 1 Then perm_ = FileSystemRights.FullControl

            Dim user_ As String = Environment.UserName

            Dim FolderPath As String = folder_
            Dim UserAccount As String = user_

            Dim FolderInfo As IO.DirectoryInfo = New IO.DirectoryInfo(FolderPath)
            Dim FolderAcl As New DirectorySecurity
            FolderAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            If remove_existing = True Then FolderAcl.SetAccessRuleProtection(True, False)
            FolderInfo.SetAccessControl(FolderAcl)
        Catch
        End Try
    End Sub

    Private Shared Sub ToMachineStartup(file_ As String, key_ As String)
        If file_.Length < 1 Or key_.Length < 1 Then Exit Sub
        Try
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", key_, file_)
        Catch x As Exception
        End Try
    End Sub
    Public Shared Sub ToMachineRunOnce(file_ As String, key_ As String)
        If file_.Length < 1 Or key_.Length < 1 Then Exit Sub
        Try
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce", key_, file_)
        Catch x As Exception
        End Try
    End Sub
    Public Shared Sub ToRunOnce(file_ As String, key_ As String)
        ToMachineRunOnce(file_, key_)
    End Sub
    Public Shared Sub ToStartup(file_ As String, key_ As String)
        ToMachineStartup(file_, key_)
    End Sub

    Public Sub RemoveFromStartup(file_ As String, key_ As String)
        If file_.Length < 1 Or key_.Length < 1 Then Exit Sub

        Using key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software")
            key.DeleteSubKey("Software\Microsoft\Windows\CurrentVersion\Run\" & key_)
        End Using

        Dim exists As Boolean = False
        Try
            If My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run\" & key_) IsNot Nothing Then
            End If
        Catch
        Finally
        End Try

    End Sub
    Public Shared Sub RegistryRemoveKey(section As RegistryKey, subKey_ As String)
        Try
            section.DeleteSubKey(subKey_)
        Catch ex As Exception
        End Try

    End Sub

    Public Sub SetAttribute(file_OR_directory As String, Optional show_ As Boolean = False, Optional remove_ As Boolean = True)
        If show_ = False And remove_ = False Then
            Try
                SetAttr(file_OR_directory, FileAttribute.Hidden)
                Exit Sub
            Catch ex As Exception
            End Try
        End If

        If show_ = True Then
            Try
                SetAttr(file_OR_directory, FileAttribute.Normal)
                Exit Sub
            Catch ex As Exception
            End Try
        Else
        End If

        If remove_ = True Then
            Try
                SetAttr(file_OR_directory, FileAttribute.Hidden + FileAttribute.System)
                Exit Sub
            Catch ex As Exception
            End Try
        Else
        End If

    End Sub

    ''' <summary>
    ''' Creates a shortcut
    ''' </summary>
    ''' <param name="where_the_shortcut_points_to">The file, e.g. the executable.</param>
    ''' <param name="where_to_put_the_shortcut">The target folder where the shortcut is going to be created.</param>
    ''' <param name="filename_for_the_shortcut">The filename of the shortcut</param>
    ''' <param name="icon_file">Optional icon file.</param>
    ''' <example>
    ''' Dim exe as String = "c:\file.pdf"
    ''' CreateShortcut(exe, Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), icon_file)
    ''' </example>
    Public Shared Sub CreateShortcut(where_the_shortcut_points_to As String, where_to_put_the_shortcut As String, filename_for_the_shortcut As String, Optional icon_file As String = Nothing)
        Dim wsh As Object = CreateObject("WScript.Shell")

        wsh = CreateObject("WScript.Shell")

        Dim MyShortcut, DesktopPath

        ' Read desktop path using WshSpecialFolders object

        DesktopPath = where_to_put_the_shortcut

        MyShortcut = wsh.CreateShortcut(DesktopPath & "\" & filename_for_the_shortcut & ".lnk")

        ' Set shortcut object properties and save it

        MyShortcut.TargetPath = wsh.ExpandEnvironmentStrings(where_the_shortcut_points_to)

        MyShortcut.WorkingDirectory = wsh.ExpandEnvironmentStrings(where_to_put_the_shortcut)

        MyShortcut.WindowStyle = 4

        If icon_file IsNot Nothing Then
            If icon_file.Length > 1 Then
                MyShortcut.IconLocation = wsh.ExpandEnvironmentStrings(icon_file)
            End If
        End If

        'Save the shortcut

        MyShortcut.Save()
    End Sub

    Public Shared Function AppIsOn(file_path As String) As Boolean
        Dim p() As Process = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(file_path.Trim))
        AppIsOn = p.Count > 0
    End Function
    Public Shared Function AppNotOn(file_path As String) As Boolean
        Dim p() As Process = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(file_path.Trim))
        AppNotOn = p.Count < 1
    End Function

    Public Shared Sub KillProcess(process_name As String)
        For Each proc As Process In Process.GetProcessesByName(Path.GetFileNameWithoutExtension(process_name))
            proc.Kill()
        Next
    End Sub
    Public Shared Sub TaskManager()
        StartFile("taskmgr")
    End Sub

    Public Shared Sub StartFiles(files_list_SI As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = True)
        StartApps(files_list_SI, style_, style__, checkFirst)
    End Sub


    ''' <summary>
    ''' Opens apps from semi-JSON string.
    ''' </summary>
    ''' <param name="files_list_SI">Semi-JSON string</param>
    ''' <param name="style_">For StartApp()</param>
    ''' <param name="style__">For StartApp()</param>
    ''' <param name="checkFirst">For StartApp()</param>
    Public Shared Sub StartApps(files_list_SI As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = True)
        '		If Exists(files_list_semiJSON) = False Then Return
        Dim files_list_semiJSON As String = files_list_SI
        'open apps if not already open
        Dim v As String = files_list_semiJSON.Trim
        If v.Length < 1 Then Return
        ''v = v.Replace("\", "\\")
        Dim s As New List(Of String)
        Try
            s.Clear()
        Catch ex As Exception
        End Try
        s = Si_StringToList(v)
        For i As Integer = 0 To s.Count - 1
            StartFile(s(i), style_, style__, checkFirst)
        Next

    End Sub

    Public Shared Sub StartAppsWithArguments(files_list_semiJSON As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = False)
        '		If Exists(files_list_semiJSON) = False Then Return

        'open apps if not already open
        Dim v As String = files_list_semiJSON.Trim
        If v.Length < 1 Then Return
        ''v = v.Replace("\", "\\")
        Dim s As New List(Of String)
        Try
            s.Clear()
        Catch ex As Exception
        End Try
        s = Si_StringToList(v)
        If s.Count Mod 2 <> 0 Then Return
        ''Dim arg_ As Array = {}

        For i As Integer = 0 To s.Count - 2 Step 2
            StartFileWithArgument(s(i), {s(i + 1)}, style_, style__, checkFirst)
        Next

        '		Dim file_2 As String = "'1':'notepad', '2':'C:\Users\Administrator\Desktop\Hub\2D\StrictD\1.txt', '3':'C:\Program Files (x86)\Windows Media Player\wmplayer.exe', '4':'C:\Users\Pediforte\Music\Music\Playlists\68.wpl'"

    End Sub

    Private Shared loa As List(Of String)
    Private Shared loa_counter As Integer = 0
    Private Shared loa_timer As Timer
    Private Shared loa_checkFirst As Boolean
    Private Shared loa_initial As String
    Private Shared loa_initial_timer As Timer
    Private Shared counter_ As Integer = 0
    '	Private Shared loa_initial_counter As Integer = 0
    Private Shared Sub Load_Apps()
        If loa_counter < loa.Count - 1 Then
            StartFileWithArgument(loa(loa_counter), loa(loa_counter + 1), loa_checkFirst)
            loa_counter += 2
        Else
            loa_timer.Enabled = False
        End If

    End Sub

    Private Shared Sub Load_Initial_Apps()
        StartApps(loa_initial)
    End Sub

    Private Shared Sub onInitialAppsLoaded()
        If counter_ = 1 Then
            loa_initial_timer.Enabled = False
            loa_timer.Enabled = True
        Else
            counter_ = 1
            Load_Initial_Apps()
        End If
    End Sub

    Public Shared Sub StartAppsWithArguments(files_list_semiJSON As String, interval_in_s As Integer, initial_files_list_semiJSON As String, Optional initial_interval_in_s As Integer = 15, Optional checkFirst As Boolean = False)
        '		If Exists(files_list_semiJSON) = False Then Return

        'open apps if not already open
        Dim v As String = files_list_semiJSON.Trim
        If v.Length < 1 Then Return
        ''v = v.Replace("\", "\\")
        Dim s As New List(Of String)
        Try
            s.Clear()
        Catch ex As Exception
        End Try
        s = Si_StringToList(v)
        If s.Count Mod 2 <> 0 Then Return

        loa_initial = initial_files_list_semiJSON
        Dim t As New Timer
        loa_initial_timer = t
        t.Interval = initial_interval_in_s * 1000

        loa_counter = 0
        loa = s
        loa_checkFirst = checkFirst
        Dim timer As New Timer
        loa_timer = timer
        timer.Interval = interval_in_s * 1000

        AddHandler t.Tick, AddressOf onInitialAppsLoaded
        AddHandler timer.Tick, AddressOf Load_Apps
        Load_Initial_Apps()

        t.Enabled = True
    End Sub
    Public Shared Sub StartAppsWithArguments(files_list_semiJSON As String, interval_in_s As Integer, Optional checkFirst As Boolean = False)
        '		If Exists(files_list_semiJSON) = False Then Return

        'open apps if not already open
        Dim v As String = files_list_semiJSON.Trim
        If v.Length < 1 Then Return
        ''v = v.Replace("\", "\\")
        Dim s As New List(Of String)
        Try
            s.Clear()
        Catch ex As Exception
        End Try
        s = Si_StringToList(v)
        If s.Count Mod 2 <> 0 Then Return

        loa_counter = 0
        loa = s
        loa_checkFirst = checkFirst
        Dim timer As New Timer
        loa_timer = timer
        timer.Interval = interval_in_s * 1000
        timer.Enabled = True
        AddHandler timer.Tick, AddressOf Load_Apps

    End Sub

    Public Shared Sub StopApps(files_list_semiJSON As String, Optional useShell As Boolean = False)
        ExitApps(files_list_semiJSON, useShell)
    End Sub
    <DllImport("User32.dll")>
    Public Shared Function ShowWindow(ByVal hwnd As IntPtr, ByVal cmd As Integer) As Boolean

    End Function

    Public Shared Function ShowApp(process_name As String, state_ As AppWindowState) As Boolean
        If process_name.Length < 1 Then Return False

        process_name = Path.GetFileNameWithoutExtension(process_name)

        Try
            Dim p() As Process = Process.GetProcessesByName(process_name)
            For Each pr As Process In p
                ShowWindow(pr.MainWindowHandle, state_)
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Shared Function ShowDesktop(Optional show_ As Boolean = False) As Boolean
        Dim show_or_hide As Integer = 6
        If show_ = True Then show_or_hide = 9

        Try
            For Each p As Process In Process.GetProcesses
                ShowWindow(p.MainWindowHandle, show_or_hide)
            Next
            Return True
        Catch
            Return False
        End Try
    End Function


    Private Shared Sub ExitApps(files_list_semiJSON As String, Optional useShell As Boolean = False)
        '		If Exists(files_list_semiJSON) = False Then Return

        Dim v As String = files_list_semiJSON.Trim
        If v.Length < 1 Then Return
        ''v = v.Replace("\", "\\")
        Dim s As New List(Of String)
        Try
            s.Clear()
        Catch ex As Exception
        End Try
        s = Si_StringToList(v)
        For i As Integer = 0 To s.Count - 1
            If useShell = True Then
                'taskkill /f /im "explorer.exe"
                Shell("taskkill /f /im """ & Path.GetFileName(s(i)) & """")
            Else
                ExitApp(Path.GetFileNameWithoutExtension(s(i)))
            End If
        Next

    End Sub


    ''' <summary>
    ''' Depracated. Use StartFile() or StartApps instead.
    ''' </summary>
    ''' <param name="file_">The path to the file.</param>
    ''' <param name="style_">ProcessWindowStyle (normal, hidden etc)</param>
    ''' <param name="style__">Ignore ProcessWindowStyle, completely hide it.</param>
    ''' <param name="checkFirst">Check if an instance of the file is already running (in Tasks).</param>
    ''' <returns>True if file exists and is opened, false if not.</returns>

    Private Function StartApp(file_ As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = True) As Boolean
        If checkFirst = True Then
            Try
                If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
                    Exit Function
                End If
            Catch
            End Try
        End If

        Dim s As New ProcessStartInfo
        If style__ = True Then
            s.WindowStyle = ProcessWindowStyle.Hidden
        Else
            s.WindowStyle = style_
        End If
        s.FileName = file_
        Try
            Process.Start(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

2:
    End Function

    ''' <summary>
    ''' Opens a file. If checkFirst is True, then the program file won't run if it is already running; this is ignored if it's not a program file.
    ''' </summary>
    ''' <param name="file_">The path to the file.</param>
    ''' <param name="style_">ProcessWindowStyle (normal, hidden etc)</param>
    ''' <param name="style__">Ignore ProcessWindowStyle, completely hide it.</param>
    ''' <param name="checkFirst">Check if an instance of the file is already running (in Tasks).</param>
    ''' <returns>True if file exists and is opened, false if not.</returns>
    Public Shared Function StartFile(file_ As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = False) As Boolean
        If checkFirst = True Then
            Try
                If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
                    Exit Function
                End If
            Catch
            End Try
        End If

        Dim s As New ProcessStartInfo
        If style__ = True Then
            s.WindowStyle = ProcessWindowStyle.Hidden
        Else
            s.WindowStyle = style_
        End If
        s.FileName = file_
        Try
            Process.Start(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

2:
    End Function
    ''' <summary>
    ''' Opens a file with argument. If checkFirst is True, then the program file won't run if it is already running; this is ignored if it's not a program file. Same as StartFileWithArgument.
    ''' </summary>
    ''' <param name="file_"></param>
    ''' <param name="arg_"></param>
    ''' <param name="checkFirst"></param>
    ''' <returns></returns>
    Public Shared Function StartFile(file_ As String, arg_ As String, Optional checkFirst As Boolean = False) As Boolean
        If checkFirst = True Then
            Try
                If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
                    Exit Function
                End If
            Catch
            End Try
        End If

        Try
            Process.Start(file_, arg_)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Sub StartFileThenExit(fileToStart As String, fileToStop As String, Optional startFirst As Boolean = True)
        If fileToStart.Trim.Length < 1 Then Return
        If startFirst Then
            StartFile(fileToStart)
            If fileToStop.Trim.Length > 0 Then ExitApp(fileToStop)
        Else
            If fileToStop.Trim.Length > 0 Then ExitApp(fileToStop)
            StartFile(fileToStart)
        End If
    End Sub
    Public Shared Sub StartFileThenExit(fileToStart As String)
        If fileToStart.Trim.Length < 1 Then Return
        StartFile(fileToStart)
        ExitApp()
    End Sub

    Public Shared Function StartFileWithArgument(file_ As String, Optional arg_ As String = Nothing, Optional checkFirst As Boolean = False) As Boolean
        If checkFirst = True Then
            Try
                If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
                    Exit Function
                End If
            Catch
            End Try
        End If

        Try
            Process.Start(file_, arg_)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function StartFileWithArgument(file_ As String, list_of_arguments As Array, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = False) As Boolean
        Dim arg_ As Array = list_of_arguments, arg__ As String
        If checkFirst = True Then
            Try
                If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
                    Exit Function
                End If
            Catch
            End Try
        End If

        Dim s As New ProcessStartInfo
        If style__ = True Then
            s.WindowStyle = ProcessWindowStyle.Hidden
        Else
            s.WindowStyle = style_
        End If
        s.FileName = file_
        arg__ = array_to_string(arg_)
        s.Arguments = arg__
        Try
            Process.Start(s)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Shared Function array_to_string(array As Array) As String
        Dim s = ""
        Try
            For i = 0 To array.Length - 1
                s &= array(i) & " "
            Next
        Catch
        End Try
        Return s.Trim
    End Function

    Public Shared Function SaveFile(filter As FileKind, Optional title As String = "Choose where to save your file.", Optional filename As String = "") As String
        Dim s As New SaveFileDialog
        With s
            If filename.Length > 0 Then .FileName = filename
            .Filter = FilterStringFromFileKind(filter)
            .OverwritePrompt = True
            .Title = title
            .ShowDialog()
            If .FileName.ToString.Length > 0 Then
                Return .FileName
            Else
                Return ""
            End If
        End With
    End Function

    Public Shared Sub SaveFile(content As String, filter As FileKind, Optional title As String = "Choose where to save your file.", Optional filename As String = "")
        Dim s As String = SaveFile(filter, title, filename)
        If s.Length > 0 Then WriteText(s, content, False, True)
    End Sub

    Private Shared Function SaveFile(_filename, _filter) As Array
        Dim f_ As New SaveFileDialog
        Dim return_() As String

        With f_
            .FileName = _filename
            .Filter = _filter
            .ShowDialog()

            If .FileName.Trim <> "" Then
                return_ = {True, .FileName}
            ElseIf .FileName.Trim = "" Then
                return_ = {False, .FileName}
            End If
        End With
        Return return_
    End Function

    Public Shared Function DiskLocation(default_string As String, Optional description_ As String = "", Optional root_folder As String = "", Optional show_new_folder_button As Boolean = True) As Array
        Dim f_ As New FolderBrowserDialog
        Dim return_() As String

        With f_
            .Description = description_
            If root_folder.Length > 0 Then .RootFolder = root_folder
            .ShowNewFolderButton = show_new_folder_button

            .ShowDialog()

            If .SelectedPath.Trim <> "" Then
                return_ = {True, .SelectedPath}
            ElseIf .SelectedPath.Trim = "" Then
                return_ = {False, default_string}
            End If
        End With
        Return return_
    End Function

    Public Shared Sub Hibernate()
        Try
            Application.SetSuspendState(PowerState.Hibernate, True, True)
        Catch
        End Try
    End Sub

    Public Shared Sub LogOff()
        Try
            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-l")
        Catch
        End Try
    End Sub
    Public Shared Sub Restart()
        Try
            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-r -f -t 0")
        Catch
        End Try
    End Sub
    Public Shared Sub Shutdown()
        Try
            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-s -f -t 0")
        Catch
        End Try
    End Sub

    Public Shared Sub Sleep()
        Try
            Application.SetSuspendState(PowerState.Suspend, True, False)
        Catch ex As Exception

        End Try
    End Sub
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Sub LockWorkStation()

    End Sub
    Public Shared Sub LockPC()
        LockWorkStation()
    End Sub

    Public Shared Sub ExitApp(Optional process_name As String = "")
        'taskkill /f /im "explorer.exe"

        Try
            Select Case process_name.Length
                Case < 1
                    Environment.Exit(0)
                Case Else
                    If AppIsOn(process_name) Then KillProcess(process_name)
            End Select
        Catch
        End Try
    End Sub

    Public Shared Sub CopyFiles(src As String, tgt As String, Optional dialogs As FileIO.UIOption = FileIO.UIOption.AllDialogs)
        If src.Trim.Length < 1 Or tgt.Trim.Length < 1 Then Return
        My.Computer.FileSystem.CopyDirectory(src, tgt, dialogs, FileIO.UICancelOption.DoNothing)
    End Sub

    Public Shared Shadows Function RenameFile(src_file_path, tgt_file_name_plus_extension) As Boolean
        Try
            My.Computer.FileSystem.RenameFile(src_file_path, tgt_file_name_plus_extension)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Private Shared Property DoEvent_PlayAudio_file As String
    Private Shared Property DoEvent_PlayAudio_mode As AudioPlayMode
    Private Shared Sub DoEvent_PlayAudio()
        My.Computer.Audio.Play(DoEvent_PlayAudio_file, DoEvent_PlayAudio_mode)
    End Sub
    Private Shared thread '' As New System.Threading.Thread(AddressOf DoEvent_PlayAudio)

    Public Shared Sub PlayAudioFile(file_ As String, Optional mode_ As AudioPlayMode = AudioPlayMode.WaitToComplete, Optional threaded_ As Boolean = True)
        If file_ Is Nothing Then Exit Sub
        If file_.Length < 1 Then Exit Sub
        Try
            If threaded_ = True Then
                DoEvent_PlayAudio_file = file_
                DoEvent_PlayAudio_mode = mode_
                thread = New System.Threading.Thread(AddressOf DoEvent_PlayAudio)
                thread.Start()
            Else
                My.Computer.Audio.Play(file_, mode_)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Shared Sub StopPlayingAudioFile()
        Try
            My.Computer.Audio.Stop()
        Catch ex As Exception
        End Try
    End Sub

    Public Shared Function Exists(file_ As String) As Boolean
        If file_.Length < 1 Then
            Return False
            Exit Function
        End If
        Exists = My.Computer.FileSystem.FileExists(file_) Or My.Computer.FileSystem.DirectoryExists(file_)
    End Function


#End Region

#Region "Si_JSON"
    ''' <summary>
    ''' Converts the items of ListBox to JSON. You can save the output directly to the database. Not intended to be used as dictionary. Opposite of DatabaseToControlJSON.
    ''' </summary>
    ''' <param name="l">ListBox with items</param>
    ''' <returns>String to save to database</returns>
    Public Shared Function Si_ControlToDatabaseJSON(l As ListBox) As String
        Dim val_ As String = ""
        Dim ltemp As ListBox = l
        If ltemp.Items.Count > 0 Then
            With ltemp
                .SelectedIndex = 0
                For k As Integer = 0 To .Items.Count - 1
                    val_ &= "'" & .SelectedItem & "':'" & .SelectedItem & "'"

                    If k <> .Items.Count - 1 Then
                        val_ &= ", "
                        .SelectedIndex += 1
                    End If
                Next
            End With
        End If
        Return val_
    End Function

    ''' <summary>
    ''' Takes database Si_JSON string and populates control (ListBox, ComboBox or Textbox) with it. Opposite of ControlToDatabaseJSON. You can use NFunctions.DatabaseToListJSON() instead if you want it as a list.
    ''' </summary>
    ''' <param name="val_">Database JSON string</param>
    ''' <param name="control_">Control to bind to</param>
    ''' <param name="step_">Should the values be treated strictly as dictionary?</param>

    Public Shared Sub Si_BindListToControl(val_ As String, Optional control_ As Control = Nothing, Optional step_ As Integer = 1)
        If val_.Length < 1 Then Exit Sub

        Clear(control_)

        Dim s As New List(Of String)
        Try
            s.Clear()
        Catch ex As Exception
        End Try
        s = Si_StringToList(val_.Trim)

        Dim c As ComboBox
        Dim l_ As ListBox
        Dim t As TextBox

        If TypeOf control_ Is ComboBox Then
            c = control_
            c.DataSource = s
            Exit Sub
        ElseIf TypeOf control_ Is ListBox Then
            l_ = control_
            l_.DataSource = s
            Exit Sub
        End If

        If TypeOf control_ Is TextBox Then
            t = control_
            t.Text = ""
            For i As Integer = 0 To s.Count - 1 Step step_
                t.Text &= s(i).ToString & vbCrLf
            Next
        End If
    End Sub

#End Region

#Region "Nationality"
    Public Sub CountriesDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = True)
        If d_.Items.Count > 0 Then Exit Sub

        'Dim web_methods_ As New Web_Module.Methods
        '		Dim d_w As New Web_Module.DataConnectionWeb

        Clear(d_)

        If FirstIsEmpty Then d_.Items.Add("")

        Dim l As List(Of String) = CountriesList()
        For i As Integer = 0 To l.Count - 1
            d_.Items.Add(l(i).ToString)
        Next
    End Sub

    Public Sub StatesOfNigeriaDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = True)
        If d_.Items.Count > 0 Then Exit Sub

        'Dim web_methods_ As New Web_Module.Methods
        '		Dim d_w As New Web_Module.DataConnectionWeb

        Clear(d_)

        If FirstIsEmpty Then d_.Items.Add("")

        Dim l As List(Of String) = StatesOfNigeriaList
        For i As Integer = 0 To l.Count - 1
            d_.Items.Add(l(i).ToString)
        Next
    End Sub
    Public Sub LGAsDrop(con_string__ As String, d_ As ComboBox, Optional state_ As String = Nothing, Optional country_ As String = "Nigeria")
        Dim q As String
        If state_ IsNot Nothing Then
            BindProperty(d_, PropertyToBind.Items, BuildSelectString_DISTINCT("MyAdminDropdownState", {"LGAsOfNigeria"}, {"Countries", "StatesOfNigeria"}), con_string__, {"Countries", country_, "StatesOfNigeria", state_}, "LGAsOfNigeria")
            Exit Sub
        End If

        BindProperty(d_, PropertyToBind.Items, BuildSelectString_DISTINCT("MyAdminDropdownState", {"LGAsOfNigeria"}, Nothing), con_string__, Nothing, "LGAsOfNigeria")
    End Sub

#End Region

#Region "Bindings"
    ''' <summary>
    ''' Displays data on DataGridView.
    ''' </summary>
    ''' <param name="g_">DataGridView to bind to</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <example>Display(DataGridView, SQL_Query, Connection_String, Select_Parameters)</example>
    ''' <returns>g_</returns>

    Public Shared Function Display(g_ As DataGridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As DataGridView
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            g_.DataSource = Nothing
        Catch ex As Exception
        End Try

        Try

            Dim connection As New SqlConnection(connection_string)
            Dim sql As String = query

            Dim Command = New SqlCommand(sql, connection)

            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If

            Dim da As New SqlDataAdapter(Command)
            Dim dt As New DataTable
            da.Fill(dt)

            g_.DataSource = dt
        Catch
        End Try

        Return g_


        '		d.GData(gPayment, Payment_, g_con)

        '		Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
        '		d.GData(gPayment, School_, m_con, select_parameter_keys_values)

    End Function

    ''' <summary>
    ''' Binds ComboBox to database column.
    ''' </summary>
    ''' <param name="d_">ComboBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Column</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <param name="First_Element_Is_Empty">Should first element of ComboBox appear empty?</param>
    ''' <returns></returns>
    Private Shared Function DData(d_ As ComboBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional First_Element_Is_Empty As Boolean = True) As ComboBox

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            d_.DataSource = Nothing
        Catch ex As Exception

        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        d_.DataSource = dt
        d_.DisplayMember = data_text_field

        If First_Element_Is_Empty Then d_.SelectedIndex = -1
        Return d_
    End Function

    ''' <summary>
    ''' Binds ComboBox Text property to database field.
    ''' </summary>
    ''' <param name="d_">ComboBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>d_</returns>
    Private Shared Function DText(d_ As ComboBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As ComboBox
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            d_.DataSource = Nothing
        Catch ex As Exception

        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        d_.DataBindings.Add(b)

        Return d_
    End Function

    ''' <summary>
    ''' Binds CheckBox Checked property to database field.
    ''' </summary>
    ''' <param name="h_">CheckBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>h_</returns>
    Private Shared Function HData(h_ As CheckBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As CheckBox

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Checked", dt, data_text_field)
        h_.DataBindings.Add(b)


        Return h_
    End Function

    ''' <summary>
    ''' Binds CheckBox Text property to database field.
    ''' </summary>
    ''' <param name="h_">CheckBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>h_</returns>
    Private Shared Function HText(h_ As CheckBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As CheckBox

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        h_.DataBindings.Add(b)

        Return h_

    End Function

    ''' <summary>
    ''' Binds ListBox to database column.
    ''' </summary>
    ''' <param name="l_">ListBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Column</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>l_</returns>
    Private Shared Function LData(l_ As ListBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As ListBox
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            l_.DataSource = Nothing
        Catch ex As Exception

        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        l_.DataSource = dt
        l_.DisplayMember = data_text_field

        Return l_
    End Function

    ''' <summary>
    ''' Binds PictureBox Image property to database field.
    ''' </summary>
    ''' <param name="p_">PictureBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>p_</returns>
    Private Shared Function PImage(p_ As PictureBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As PictureBox
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            p_.Image = Nothing
        Catch ex As Exception
        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Image", dt, data_text_field, True)
        p_.DataBindings.Add(b)

        Return p_
    End Function

    ''' <summary>
    ''' Binds PictureBox BackgroundImage property to database field.
    ''' </summary>
    ''' <param name="p_">PictureBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>p_</returns>
    Private Shared Function PBackgroundImage(p_ As PictureBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As PictureBox
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            p_.BackgroundImage = Nothing
        Catch ex As Exception
        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("BackgroundImage", dt, data_text_field, True)
        p_.DataBindings.Add(b)

        Return p_
    End Function

    ''' <summary>
    ''' Binds Button Text property to database field.
    ''' </summary>
    ''' <param name="b_">Button</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>b_</returns>
    Private Shared Function BData(b_ As Button, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As Button
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        b_.DataBindings.Add(b)

        Return b_
    End Function

    ''' <summary>
    ''' Binds DateTimePicker Value property to database field.
    ''' </summary>
    ''' <param name="date_">DateTimePicker</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>date_</returns>
    Private Shared Function DATEData(date_ As DateTimePicker, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As DateTimePicker
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Value", dt, data_text_field)
        date_.DataBindings.Add(b)

        Return date_
    End Function

    ''' <summary>
    ''' Binds Label Text property to database field.
    ''' </summary>
    ''' <param name="label_">Label</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>label_</returns>
    Private Shared Function LABELData(label_ As Label, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As Label

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        label_.DataBindings.Add(b)


        Return label_

    End Function

    ''' <summary>
    ''' Binds TextBox Text property to database field.
    ''' </summary>
    ''' <param name="t_">TextBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>t_</returns>
    Private Shared Function TData(t_ As TextBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As TextBox

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        t_.DataBindings.Add(b)


        Return t_

    End Function

    ''' <summary>
    ''' Attaches List as DataSource to ComboBox or ListBox.
    ''' </summary>
    ''' <param name="TheControl">ComboBox or ListBox</param>
    ''' <param name="TheList">List</param>
    ''' <returns></returns>
    'Public Shared Function BindItems(TheControl As Control, TheList As List(Of String)) As Control
    '    If TypeOf TheControl Is ComboBox Then
    '        CType(TheControl, ComboBox).DataSource = TheList
    '    ElseIf TypeOf TheControl Is ListBox Then
    '        CType(TheControl, ListBox).DataSource = TheList
    '    End If
    '    Return TheControl
    'End Function

    ''' <summary>
    ''' Attaches List as DataSource to ComboBox or ListBox.
    ''' </summary>
    ''' <param name="TheControl">ComboBox or ListBox</param>
    ''' <param name="TheList">List</param>
    ''' <returns></returns>
    Public Shared Function BindItems(TheControl As Control, TheList As List(Of String), Optional AsDataSource As Boolean = True) As Control
        If AsDataSource Then
            If TypeOf TheControl Is ComboBox Then
                CType(TheControl, ComboBox).DataSource = TheList
            ElseIf TypeOf TheControl Is ListBox Then
                CType(TheControl, ListBox).DataSource = TheList
            End If
        Else
            If TypeOf TheControl Is ComboBox Then
                CType(TheControl, ComboBox).DataSource = Nothing
                CType(TheControl, ComboBox).Items.Clear()
                For i = 0 To TheList.Count - 1
                    CType(TheControl, ComboBox).Items.Add(TheList(i))
                Next
            ElseIf TypeOf TheControl Is ListBox Then
                CType(TheControl, ListBox).DataSource = Nothing
                CType(TheControl, ListBox).Items.Clear()
                For i = 0 To TheList.Count - 1
                    CType(TheControl, ListBox).Items.Add(TheList(i))
                Next
            End If

        End If
        Return TheControl
    End Function
    ''' <summary>
    ''' Attaches List as DataSource to ComboBox or ListBox.
    ''' </summary>
    ''' <param name="TheControl">ComboBox or ListBox</param>
    ''' <param name="TheList">List</param>
    ''' <param name="InitialSelectedIndexIsNegativeOne">Should the SelectedIndex be set to -1? Only applies to ComboBox</param>
    ''' <returns></returns>
    Public Shared Function BindProperty(TheControl As Control, TheList As List(Of String), InitialSelectedIndexIsNegativeOne As Boolean) As Control
        If TypeOf TheControl Is ComboBox Then
            CType(TheControl, ComboBox).DataSource = TheList
            CType(TheControl, ComboBox).SelectedIndex = If(InitialSelectedIndexIsNegativeOne, -1, 0)
        ElseIf TypeOf TheControl Is ListBox Then
            CType(TheControl, ListBox).DataSource = TheList
        End If

        Return TheControl
    End Function

    ''' <summary>
    ''' Populates ComboBox or ListBox with items of List.
    ''' </summary>
    ''' <param name="TheControl">ComboBox or ListBox</param>
    ''' <param name="TheList">List</param>
    ''' <param name="InitialSelectedIndexIsNegativeOne">Should the SelectedIndex be set to -1? Only applies to ComboBox</param>
    ''' <param name="ReplaceUnderscoresWithSpace">If using enum with values having more than one word separted by underscore, this replaces the underscores with space</param>
    ''' <returns>TheControl</returns>
    Public Shared Function BindProperty(TheControl As Control, TheList As List(Of String), Optional InitialSelectedIndexIsNegativeOne As Boolean = True, Optional ReplaceUnderscoresWithSpace As Boolean = False) As Control
        If TheList.Count < 1 Or TheControl Is Nothing Then Return Nothing

        If TypeOf TheControl Is ComboBox Then
            Dim c As ComboBox = CType(TheControl, ComboBox)
            c.DataSource = Nothing
            c.Items.Clear()
            For i = 0 To TheList.Count - 1
                c.Items.Add(If(ReplaceUnderscoresWithSpace, TheList(i).Replace("_", " "), TheList(i)))
            Next

            c.SelectedIndex = If(InitialSelectedIndexIsNegativeOne, -1, 0)

        ElseIf TypeOf TheControl Is ListBox Then
            Dim c As ListBox = CType(TheControl, ListBox)
            c.DataSource = Nothing
            c.Items.Clear()
            For i = 0 To TheList.Count - 1
                c.Items.Add(If(ReplaceUnderscoresWithSpace, TheList(i).Replace("_", " "), TheList(i)))
            Next
        End If
        Return TheControl
    End Function

    ''' <summary>
    ''' Attaches List as DataSource to ComboBox or ListBox. Discontinued. Use other methods of same name as suitable.
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="list_"></param>
    ''' <returns></returns>
    Private Shared Function BindPropertyEx(control_ As Control, list_ As Object, Optional TextIsFirstItem As Boolean = True, Optional FirstItemIs As String = Nothing, Optional bindAsDatasourceNotList As Boolean = False) As Control
        If bindAsDatasourceNotList Then
            Try
                CType(control_, ComboBox).DataSource = list_
            Catch ex As Exception

            End Try

            Try
                CType(control_, ListBox).DataSource = list_
            Catch ex As Exception

            End Try
        Else

            Try

                CType(control_, ComboBox).Items.Clear()
                CType(control_, ComboBox).DataSource = Nothing
                If FirstItemIs IsNot Nothing Then
                    If CType(FirstItemIs, String).Trim.Length > 0 Then
                        CType(control_, ComboBox).Items.Add(FirstItemIs)
                    End If
                End If

                Dim text As String

                For i = 0 To list_.Count - 1
                    If text Is Nothing Then text = list_(i).ToString.Trim
                    CType(control_, ComboBox).Items.Add(list_(i).ToString.Trim)
                Next

                CType(control_, ComboBox).Text = If(TextIsFirstItem, text, "")

            Catch ex As Exception
            End Try
            Try
                CType(control_, ListBox).Items.Clear()
                CType(control_, ListBox).DataSource = Nothing
                If FirstItemIs IsNot Nothing Then
                    If CType(FirstItemIs, String).Trim.Length > 0 Then
                        CType(control_, ListBox).Items.Add(FirstItemIs)
                    End If
                End If

                For i = 0 To list_.Count - 1
                    CType(control_, ListBox).Items.Add(list_(i).ToString.Trim)
                Next

            Catch ex As Exception
            End Try

        End If
        Return control_

    End Function


    ''' <summary>
    ''' Binds Control property to database column/field.
    ''' </summary>
    ''' <param name="control_">Control</param>
    ''' <param name="property__">Property to bind to (Text, Value, Image, BackgroundImage or empty to bind to column)</param>
    ''' <param name="query_">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters (Nothing, if not needed)</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="First_Item_Is_Empty">Should first element of ComboBox appear empty?</param>
    ''' <returns>control_</returns>
    Public Shared Function BindProperty(control_ As Control, property__ As PropertyToBind, query_ As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional data_text_field As String = "", Optional First_Item_Is_Empty As Boolean = True) As Control
        Dim property_ As String = ""
        Select Case property__
            Case PropertyToBind.BackgroundImage
                property_ = property__.ToString.ToLower
            Case PropertyToBind.Check
                property_ = "checked"
            Case PropertyToBind.Image
                property_ = property__.ToString.ToLower
            Case PropertyToBind.Text
                property_ = "text"
        End Select

        'c, text
        'c, checked
        If TypeOf control_ Is CheckBox Then
            If property_.ToLower = "text" Then
                Return HText(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            ElseIf property_.ToLower = "checked" Then
                Return HData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            End If
        End If
        'g
        If TypeOf control_ Is DataGridView Then
            Return Display(control_, query_, connection_string, select_parameter_keys_values_)
        End If
        'd, text
        'd, data
        If TypeOf control_ Is ComboBox Then
            If property_.ToLower = "text" Then
                Return DText(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            Else
                Return DData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_, First_Item_Is_Empty)
            End If
        End If
        'l
        If TypeOf control_ Is ListBox Then
            Return LData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
        'p, image
        'p, backgroundImage
        If TypeOf control_ Is PictureBox Then
            If property_.ToLower = "image" Then
                Return PImage(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            Else
                Return PBackgroundImage(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            End If
        End If
        'b, text
        If TypeOf control_ Is Button Then
            Return BData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
        'date, value
        If TypeOf control_ Is DateTimePicker Then
            Return DATEData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
        'l, text
        If TypeOf control_ Is Label Then
            Return LABELData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
        't, text
        If TypeOf control_ Is TextBox Then
            Return TData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
    End Function

#End Region

#Region "Stat"
    Private Sub StatFromLoad(stat_ As TextBox)
        stat_.Text = "To begin, locate a record."
    End Sub

    Private Sub StatFromSearch(stat_ As System.Windows.Forms.Control, g_ As DataGridView, term_ As String, Optional phrase_ As String = "matching selected ", Optional keyword_singular As String = "record", Optional keyword_plural As String = "records")
        Dim var As String = ""
        If phrase_.Substring(0, 1) <> " " Then var = " " & phrase_ : phrase_ = var
        If phrase_.Substring(phrase_.Length - 1) <> " " Then phrase_ &= " "
        Select Case g_.Rows.Count
            Case 1
                stat_.Text = "1 " & keyword_singular & phrase_ & term_
            Case Else
                stat_.Text = g_.Rows.Count & " " & keyword_plural & phrase_ & term_
        End Select

    End Sub
#End Region


#Region "Print"
    Private Function GridInfo(grid_ As DataGridView, stat_ As String, app_ As String) As String
        Dim str$ = "Printout from " & app_ & vbCrLf & "Date: " & Now.ToShortDateString & ",  Time: " & Now.ToLongTimeString & vbCrLf & vbCrLf & vbCrLf & stat_ & vbCrLf & vbCrLf

        With grid_
            For pg_r% = 0 To .Rows.Count - 1
                str &= "#" & pg_r + 1 & vbCrLf
                For pg_c% = 0 To .Columns.Count - 1
                    If .Columns.Item(pg_c).Visible = True Then str &= .Columns.Item(pg_c).HeaderText & ":" & vbCrLf & .Rows(pg_r).Cells(pg_c).Value.ToString & vbCrLf
                Next
                str &= vbCrLf
            Next
        End With

        Return str
    End Function

    Private Sub GridToFile(grid_ As DataGridView, stat_ As String, app_ As String)
        If grid_.Rows.Count < 1 Then Exit Sub
        Dim str_$ = GridInfo(grid_, stat_, app_)
        SaveToFile(str_, app_)
    End Sub
    Private Sub SaveToFile(str As String, app As String)
        Dim file_location_ As String = My.Application.Info.DirectoryPath & "\" & app & "\Saved Records"
        Try
            My.Computer.FileSystem.CreateDirectory(file_location_)
        Catch ex As Exception

        End Try
        Dim filename As String = file_location_ & "\" & Now.Year & "_" & Now.Month & "_" & Now.Day & "_" & Now.Hour & "_" & Now.Minute & "_" & Now.Second & "_" & Now.Millisecond & ".txt"
        Try
            My.Computer.FileSystem.WriteAllText(filename, str, False)
            If MsgBox("Information successfully saved to " & filename & vbCrLf & vbCrLf & "Open it?", MsgBoxStyle.YesNo + MsgBoxStyle.Information) = MsgBoxResult.Yes Then
                Try
                    Process.Start(filename)
                Catch ex As Exception
                    MsgBox("Could not open the file. One or more errors occured.", MsgBoxStyle.Information)
                End Try
            End If
        Catch ex_ As Exception
            'ReturnFeedback("There was a problem while trying to process your request. Please veriy that the operation was successful.")
        End Try
    End Sub

#End Region

#Region "Sequel"
    Public Shared Function Display(filename As String, p As Object, Optional use_image As Boolean = False)
        Try
            If use_image Then
                CType(p, PictureBox).Image = Image.FromFile(filename)
            Else
                CType(p, PictureBox).BackgroundImage = Image.FromFile(filename)
            End If
        Catch ex As Exception

        End Try

    End Function
    Public Shared Function DisplayFromExcel(g_ As DataGridView, query As String, connection_string As String) As DataGridView
        Try
            With g_
                .DataSource = Nothing
            End With
        Catch ex As Exception

        End Try
        Dim table As DataTable = iNovation.Code.Sequel.QDataTableFromExcel(query, connection_string)
        g_.DataSource = table
        Return g_
    End Function
    'Public Shared Function DisplayFromExcel(g_ As DataGridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As DataGridView
    '    Try
    '        With g_
    '            .DataSource = Nothing
    '        End With
    '    Catch ex As Exception

    '    End Try
    '    Dim table As DataTable = iNovation.Code.Sequel.QDataTableFromExcel(query, connection_string, select_parameter_keys_values_)
    '    g_.DataSource = table
    '    Return g_
    'End Function

#End Region

#Region "Network"
    ''' <summary>
    ''' Same as IPAddresses, except it returns String by default
    ''' </summary>
    ''' <param name="machine_name"></param>
    ''' <param name="_return"></param>
    ''' <param name="line_break_delimeter"></param>
    ''' <returns></returns>
    Public Shared Function IP(Optional machine_name As String = Nothing, Optional _return As ReturnInfo = ReturnInfo.AsString, Optional line_break_delimeter As Object = ", ")
        Return IPAddresses(machine_name, _return, line_break_delimeter)
    End Function

    ''' <summary>
    ''' returns list of IP address on the network of machine_name or of the current machine_name if machine_name is not specified.
    ''' </summary>
    ''' <param name="machine_name">name of machine</param>
    ''' <param name="_return">List(Of String) or Array</param>
    ''' <returns>List(Of String) or Array</returns>
    Public Shared Function IPAddresses(Optional machine_name As String = Nothing, Optional _return As ReturnInfo = ReturnInfo.AsListOfString, Optional line_break_delimeter As Object = ", ")
        Dim server As String = Environment.MachineName
        If machine_name IsNot Nothing Then server = machine_name
        Dim return_
        Dim c As String = ""
        Dim l_string As New List(Of String)
        Dim suffix As String = ", "
        If line_break_delimeter.ToString.Length > 0 Then suffix = line_break_delimeter

        Try
            Dim ASCII As New System.Text.ASCIIEncoding()

            ' Get server related information.
            Dim heserver As IPHostEntry = Dns.Resolve(server)

            ' Loop on the AddressList
            Dim curAdd As IPAddress, counter As Integer = 0
            For Each curAdd In heserver.AddressList

                ' Display the server IP address in the standard format. In 
                ' IPv4 the format will be dotted-quad notation, in IPv6 it will be
                ' in in colon-hexadecimal notation.
                If _return = ReturnInfo.AsString Then
                    c &= curAdd.ToString()
                ElseIf _return = ReturnInfo.AsArray Or _return = ReturnInfo.AsListOfString Then
                    l_string.Add(curAdd.ToString())
                End If
                counter += 1
                If counter < heserver.AddressList.Count Then
                    c &= suffix
                End If
            Next curAdd
            If _return = ReturnInfo.AsString Then
                Return c
            ElseIf _return = ReturnInfo.AsArray Then
                Return l_string.ToArray
            ElseIf _return = ReturnInfo.AsListOfString Then
                Return l_string
            End If
        Catch e As Exception

        End Try
    End Function 'IPAddresses

#End Region

#Region "Other Functions"
    Public Shared Function DateToDotNet(date_time As Object, Optional format_ As DateFormat = DateFormat.ShortDate) As String
        Select Case format_
            Case DateFormat.LongDate
                Return Date.Parse(date_time).ToLongDateString
            Case Else
                Return Date.Parse(date_time).ToShortDateString
        End Select
    End Function

    Public Shared Function DateToSQL(date_time As Object, Optional short_ As Boolean = True)

        If date_time.ToString.Contains("-") Then Return date_time

        Dim d As Date = Date.Parse(date_time)
        Dim month = d.Month
        Dim day = d.Day
        Dim year = d.Year
        Dim hour = d.Hour
        Dim minute = d.Minute
        Dim second = d.Second
        Dim milli_second = "00" & d.Millisecond
        '2020-06-09 18:33:44.000
        Return year & "-" & LeadingZero(month) & "-" & LeadingZero(day) _
            & " 00:00:00.000"
    End Function

    'Public Shared Function DateToShort(obj_ As Object, Optional convert_date_to_short As Boolean = True) As String
    '	Try
    '		If TypeOf obj_ Is Date Or TypeOf obj_ Is DateTime Or IsDate(obj_) Then
    '			If convert_date_to_short Then
    '				Try
    '					Return Date.Parse(Date.Parse(obj_).ToShortDateString())
    '				Catch
    '				End Try
    '			Else
    '				Try
    '					Return obj_.ToString
    '				Catch
    '				End Try
    '			End If
    '		Else
    '			Try
    '				Return obj_.ToString
    '			Catch
    '			End Try
    '		End If
    '	Catch
    '	End Try
    'End Function
    'Public Shared Function DateToShort_(obj_ As Object, Optional convert_date_to_short As Boolean = True, Optional use_short_format As Boolean = True) As String
    '	If TypeOf obj_ Is Date Or TypeOf obj_ Is DateTime Or IsDate(obj_) Then
    '		If convert_date_to_short Then
    '			If use_short_format Then
    '				Return Date.Parse(Date.Parse(obj_).ToShortDateString())
    '			Else
    '				Return Date.Parse(Date.Parse(obj_).ToLongDateString())
    '			End If
    '		End If
    '	Else
    '		Return obj_.ToString
    '	End If
    'End Function

    Public Shared Function ListFiles(directory_ As String, Optional ext_ As String = "*.txt", Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchTopLevelOnly) As List(Of String)
        Dim Files__ As ReadOnlyCollection(Of String)
        Try
            Files__ = My.Computer.FileSystem.GetFiles(directory_, search_depth, ext_)
            If Files__.Count > 0 Then Return Files__.ToList()
        Catch
        End Try
    End Function

    Public Shared Function GetFiles(listbox As ListBox, directory_ As String, fileType As FileType, Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories, Optional fileNamesOnly As Boolean = True) 'As List(Of String)
        Try
            Dim ls As ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(directory_, search_depth, FileTypeToString(fileType))
            Dim lt As New List(Of String)
            With ls
                For i = 0 To .Count - 1
                    If fileNamesOnly Then
                        lt.Add(Path.GetFileNameWithoutExtension(ls(i)))
                    Else
                        lt.Add(ls(i))
                    End If
                Next
            End With
            Try
                listbox.Items.Clear()
            Catch ex As Exception

            End Try

            With lt
                For i = 0 To .Count - 1
                    listbox.Items.Add(lt(i))
                Next
            End With
            Return lt
        Catch x As System.UnauthorizedAccessException
            Return "One or more folders could not be searched because the machine disallowed it. Try narrowing down to a particular folder."
        End Try
    End Function

    Public Shared Function GetFiles(listbox As ListBox, directory_ As String, fileType As String, Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories, Optional fileNamesOnly As Boolean = True) 'As List(Of String)
        Try
            Dim ls As ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(directory_, search_depth, fileType)
            Dim lt As New List(Of String)
            With ls
                For i = 0 To .Count - 1
                    If fileNamesOnly Then
                        lt.Add(Path.GetFileNameWithoutExtension(ls(i)))
                    Else
                        lt.Add(ls(i))
                    End If
                Next
            End With
            Try
                listbox.Items.Clear()
            Catch ex As Exception

            End Try

            With lt
                For i = 0 To .Count - 1
                    listbox.Items.Add(lt(i))
                Next
            End With
            Return lt
        Catch x As System.UnauthorizedAccessException
            Return "One or more folders could not be searched because the machine disallowed it. Try narrowing down to a particular folder."
        End Try
    End Function

    Public Shared Function GetFiles(directory_ As String, fileType As FileType, Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories, Optional fileNamesOnly As Boolean = False) 'As List(Of String)
        Try
            Dim ls As ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(directory_, search_depth, FileTypeToString(fileType))
            Dim lt As New List(Of String)
            With ls
                For i = 0 To .Count - 1
                    If fileNamesOnly Then
                        lt.Add(Path.GetFileNameWithoutExtension(ls(i)))
                    Else
                        lt.Add(ls(i))
                    End If
                Next
            End With
            Return lt
        Catch x As System.UnauthorizedAccessException
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Get Files from directory
    ''' </summary>
    ''' <param name="directory_"></param>
    ''' <param name="fileType">for example, *.jpg</param>
    ''' <param name="search_depth"></param>
    ''' <param name="fileNamesOnly"></param>
    ''' <returns></returns>
    Public Shared Function GetFiles(directory_ As String, fileType As String, Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories, Optional fileNamesOnly As Boolean = False) 'As List(Of String)

        Try
            Dim ls As ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(directory_, search_depth, fileType)
            Dim lt As New List(Of String)
            With ls
                For i = 0 To .Count - 1
                    If fileNamesOnly Then
                        lt.Add(Path.GetFileNameWithoutExtension(ls(i)))
                    Else
                        lt.Add(ls(i))
                    End If
                Next
            End With
            Return lt
        Catch x As System.UnauthorizedAccessException
            Return False
        End Try
    End Function

    Private Shared Function GetFiles(listbox_ As ListBox, directory_ As String, file_type As String, Optional ClearBeforeFill As Boolean = False) As Boolean
        If ClearBeforeFill = True Then listbox_.Items.Clear()
        Dim _Files As ReadOnlyCollection(Of String)
        Try
            If My.Computer.FileSystem.DirectoryExists(directory_) = True And My.Computer.FileSystem.GetFiles(directory_, FileIO.SearchOption.SearchAllSubDirectories, file_type).Count > 1 Then
                _Files = My.Computer.FileSystem.GetFiles(directory_, FileIO.SearchOption.SearchAllSubDirectories, file_type)
                For i% = 0 To _Files.Count - 1
                    listbox_.Items.Add(_Files.Item(i))
                Next
                Return True
            Else
                Return False
            End If
        Catch
            Return False
        End Try
    End Function

    Public Shared Function GetDirectories(directory_ As String, Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories) As ReadOnlyCollection(Of String)
        Dim folders As ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetDirectories(directory_, search_depth)
        If folders.Count > 0 Then
            Return folders
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetFolders(directory_ As String, Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories) As List(Of String)
        Return GetDirectories(directory_, search_depth).ToList
    End Function


    ''' <summary>
    ''' No longer supported. Retained only for legacy purposes. Prefer FileKind.
    ''' </summary>
    Public Enum FileDialogExpectedFormat
        audio
        video
        picture
        executable
        any
        batch
        icon
        installer
        code
        text
        supported
    End Enum


    Public Shared Function GetFile(FileDialogExpectedFormat_ As FileKind, Optional title_text As String = "Select File") As String
        Dim f_ As New OpenFileDialog
        f_.Multiselect = False
        f_.Filter = FilterStringFromFileKind(FileDialogExpectedFormat_)
        f_.ShowDialog()
        If f_.FileName.Length > 0 Then
            Return f_.FileName
        Else
            Return ""
        End If
    End Function

    Public Shared Function GetFile(FileDialogExpectedFormats_ As List(Of FileKind), Optional title_text As String = "Select File") As String
        Dim f_ As New OpenFileDialog
        f_.Multiselect = False
        f_.Filter = FilterStringFromFileKind(FileDialogExpectedFormats_)
        f_.ShowDialog()
        If f_.FileName.Length > 0 Then
            Return f_.FileName
        Else
            Return ""
        End If
    End Function

    Public Shared Function GetFile(FilterString As String, Optional title_text As String = "Select File") As String
        Dim f_ As New OpenFileDialog
        f_.Multiselect = False
        f_.Filter = FilterString
        f_.ShowDialog()
        If f_.FileName.Length > 0 Then
            Return f_.FileName
        Else
            Return ""
        End If
    End Function



    ''' <summary>
    ''' No longer supported. For legacy purposes only. Prefer GetFile(FileDialogExpectedFormat_ As FileKind, Optional title_text As String = "Select File") and/or variants.  
    ''' </summary>
    ''' <param name="FileDialogExpectedFormat_"></param>
    ''' <returns></returns>
    Public Shared Function GetFile(FileDialogExpectedFormat_ As FileDialogExpectedFormat, Optional title_text As String = "Select File") As String
        Dim audio_video_picture_exec_all_combined As FileDialogExpectedFormat = FileDialogExpectedFormat_
        ''        Dim fw As New FormatWindow

        Dim f_ As New OpenFileDialog
        Select Case audio_video_picture_exec_all_combined
            Case FileDialogExpectedFormat.audio
            Case FileDialogExpectedFormat.video
            Case FileDialogExpectedFormat.picture
                f_.Filter = "Images|*.GIF;*.JPG;*.JPEG;*.TIF;*.BMP;*.PNG;*.ICO"
                f_.Title = title_text
            Case FileDialogExpectedFormat.icon
                f_.Filter = "Icons|*.ico;*.exe"
                f_.Title = title_text
            Case FileDialogExpectedFormat.executable
                f_.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
                f_.Filter = "App Launcher Files|*.exe"
                f_.Title = title_text
                f_.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
            Case FileDialogExpectedFormat.any
                f_.Filter = "Any File|*.*"
                f_.Title = title_text
            Case FileDialogExpectedFormat.batch
                f_.Filter = "Batch File|*.bat"
                f_.Title = title_text
            Case FileDialogExpectedFormat.code
                f_.Filter = "Code|*.html;*.htm;*.css;*.js;*.py;*.java;*.cs;*.vb;*.md;*.json,*.bat"
            Case FileDialogExpectedFormat.text
                f_.Filter = "Text File|*.txt"
                f_.Title = title_text
            Case FileDialogExpectedFormat.supported
                f_.Filter = "Supported Files|*.txt;*.html;*.htm;*.css;*.js;*.py;*.java;*.cs;*.vb;*.md;*.json;*.bat;*.idf"
                f_.Title = title_text


        End Select
        f_.Multiselect = False
        f_.ShowDialog()
        If f_.FileName.Length > 0 Then
            Return f_.FileName
        Else
            Return ""
        End If
    End Function


    ''' <summary>
    ''' No longer supported. For legacy purposes only.
    ''' </summary>
    Public Shared Function GetFile(textBox As Control, FileDialogExpectedFormat_ As FileDialogExpectedFormat, Optional title_text As String = "Select File") As String
        Dim s As String = GetFile(FileDialogExpectedFormat_, title_text)
        If s.Length > 0 Then
            If textBox IsNot Nothing Then
                textBox.Text = ReadText(s)
            End If
            Return s
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' Gets file name and extension from OpenFileDialog. 
    ''' </summary>
    ''' <param name="audio_video_picture_exec_all_combined">What it is expected to find.</param>
    ''' <returns>String Array {0:True if no error or False if otherwise (convert to bool), 1:File Path, 2:File Extension</returns>
    Public Shared Function GetFileAndExtension(Optional audio_video_picture_exec_all_combined As FileKind = FileKind.AnyFileKind, Optional title_text As String = "Select File") As Array
        '		If audio_video_picture_exec_all_combined.Length < 1 Then audio_video_picture_exec_all_combined = "all"
        Dim f_ As New OpenFileDialog
        Dim return_() As String
        With f_
            .Filter = FilterStringFromFileKind(audio_video_picture_exec_all_combined)
            f_.Multiselect = False
            f_.ShowDialog()
            If .FileName.Trim <> "" Then
                return_ = {True, .FileName, Path.GetExtension(.FileName)}
            ElseIf .FileName.Trim = "" Then
                return_ = {False, "", ""}
            End If
        End With
        Return return_

    End Function

    ''' <summary>
    ''' Gets file name and extension from OpenFileDialog. 
    ''' </summary>
    ''' <param name="fileKinds">What it is expected to find.</param>
    ''' <returns>String Array {0:True if no error or False if otherwise (convert to bool), 1:File Path, 2:File Extension</returns>
    Public Shared Function GetFileAndExtension(fileKinds As List(Of FileKind), Optional title_text As String = "Select File") As Array
        '		If audio_video_picture_exec_all_combined.Length < 1 Then audio_video_picture_exec_all_combined = "all"
        Dim f_ As New OpenFileDialog
        Dim return_() As String
        With f_
            .Filter = FilterStringFromFileKind(fileKinds)
            f_.Multiselect = False
            f_.ShowDialog()
            If .FileName.Trim <> "" Then
                return_ = {True, .FileName, Path.GetExtension(.FileName)}
            ElseIf .FileName.Trim = "" Then
                return_ = {False, "", ""}
            End If
        End With
        Return return_

    End Function







    ''' <summary>
    ''' Gets Folder Path from FolderBrowserDialog. Use if no language is involved.
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="description_"></param>
    ''' <param name="ShowNewFolderButton_"></param>
    ''' <returns></returns>
    Public Shared Function GetFolder(Optional description_ As String = "Select Folder", Optional ShowNewFolderButton_ As Boolean = True, Optional control_ As Control = Nothing) As String

        Dim t
        If control_ IsNot Nothing And TypeOf control_ Is TextBox Then t = control_
        If control_ IsNot Nothing And TypeOf control_ Is ComboBox Then t = control_
        If control_ IsNot Nothing And TypeOf control_ IsNot TextBox And TypeOf control_ IsNot ComboBox Then Exit Function

        Dim f_ As New FolderBrowserDialog

        With f_
            .Description = description_
            .ShowNewFolderButton = ShowNewFolderButton_
            .ShowDialog()

            If .SelectedPath.Length > 0 Then
                If t IsNot Nothing Then t.Text = .SelectedPath
                Return .SelectedPath
                'ElseIf .SelectedPath.Length < 1 Then
                '	On Error Resume Next
                '	Return t.text
            End If

        End With
    End Function

    ''' <summary>
    ''' Gets text from InputBox. Use this if no language is involved.
    ''' </summary>
    ''' <param name="prompt_">Message to display.</param>
    ''' <param name="default_response">String to return if none is given by user.</param>
    ''' <param name="title_">Title of the input box.</param>
    ''' <param name="ReturnDefaultResponseIfNoResponse">When set to true, it returns value of default_response if user types nothing in input box or presses cancel on the input box</param>
    ''' <returns></returns>
    Public Shared Function GetText(prompt_ As String, Optional default_response As String = "", Optional title_ As String = "", Optional ReturnDefaultResponseIfNoResponse As Boolean = False) As String
        Dim response As String = InputBox(prompt_, title_, default_response)
        Return If(response.Trim.Length > 0, response, If(ReturnDefaultResponseIfNoResponse, default_response, ""))
    End Function

    ''' <summary>
    ''' Gets file's content. Same as ReadText.
    ''' </summary>
    ''' <param name="file_"></param>
    ''' <returns></returns>
    Public Function file_content(file_ As String) As String
        Return ReadText(file_) : Return True : Exit Function
    End Function

#End Region

#Region "Drag-Drop"

#Region "DragDrop-PictureBox"

    ''' <summary>
    '''Declare constant for use in detecting whether the Ctrl key was pressed during the drag operation. 
    ''' </summary>
    Const CtrlMask As Byte = 8

    ''' <summary>
    ''' Initiates DragDrop, called from _MouseDown (PictureBox_DragEnter and PictureBox_DragDrop must be handled as well).
    ''' </summary>
    ''' <example>PictureBoxMouseDown(sender, e, picLeft, picRight)</example>
    Public Shared Sub PictureBoxMouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs, ByVal picLeft As PictureBox, ByVal picRight As PictureBox)
        picLeft.AllowDrop = True
        picRight.AllowDrop = True

        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim pic As PictureBox = CType(sender, PictureBox)
            ' Invoke the drag and drop operation.
            If Not pic.BackgroundImage Is Nothing Then
                pic.DoDragDrop(pic.BackgroundImage, DragDropEffects.Move Or DragDropEffects.Copy)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Handles the DragEnter event for both PictureBox controls. DragEnter is the
    ''' event that fires when an object is dragged into the control's bounds.
    ''' </summary>
    ''' <example>PictureBoxDragEnter(sender, e)</example>
    Public Shared Sub PictureBoxDragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs)
        ' Check to be sure that the drag content is the correct type for this 
        ' control. If not, reject the drop.
        If (e.Data.GetDataPresent(DataFormats.Bitmap)) Then
            ' If the Ctrl key was pressed during the drag operation then perform
            ' a Copy. If not, perform a Move.
            If (e.KeyState And CtrlMask) = CtrlMask Then
                e.Effect = DragDropEffects.Copy
            Else
                '                e.Effect = DragDropEffects.Move
                e.Effect = DragDropEffects.Copy
            End If
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    ''' <summary>
    ''' This method handles the DragDrop event for both PictureBox controls. 
    ''' </summary>
    ''' <example>If PictureBoxDragDrop(sender, e) Then MsgBox("CallBack!")</example>
    ''' <returns>True, to signal end of DragDrop, which is usable as a callback.</returns>
    Public Shared Function PictureBoxDragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs, Optional callBack As Boolean = False) As Boolean
        ' Display the image in the selected PictureBox control.
        Dim pic As PictureBox = CType(sender, PictureBox)
        pic.BackgroundImage = CType(e.Data.GetData(DataFormats.Bitmap), Bitmap)

        ' The image in the other PictureBox (that is, the PictureBox that was
        ' not the sender in the DragDrop event) is removed if the user executes a Move.
        ' The action is a Move if the Ctrl key was not pressed.
        ''If (e.KeyState And CtrlMask) <> CtrlMask Then
        ''    If sender Is picLeft Then
        ''        picRight.BackgroundImage = Nothing
        ''    Else
        ''        picLeft.BackgroundImage = Nothing
        ''    End If
        ''End If
        Return True
    End Function


#End Region

#Region "DragDrop-TextBox"

    ''' <summary>
    ''' DragDrop, called from _DragEnter (TextDragDrop must be called as well).
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Public Shared Sub TextDragEnter(sender As Object, e As DragEventArgs)
        ' Check to be sure that the drag content is the correct type for this 
        ' control. If not, reject the drop.
        Try
            If (e.Data.GetDataPresent(DataFormats.Text)) Then
                ' If the Ctrl key was pressed during the drag operation then perform
                ' a Copy. If not, perform a Move.
                '               If (e.KeyState And CtrlMask) = CtrlMask Then
                '                    e.Effect = DragDropEffects.Copy
                '          Else
                e.Effect = DragDropEffects.Copy
                '                    e.Effect = DragDropEffects.Move
                '             End If
            Else
                e.Effect = DragDropEffects.None
            End If
        Catch
        End Try

    End Sub
    Public Shared Sub TextDragDrop(sender As Object, e As DragEventArgs)
        Dim str_ As String = e.Data.GetData(DataFormats.Text).ToString
        Dim t As TextBox = sender
        If t.Text.Length < 1 Then
            t.Text = str_
        Else
            t.Text &= vbCrLf & vbCrLf & str_
        End If
    End Sub
#End Region

#End Region

#Region "Lists - ListBox"
    Public Shared Function ListsContain(list As ListBox, what As String) As Boolean
        Return list.Items.Contains(what)
    End Function

    Public Shared Sub ListsIncludeItem(left_list As ListBox, right_list As ListBox, Optional retain_in_left As Boolean = False)
        With left_list
            If .Items.Count < 1 Or .SelectedIndex < 0 Or .SelectedIndex > .Items.Count - 1 Then Exit Sub
            right_list.Items.Add(.SelectedItem.ToString)
            If retain_in_left = False Then .Items.RemoveAt(.SelectedIndex)
        End With
    End Sub

    Public Shared Sub ListsIncludeAllItems(left_list As ListBox, right_list As ListBox)
        With left_list
            If .Items.Count < 1 Then Exit Sub
            .SelectedIndex = 0
            For ia As Integer = 0 To .Items.Count - 1
                right_list.Items.Add(.Items.Item(ia).ToString)
            Next
            .Items.Clear()
        End With
    End Sub

    Public Shared Sub ListsExcludeItem(left_list As ListBox, right_list As ListBox)
        With right_list
            If .Items.Count < 1 Or .SelectedIndex < 0 Or .SelectedIndex > .Items.Count - 1 Then Exit Sub
            left_list.Items.Add(.SelectedItem.ToString)
            .Items.RemoveAt(.SelectedIndex)
        End With
    End Sub

    Public Shared Sub ListsExcludeAllItems(left_list As ListBox, right_list As ListBox)
        With right_list
            If .Items.Count < 1 Then Exit Sub
            .SelectedIndex = 0
            For r As Integer = 0 To .Items.Count - 1
                left_list.Items.Add(.Items.Item(r).ToString)
            Next
            .Items.Clear()
        End With
    End Sub
    Public Shared Sub ListsAddItem_Multiline(list_ As ListBox, item_ As TextBox, Optional message_if_already_exists As String = "", Optional allow_duplicate As Boolean = False, Optional clear_after_add As Boolean = True)
        Dim left_list As ListBox = list_

        If item_.Text.Trim.Length < 1 Then Exit Sub

        If left_list.Items.Contains(item_.Text.Trim) Then
            If allow_duplicate Then
                If message_if_already_exists.Length > 0 Then
                    If MsgBox(message_if_already_exists, MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                Else
                End If
            Else
                Exit Sub
            End If
        End If

        Dim x As String = item_.Text.Trim.Replace(vbCr, "<New Line>")
        Dim y As String = x.Replace(vbCrLf, "<New Line>")
        Dim item__ As String = y

        With left_list
            .Items.Add(item__)
            item__ = ""
            x = "" : y = ""
            If clear_after_add = True Then item_.Text = ""
            .SelectedIndex = .Items.Count - 1
        End With
    End Sub
    Public Shared Function ListsAddItem(list_ As ListBox, item_ As String, Optional allow_duplicate As Boolean = False) As Boolean
        Dim return_ As Boolean = True
        Try

            If item_.Trim.Length > 0 Then
                If list_.Items.Contains(item_.Trim) Then
                    If allow_duplicate Then
                        list_.Items.Add(item_.Trim)
                    Else
                        return_ = False
                    End If
                Else
                    list_.Items.Add(item_.Trim)
                End If
            End If
        Catch ex As Exception

        End Try
        Return return_
    End Function

    Public Shared Function ListsAddItem(list_ As ListBox, item_ As TextBox, Optional message_if_already_exists As String = "", Optional allow_duplicate As Boolean = False, Optional clear_after_add As Boolean = True) As Boolean
        Try
            Dim left_list As ListBox = list_
            If item_.Text.Trim.Length > 0 Then
                If left_list.Items.Contains(item_.Text.Trim) Then
                    If allow_duplicate Then
                        Return AddTheItem(list_, item_, clear_after_add)
                        'warn user
                        'If message_if_already_exists.Length > 0 Then
                        '	If MsgBox(message_if_already_exists, MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then
                        '		Exit Sub
                        '	End If
                        'Else
                        'End If
                    Else
                        Return False
                    End If
                Else
                    Return AddTheItem(list_, item_, clear_after_add)
                End If
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Shared Function AddTheItem(list_ As ListBox, item_ As TextBox, Optional clear_after_add As Boolean = True) As Boolean
        Try
            Dim item__ As String = item_.Text.Trim
            Dim left_list As ListBox = list_

            With left_list
                .Items.Add(item__)
                item__ = ""
                If clear_after_add = True Then item_.Text = ""
                .SelectedIndex = .Items.Count - 1
            End With
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Shared Function ListsRetrieveItem(list_ As ListBox, Optional item_ As TextBox = Nothing)
        Try
            Dim LeftList As ListBox = list_
            With LeftList
                If .Items.Count = 0 Or .SelectedIndex < 0 Then
                    Return ""
                Else
                    Try
                        If item_ IsNot Nothing Then
                            item_.ReadOnly = True
                            item_.Text = LeftList.SelectedItem.ToString.Replace("<New Line>", vbCrLf)
                        End If
                    Catch ex As Exception
                    End Try
                    Return LeftList.SelectedItem.ToString.Replace("<New Line>", vbCrLf)
                End If
            End With
        Catch ex As Exception

        End Try


        '		RetrieveItem(TheListBox, TheTextBox)
        '		Or
        '		TheTextBox.Text = RetrieveItem(TheListBox)

    End Function

    ''' <summary>
    ''' Moves an item up the ListBox.
    ''' </summary>
    ''' <param name="list_"></param>

    Public Shared Sub ListsMoveItemUp(list_ As ListBox)
        Try
            Dim LeftList As ListBox = list_
            With LeftList
                If .Items.Count = 0 Or .SelectedIndex <= 0 Then
                    Exit Sub
                End If
                .Hide()
            End With

            Dim selected_index As Integer
            Dim before_selected_index As Integer

            Dim selected_item
            Dim before_selected_item

            Dim right_list As New List(Of String)
            Try
                right_list.Clear()
            Catch ex As Exception
            End Try

            With LeftList
                before_selected_item = .Items.Item(.SelectedIndex - 1)
                selected_item = .Items.Item(.SelectedIndex)
                before_selected_index = .SelectedIndex - 1
                selected_index = .SelectedIndex
                .SelectedIndex = 0
                With .Items
                    For i As Integer = 0 To .Count - 1
                        right_list.Add(LeftList.Items.Item(LeftList.SelectedIndex))
                        Try
                            LeftList.SelectedIndex += 1
                        Catch
                        End Try
                    Next
                    .Clear()
                End With
                With right_list
                    For j As Integer = 0 To .Count - 1
                        If j = before_selected_index Then
                            LeftList.Items.Add(selected_item)
                        ElseIf j = selected_index Then
                            LeftList.Items.Add(before_selected_item)
                        Else
                            LeftList.Items.Add(right_list(j))
                        End If
                    Next
                End With

                .SelectedIndex = before_selected_index
            End With

            Try
                right_list.Clear()
            Catch ex As Exception
            End Try

            LeftList.Show()
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Function ListsShowSelected(list_ As ListBox, Optional label_ As Control = Nothing) As String
        Try
            If list_.Items.Count = 0 Or list_.SelectedIndex < 0 Then
                If label_ IsNot Nothing Then
                    label_.Text = ""
                End If
                Return ""
            Else
                If label_ IsNot Nothing Then
                    Try
                        label_.Text = (list_.SelectedIndex + 1).ToString
                    Catch ex As Exception
                    End Try
                End If
                Return (list_.SelectedIndex + 1).ToString
            End If

        Catch ex As Exception

        End Try


        '		Me.Text = ShowSelected(TheListBox)
        '		Or, for both
        '		Me.Text = ShowSelected(TheListBox, TheLabel)
        '		Or, for label alone
        '		ShowSelected(TheListBox, TheLabel)


    End Function
    ''' <summary>
    ''' Removes an item from the ListBox.
    ''' </summary>
    ''' <param name="list_"></param>
    Public Shared Sub ListsRemoveItem(list_ As ListBox)
        Try
            Dim selected_ As Integer = list_.SelectedIndex
            If list_.Items.Count > 0 And list_.SelectedIndex >= 0 Then
                list_.Items.RemoveAt(list_.SelectedIndex)
                Try
                    list_.SelectedIndex = selected_
                Catch ex As Exception
                    Try
                        list_.SelectedIndex = selected_ - 1
                    Catch
                        Try
                            list_.SelectedIndex = list_.Items.Count - 1
                        Catch
                        End Try
                    End Try
                End Try
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Shared Function ListsIsSelected(list_ As ListBox) As Boolean
        Dim return_ As Boolean = False
        Try
            With list_
                If .Items.Count > 0 Then
                    If .SelectedIndex > -1 Then return_ = True
                End If
            End With
        Catch ex As Exception

        End Try

        Return return_
    End Function
    ''' <summary>
    ''' Removes all items of ListBox.
    ''' </summary>
    ''' <param name="list_"></param>
    ''' <returns></returns>
    Public Shared Function ListsClearList(list_ As ListBox)
        Try
            list_.DataSource = Nothing
            list_.Items.Clear()
        Catch ex As Exception

        End Try
        Return list_

    End Function

    Public Shared Sub ListsClearLists(lists As Array)
        Try
            For i = 0 To lists.Length - 1
                ListsClearList(lists(i))
                If TypeOf lists(i) Is ComboBox Then
                    CType(lists(i), ComboBox).Text = ""
                End If
                ''ListsClearList(lists(i))
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Shared Function ListsItemsToList(TheListBox As ListBox) As List(Of String)
        Dim result As New List(Of String)
        If TheListBox.Items.Count < 1 Then
            Return result
        Else
            With TheListBox.Items
                For i = 0 To .Count - 1
                    Dim current As String = .Item(i).ToString.Trim
                    If current.Length > 0 Then
                        result.Add(current)
                    End If
                Next
            End With
        End If
        Return result
    End Function
    Public Shared Function ListsItemsToList(TheComboBox As ComboBox) As List(Of String)
        Dim result As New List(Of String)
        If TheComboBox.Items.Count < 1 Then
            Return result
        Else
            With TheComboBox.Items
                For i = 0 To .Count - 1
                    Dim current As String = .Item(i).ToString.Trim
                    If current.Length > 0 Then
                        result.Add(current)
                    End If
                Next
            End With
        End If
        Return result
    End Function

#End Region

#Region "Lists - ComboBox"
    Public Shared Function ListsContain(list As ComboBox, what As String) As Boolean
        Return list.Items.Contains(what) ''Or list.Items.Contains(what.ToLower)
    End Function

    Public Shared Sub ListsIncludeItem(left_list As ComboBox, right_list As ComboBox, Optional retain_in_left As Boolean = False)
        With left_list
            If .Items.Count < 1 Or .SelectedIndex < 0 Or .SelectedIndex > .Items.Count - 1 Then Exit Sub
            right_list.Items.Add(.SelectedItem.ToString)
            If retain_in_left = False Then .Items.RemoveAt(.SelectedIndex)
        End With
    End Sub

    Public Shared Sub ListsIncludeAllItems(left_list As ComboBox, right_list As ComboBox)
        With left_list
            If .Items.Count < 1 Then Exit Sub
            .SelectedIndex = 0
            For ia As Integer = 0 To .Items.Count - 1
                right_list.Items.Add(.Items.Item(ia).ToString)
            Next
            .Items.Clear()
        End With
    End Sub

    Public Shared Sub ListsExcludeItem(left_list As ComboBox, right_list As ComboBox)
        With right_list
            If .Items.Count < 1 Or .SelectedIndex < 0 Or .SelectedIndex > .Items.Count - 1 Then Exit Sub
            left_list.Items.Add(.SelectedItem.ToString)
            .Items.RemoveAt(.SelectedIndex)
        End With
    End Sub

    Public Shared Sub ListsExcludeAllItems(left_list As ComboBox, right_list As ComboBox)
        With right_list
            If .Items.Count < 1 Then Exit Sub
            .SelectedIndex = 0
            For r As Integer = 0 To .Items.Count - 1
                left_list.Items.Add(.Items.Item(r).ToString)
            Next
            .Items.Clear()
        End With
    End Sub
    Public Shared Sub ListsAddItem_Multiline(list_ As ComboBox, item_ As TextBox, Optional message_if_already_exists As String = "", Optional allow_duplicate As Boolean = False, Optional clear_after_add As Boolean = True)
        Dim left_list As ComboBox = list_

        If item_.Text.Trim.Length < 1 Then Exit Sub

        If left_list.Items.Contains(item_.Text.Trim) Then
            If allow_duplicate Then
                If message_if_already_exists.Length > 0 Then
                    If MsgBox(message_if_already_exists, MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                Else
                End If
            Else
                Exit Sub
            End If
        End If

        Dim x As String = item_.Text.Trim.Replace(vbCr, "<New Line>")
        Dim y As String = x.Replace(vbCrLf, "<New Line>")
        Dim item__ As String = y

        With left_list
            .Items.Add(item__)
            item__ = ""
            x = "" : y = ""
            If clear_after_add = True Then item_.Text = ""
            .SelectedIndex = .Items.Count - 1
        End With
    End Sub
    Public Shared Function ListsAddItem(list_ As ComboBox, item_ As String, Optional allow_duplicate As Boolean = False) As Boolean
        Dim return_ As Boolean = True
        Try

            If item_.Trim.Length > 0 Then
                If list_.Items.Contains(item_.Trim) Then
                    If allow_duplicate Then
                        list_.Items.Add(item_.Trim)
                    Else
                        return_ = False
                    End If
                Else
                    list_.Items.Add(item_.Trim)
                End If
            End If
        Catch ex As Exception

        End Try
        Return return_
    End Function

    Public Shared Function ListsAddItem(list_ As ComboBox, item_ As TextBox, Optional message_if_already_exists As String = "", Optional allow_duplicate As Boolean = False, Optional clear_after_add As Boolean = True) As Boolean
        Try
            Dim left_list As ComboBox = list_
            If item_.Text.Trim.Length > 0 Then
                If left_list.Items.Contains(item_.Text.Trim) Then
                    If allow_duplicate Then
                        Return AddTheItem(list_, item_, clear_after_add)
                        'warn user
                        'If message_if_already_exists.Length > 0 Then
                        '	If MsgBox(message_if_already_exists, MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then
                        '		Exit Sub
                        '	End If
                        'Else
                        'End If
                    Else
                        Return False
                    End If
                Else
                    Return AddTheItem(list_, item_, clear_after_add)
                End If
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Shared Function AddTheItem(list_ As ComboBox, item_ As TextBox, Optional clear_after_add As Boolean = True) As Boolean
        Try
            Dim item__ As String = item_.Text.Trim
            Dim left_list As ComboBox = list_

            With left_list
                .Items.Add(item__)
                item__ = ""
                If clear_after_add = True Then item_.Text = ""
                .SelectedIndex = .Items.Count - 1
            End With
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Shared Function ListsRetrieveItem(list_ As ComboBox, Optional item_ As TextBox = Nothing)
        Try
            Dim LeftList As ComboBox = list_
            With LeftList
                If .Items.Count = 0 Or .SelectedIndex < 0 Then
                    Return ""
                Else
                    Try
                        If item_ IsNot Nothing Then
                            item_.ReadOnly = True
                            item_.Text = LeftList.SelectedItem.ToString.Replace("<New Line>", vbCrLf)
                        End If
                    Catch ex As Exception
                    End Try
                    Return LeftList.SelectedItem.ToString.Replace("<New Line>", vbCrLf)
                End If
            End With
        Catch ex As Exception

        End Try


        '		RetrieveItem(TheComboBoxBox, TheTextBox)
        '		Or
        '		TheTextBox.Text = RetrieveItem(TheComboBoxBox)

    End Function
    ''' <summary>
    ''' Moves an item up the ComboBox.
    ''' </summary>
    ''' <param name="list_"></param>

    Public Shared Sub ListsMoveItemUp(list_ As ComboBox)
        Try
            Dim LeftList As ComboBox = list_
            With LeftList
                If .Items.Count = 0 Or .SelectedIndex <= 0 Then
                    Exit Sub
                End If
                .Hide()
            End With

            Dim selected_index As Integer
            Dim before_selected_index As Integer

            Dim selected_item
            Dim before_selected_item

            Dim right_list As New List(Of String)
            Try
                right_list.Clear()
            Catch ex As Exception
            End Try

            With LeftList
                before_selected_item = .Items.Item(.SelectedIndex - 1)
                selected_item = .Items.Item(.SelectedIndex)
                before_selected_index = .SelectedIndex - 1
                selected_index = .SelectedIndex
                .SelectedIndex = 0
                With .Items
                    For i As Integer = 0 To .Count - 1
                        right_list.Add(LeftList.Items.Item(LeftList.SelectedIndex))
                        Try
                            LeftList.SelectedIndex += 1
                        Catch
                        End Try
                    Next
                    .Clear()
                End With
                With right_list
                    For j As Integer = 0 To .Count - 1
                        If j = before_selected_index Then
                            LeftList.Items.Add(selected_item)
                        ElseIf j = selected_index Then
                            LeftList.Items.Add(before_selected_item)
                        Else
                            LeftList.Items.Add(right_list(j))
                        End If
                    Next
                End With

                .SelectedIndex = before_selected_index
            End With

            Try
                right_list.Clear()
            Catch ex As Exception
            End Try

            LeftList.Show()
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Function ListsShowSelected(list_ As ComboBox, Optional label_ As Control = Nothing) As String
        Try
            If list_.Items.Count = 0 Or list_.SelectedIndex < 0 Then
                If label_ IsNot Nothing Then
                    label_.Text = ""
                End If
                Return ""
            Else
                If label_ IsNot Nothing Then
                    Try
                        label_.Text = (list_.SelectedIndex + 1).ToString
                    Catch ex As Exception
                    End Try
                End If
                Return (list_.SelectedIndex + 1).ToString
            End If

        Catch ex As Exception

        End Try


        '		Me.Text = ShowSelected(TheComboBoxBox)
        '		Or, for both
        '		Me.Text = ShowSelected(TheComboBoxBox, TheLabel)
        '		Or, for label alone
        '		ShowSelected(TheComboBoxBox, TheLabel)


    End Function
    ''' <summary>
    ''' Removes an item from the ComboBox.
    ''' </summary>
    ''' <param name="list_"></param>

    Public Shared Sub ListsRemoveItem(list_ As ComboBox)
        Try
            Dim selected_ As Integer = list_.SelectedIndex
            If list_.Items.Count > 0 And list_.SelectedIndex >= 0 Then
                list_.Items.RemoveAt(list_.SelectedIndex)
                Try
                    list_.SelectedIndex = selected_
                Catch ex As Exception
                    Try
                        list_.SelectedIndex = selected_ - 1
                    Catch
                        Try
                            list_.SelectedIndex = list_.Items.Count - 1
                        Catch
                        End Try
                    End Try
                End Try
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Shared Function ListsIsSelected(list_ As ComboBox) As Boolean
        Dim return_ As Boolean = False
        Try
            With list_
                If .Items.Count > 0 Then
                    If .SelectedIndex > -1 Then return_ = True
                End If
            End With
        Catch ex As Exception

        End Try

        Return return_
    End Function

    Public Shared Function ListsIsEmpty(source As Control) As Boolean
        If TypeOf source Is ComboBox Then
            Return CType(source, ComboBox).Items.Count < 1
        ElseIf TypeOf source Is ListBox Then
            Return CType(source, ListBox).Items.Count < 1
        End If
        ''        Return combobox_or_listbox.Items.Count < 1
    End Function
    Public Shared Function ListsIsEmpty(sources As Array, Optional controls_to_check As ControlsToCheck = ControlsToCheck.Any) As Boolean
        Dim counter_ As Integer = 0
        With sources
            For i As Integer = 0 To .Length - 1
                If ListsIsEmpty(sources(i)) Then
                    counter_ += 1
                End If
            Next
        End With
        If controls_to_check = ControlsToCheck.All Then
            Return Val(counter_) = sources.Length
        Else
            Return Val(counter_) > 0
        End If
    End Function
    ''' <summary>
    ''' Removes all items of ComboBox.
    ''' </summary>
    ''' <param name="list_"></param>
    ''' <returns></returns>
    Public Shared Function ListsClearList(list_ As ComboBox)
        Try
            list_.DataSource = Nothing
            list_.Items.Clear()
            list_.Text = ""
        Catch ex As Exception

        End Try
        Return list_

    End Function
#End Region


#Region "Sound"
    Public Enum MakeTheFile
        Hidden
        System
        Invisible
        Visible
    End Enum

    Private Declare Function record Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
    Private Shared Property filename__recording As String
    Public Shared Property elapsed_label__recording As Label
    Public Shared Property remaining_label__recording As Label
    Public Shared Property append_text__recording As Boolean
    Public Shared Property what_to_display_when_finished__recording As String
    Private Shared Property duration_in_seconds__recording As Long
    Private Shared Property make_the_file__recording As MakeTheFile
    Private Shared timer_recording__ As Timer
    Private Shared counter__recording As Long
    Public Structure DisplayControlsForRecordingDuration
        Public elapsed_label As Label
        Public remaining_label As Label
        Public append_text As Boolean
        Public what_to_display_when_finished As String
    End Structure
    Private Shared Sub Recording(directory_ As String, file_name_without_extension As String, Optional timer_ As Timer = Nothing, Optional duration_in_seconds As Long = 60, Optional file_extension As String = ".wav", Optional display_labels As DisplayControlsForRecordingDuration = Nothing, Optional auto_name_file As Boolean = False)
        If directory_.Length < 1 Then Exit Sub
        If file_name_without_extension.Length < 1 And auto_name_file = False Then Exit Sub

        If timer_ IsNot Nothing Then timer_.Enabled = False

        Try
            My.Computer.FileSystem.CreateDirectory(directory_)
        Catch ex As Exception
        End Try

        Dim filename_ As String = directory_ & "\"
        If auto_name_file = True Then
            filename_ &= String.Format("{0:00}", CStr(Now.Month)) & "." & String.Format("{0:00}", CStr(Now.Day)) & "." & String.Format("{0:0000}", CStr(Now.Year)) & "T" & String.Format("{0:00}", CStr(Now.Hour)) & "." & String.Format("{0:00}", CStr(Now.Minute)) & "." & String.Format("{0:00}", CStr(Now.Millisecond)) & file_extension
        Else
            filename_ &= file_name_without_extension & file_extension
        End If
        filename__recording = filename_

        If Not IsNothing(display_labels) Then
            elapsed_label__recording = display_labels.elapsed_label
            remaining_label__recording = display_labels.remaining_label
            append_text__recording = display_labels.append_text
            what_to_display_when_finished__recording = display_labels.what_to_display_when_finished
        End If

        Dim duration_ As Long = 60000
        If duration_in_seconds <> Nothing Then duration_ = duration_in_seconds
        If timer_ IsNot Nothing Then
            timer_recording__ = timer_
            AddHandler timer_.Tick, New EventHandler(AddressOf RecordingTimer)
            timer_.Interval = Val(duration_) * 1000 '* 60
            timer_.Enabled = True
        End If

        counter__recording = 0
        record("open new Type waveaudio Alias recsound", "", 0, 0)
        record("record recsound", "", 0, 0)

    End Sub
    Public Shared Sub StartRecording(directory_ As String, file_name_without_extension As String, Optional timer_ As Timer = Nothing, Optional duration_in_seconds As Long = 60, Optional file_extension As String = ".wav", Optional display_labels As DisplayControlsForRecordingDuration = Nothing, Optional auto_name_file As Boolean = False, Optional make_the_file As MakeTheFile = MakeTheFile.Visible)
        If directory_.Length < 1 Then Exit Sub
        If file_name_without_extension.Length < 1 And auto_name_file = False Then Exit Sub

        make_the_file__recording = make_the_file

        Recording(directory_, file_name_without_extension, timer_, duration_in_seconds, file_extension, display_labels, auto_name_file)
    End Sub
    Public Shared Sub EndRecording()
        record("save recsound " & filename__recording, "", 0, 0)
        record("close recsound", "", 0, 0)
        Select Case make_the_file__recording
            Case MakeTheFile.Hidden
                SetAttr(filename__recording, FileAttribute.Hidden)
            Case MakeTheFile.Invisible
                SetAttr(filename__recording, FileAttribute.Hidden + FileAttribute.System)
            Case MakeTheFile.System
                SetAttr(filename__recording, FileAttribute.System)
            Case MakeTheFile.Visible
                SetAttr(filename__recording, FileAttribute.Normal)
        End Select
    End Sub

    Private Shared Sub RecordingTimer()
        If counter__recording = duration_in_seconds__recording Then
            timer_recording__.Enabled = False
            EndRecording()
            Exit Sub
        End If
        Dim remaining As Long = duration_in_seconds__recording - counter__recording
        If elapsed_label__recording IsNot Nothing Then
            If append_text__recording Then
                elapsed_label__recording.Text = counter__recording & " seconds into recording"
            Else
                elapsed_label__recording.Text = counter__recording
            End If
        End If
        If remaining_label__recording IsNot Nothing Then
            If append_text__recording Then
                remaining_label__recording.Text = remaining & " seconds remaining"
            Else
                remaining_label__recording.Text = remaining
            End If
        End If
        counter__recording += 1
    End Sub

#End Region

#Region "Wallpaper"
    Private Const SPI_SETDESKWALLPAPER As Integer = &H14

    Private Const SPIF_UPDATEINIFILE As Integer = &H1

    Private Const SPIF_SENDWININICHANGE As Integer = &H2

    Private Declare Auto Function SystemParametersInfo Lib "user32.dll" (ByVal uAction As Integer, ByVal uParam As Integer, ByVal lpvParam As String, ByVal fuWinIni As Integer) As Integer

    ''' <summary>
    ''' Sets wallpaper. Stretch by default.
    ''' </summary>
    ''' <param name="filename">file to set as wallpaper</param>
    Public Shared Sub SetWallpaper(filename As String)
        Try

            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, filename, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

        Catch ignored As Exception


        End Try

    End Sub

    ''' <summary>
    ''' Sets wallpaper. Stretch by default.
    ''' </summary>
    ''' <param name="folder">where to select the file</param>
    ''' <param name="pickRandomly">select first (no reference to sorting) or any</param>
    ''' <param name="supportedFileTypes">preferred filetypes, defaults to {".jpg", ".png", ".bmp"}</param>
    Public Shared Sub SetWallpaper(folder As String, pickRandomly As Boolean, Optional depth As SearchOption = SearchOption.TopDirectoryOnly, Optional supportedFileTypes As String() = Nothing)
        Try
            If Not My.Computer.FileSystem.DirectoryExists(folder) Then Return

            Dim defaultFileTypes As String() = {".jpg", ".png", ".bmp"}

            Dim files As String() = GetFilesOfType(folder, If(supportedFileTypes IsNot Nothing, If(supportedFileTypes.Length > 1, supportedFileTypes, defaultFileTypes), defaultFileTypes), depth)

            If files.Length < 1 Then Return

            SetWallpaper(If(pickRandomly, PickRandomFile(files), files(0)))

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' Gets the files of a selected filetypes from a folder.
    ''' </summary>
    ''' <param name="folderPath"></param>
    ''' <param name="fileTypes">preferred filetypes, defaults to {".jpg", ".png", ".bmp"}</param>
    ''' <param name="depth"></param>
    ''' <returns></returns>
    Public Shared Function GetFilesOfType(folderPath As String, fileTypes As String(), Optional depth As SearchOption = SearchOption.TopDirectoryOnly) As String()
        Dim files As New List(Of String)

        For Each fileType As String In fileTypes
            files.AddRange(Directory.GetFiles(folderPath, "*" & fileType, depth))
        Next

        Return files.ToArray
    End Function
    ''' <summary>
    ''' Gets the files of a selected filetypes from a folder.
    ''' </summary>
    ''' <param name="folderPath"></param>
    ''' <param name="fileTypes">preferred filetypes, defaults to {".jpg", ".png", ".bmp"}</param>
    ''' <param name="depth"></param>
    ''' <returns></returns>
    Public Shared Function GetFilesOfTypeAsListOfString(folderPath As String, fileTypes As List(Of String), Optional depth As SearchOption = SearchOption.TopDirectoryOnly) As List(Of String)
        Dim files As New List(Of String)

        For Each fileType As String In fileTypes
            files.AddRange(Directory.GetFiles(folderPath, "*" & fileType, depth))
        Next

        Return files

    End Function
    Public Shared Function GetRandomFileFromFolder(folderPath As String, fileTypes As List(Of String), Optional depth As SearchOption = SearchOption.TopDirectoryOnly) As String
        Return PickRandomFile(folderPath, fileTypes, depth)
    End Function
    Public Shared Function PickRandomFile(files As String()) As String
        Return files(New Random().Next(0, files.Length))
    End Function

    Public Shared Function PickRandomFile(folderPath As String, fileTypes As List(Of String), Optional depth As SearchOption = SearchOption.TopDirectoryOnly) As String
        Return New Random().Next(0, GetFilesOfType(folderPath, fileTypes.ToArray, depth).Length)
    End Function

#End Region

End Class

Public Class ManagedPower
    ' GetSystemPowerStatus() is the only unmanaged API called.
    Declare Auto Function GetSystemPowerStatus Lib "kernel32.dll" _
    Alias "GetSystemPowerStatus" (ByRef sps As SystemPowerStatus) As Boolean

    Public Overrides Function ToString() As String
        Dim text As String = ""
        Dim sysPowerStatus As SystemPowerStatus
        ' Get the power status of the system
        If ManagedPower.GetSystemPowerStatus(sysPowerStatus) Then
            ' Current power source - AC/DC
            Dim currentPowerStatus = sysPowerStatus.ACLineStatus
            text += sysPowerStatus.ACLineStatus.ToString()
        End If
        Return text
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure SystemPowerStatus
        Public ACLineStatus As _ACLineStatus
        Public BatteryFlag As _BatteryFlag
        Public BatteryLifePercent As Byte
        Public Reserved1 As Byte
        Public BatteryLifeTime As System.UInt32
        Public BatteryFullLifeTime As System.UInt32
    End Structure

    Public Enum _ACLineStatus As Byte
        Battery = 0
        AC = 1
        Unknown = 255
    End Enum

    <Flags()>
    Public Enum _BatteryFlag As Byte
        High = 1
        Low = 2
        Critical = 4
        Charging = 8
        NoSystemBattery = 128
        Unknown = 255
    End Enum
End Class



