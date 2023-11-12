Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Runtime
Imports System.Runtime.InteropServices

Public Class Machine
    <DllImport("wininet.dll")>
    Private Shared Function InternetGetConnectedState(<Out> ByRef Description As Integer, ByVal ReservedValue As Integer) As Boolean

    End Function

    Public Shared Function MachineIsConnectedToInternet() As Boolean
        Dim Desc As Integer
        Return InternetGetConnectedState(Desc, 0)
    End Function



    Private Const APPCOMMAND_VOLUME_MUTE As Integer = &H80000
    Private Const APPCOMMAND_VOLUME_UP As Integer = &HA0000
    Private Const APPCOMMAND_VOLUME_DOWN As Integer = &H90000
    Private Const WM_APPCOMMAND As Integer = &H319

    <DllImport("user32.dll")>
    Public Shared Function SendMessageW(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Public Shared Sub Mute(dialog As System.Windows.Forms.Form)
        Dim h As IntPtr = dialog.Handle
        SendMessageW(h, WM_APPCOMMAND, h, CType(APPCOMMAND_VOLUME_MUTE, IntPtr))
    End Sub

    Public Shared Sub VolumeDown(dialog As System.Windows.Forms.Form)
        Dim h As IntPtr = dialog.Handle
        SendMessageW(h, WM_APPCOMMAND, h, New IntPtr(APPCOMMAND_VOLUME_DOWN))
    End Sub

    Public Shared Sub VolumeUp(dialog As System.Windows.Forms.Form)
        Dim h As IntPtr = dialog.Handle
        SendMessageW(h, WM_APPCOMMAND, h, New IntPtr(APPCOMMAND_VOLUME_UP))
    End Sub

    Public Shared Shadows Sub SetEnvironmentVariable(key As String, value As String, Optional target As EnvironmentVariableTarget = EnvironmentVariableTarget.User)
        Environment.SetEnvironmentVariable(key, value, target)
    End Sub

    Public Shared Shadows Function GetEnvironmentVariable(key As String, Optional target As EnvironmentVariableTarget = EnvironmentVariableTarget.User) As String
        Return Environment.GetEnvironmentVariable(key, target)
    End Function

    Public Shared Shadows Function GetEnvironmentVariable(key As String, default_value As String, Optional target As EnvironmentVariableTarget = EnvironmentVariableTarget.User) As String
        Return If(GetEnvironmentVariable(key, target).Trim.Length > 0, GetEnvironmentVariable(key, target), default_value)
    End Function

    Public Shared Sub SetEnvironmentVariableForUser(key As String, value As String)
        SetEnvironmentVariable(key, value, EnvironmentVariableTarget.User)
    End Sub

    Public Shared Sub SetEnvironmentVariableForSystem(key As String, value As String)
        SetEnvironmentVariable(key, value, EnvironmentVariableTarget.Machine)
    End Sub

    Public Shared Function GetEnvironmentVariableForUser(key As String, default_value As String) As String
        Return If(GetEnvironmentVariable(key, EnvironmentVariableTarget.User).Trim.Length > 0, GetEnvironmentVariable(key, EnvironmentVariableTarget.User), default_value)
    End Function

    Public Shared Function GetEnvironmentVariableForUser(key As String) As String
        Return GetEnvironmentVariable(key, EnvironmentVariableTarget.User)
    End Function

    Public Shared Function GetEnvironmentVariableForSystem(key As String) As String
        Return GetEnvironmentVariable(key, EnvironmentVariableTarget.Machine)
    End Function

    Public Shared Function GetEnvironmentVariableForSystem(key As String, default_value As String) As String
        Return If(GetEnvironmentVariable(key, EnvironmentVariableTarget.Machine).Trim.Length > 0, GetEnvironmentVariable(key, EnvironmentVariableTarget.Machine), default_value)
    End Function

End Class
