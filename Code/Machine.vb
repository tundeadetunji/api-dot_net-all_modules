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



End Class
