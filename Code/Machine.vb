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
End Class
