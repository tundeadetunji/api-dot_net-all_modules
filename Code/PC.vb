Imports System.Runtime
Imports System.Runtime.InteropServices

Public Class PC
    Private Shared Function InternetGetConnectedState(<Out> ByRef Description As Integer, ByVal ReservedValue As Integer) As Boolean

    End Function

    Public Shared Function IsConnectedToInternet() As Boolean
        Dim Desc As Integer
        Return InternetGetConnectedState(Desc, 0)
    End Function
End Class
