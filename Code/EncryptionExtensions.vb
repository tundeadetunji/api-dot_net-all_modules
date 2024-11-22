Imports System.Runtime.CompilerServices
''' <summary>
''' This class contains extension methods based on methods from iNovation.Code.Encryption.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: November 2024
''' </remarks>
Public Module EncryptionExtensions
    <Extension()>
    Public Function Encrypt(ByVal s As String, ByVal key As String) As String
        Return Encryption.Encrypt(key, s)
    End Function
    <Extension()>
    Public Function Decrypt(ByVal s As String, ByVal key As String) As String
        Return Encryption.Decrypt(key, s)
    End Function

End Module
