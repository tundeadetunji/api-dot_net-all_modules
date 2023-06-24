Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports System.Text
Imports System.Net
Imports System.IO

Public Class Poster

#Region "fields"
    Private Const default_target_filename_with_extension As String = "info.txt"

    Private ReadOnly Property target_filename_with_extension As String
    Private ReadOnly Property remoteFolderPathEscapedString As String

    Private Property _drone As Uploader

    Private ReadOnly Property _ftp_server As String
    Private ReadOnly Property _username As String
    Private ReadOnly Property _password As String


    Private ReadOnly Property drone As Uploader
        Get
            If _drone Is Nothing Then
                _drone = New Uploader(_ftp_server, _username, _password)
            End If
            Return _drone
        End Get
    End Property
#End Region

#Region "functions"
    ''' <summary>
    ''' Initializes to send/receive data from a specified remote file
    ''' </summary>
    ''' <param name="credentials">credentials.remoteFolderPathEscapedString e.g. domain.com\wwwroot\files\</param>
    Public Sub New(credentials As Stamp)
        Me.target_filename_with_extension = If(credentials.target_filename_with_extension IsNot Nothing, credentials.target_filename_with_extension, default_target_filename_with_extension)
        Me.remoteFolderPathEscapedString = ToPath(credentials.remoteFolderPathEscapedString)

        _ftp_server = credentials.ftp_server
        _username = credentials.username
        _password = credentials.password

        'drone = New Uploader(credentials.ftp_server, credentials.username, credentials.password)
    End Sub

    Public Function SendPackage(info As Envelope) As FluentFTP.FtpStatus
        Return drone.uploadBytes(target_filename_with_extension, sealedEnvelope(info), remoteFolderPathEscapedString)
    End Function

    Public Function SendPackage(info As LargeEnvelope) As FluentFTP.FtpStatus
        Return drone.uploadBytes(target_filename_with_extension, sealedLargeEnvelope(info), remoteFolderPathEscapedString)
    End Function

    Private Function ReceivePackageAndNothingElse(folder_to_store_it_on_local_file_system As String) As String
        drone.downloadFile(remoteFolderPathEscapedString & target_filename_with_extension, ToPath(folder_to_store_it_on_local_file_system) & default_target_filename_with_extension)
        Return ToPath(folder_to_store_it_on_local_file_system) & default_target_filename_with_extension
    End Function

    Public Function ReceivePackage(folder_to_store_it_on_local_file_system As String, Optional read_package_on_receive As Boolean = True) As String
        Return If(read_package_on_receive = True, ReadText(ReceivePackageAndNothingElse(folder_to_store_it_on_local_file_system), True), ReceivePackageAndNothingElse(folder_to_store_it_on_local_file_system))
    End Function

    Private Function sealedEnvelope(info As Envelope) As Byte()
        Return Encoding.ASCII.GetBytes(info.source_machine_name & vbCrLf & info.info)
    End Function
    Private Function sealedLargeEnvelope(info As LargeEnvelope) As Byte()
        Return Encoding.ASCII.GetBytes(info.source_machine_name & vbCrLf & info.target_machine_name & vbCrLf & info.isBroadcast & vbCrLf & info.info)
    End Function

    Public Shared Function Peek(path_to_remote_resource As String) As String
        Dim client As New WebClient()
        Dim stream As Stream = client.OpenRead(path_to_remote_resource)
        Dim reader As New StreamReader(stream)
        Dim content As String = reader.ReadToEnd()
        Return content
    End Function


#End Region

End Class
