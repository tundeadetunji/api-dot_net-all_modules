Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports System.Text
Public Class Poster

#Region "fields"
    Private Const default_target_filename_with_extension As String = "info.txt"

    Private ReadOnly Property target_filename_with_extension As String
    Private ReadOnly Property remoteFolderPathEscapedString As String

    Private ReadOnly Property drone As Uploader
#End Region

#Region "functions"
    ''' <summary>
    ''' credentials.remoteFolderPathEscapedString must end with "\", e.g. domain.com\wwwroot\files\
    ''' </summary>
    ''' <param name="credentials"></param>
    Public Sub New(credentials As Stamp)
        Me.target_filename_with_extension = If(credentials.target_filename_with_extension IsNot Nothing, credentials.target_filename_with_extension, default_target_filename_with_extension)
        Me.remoteFolderPathEscapedString = ToPath(credentials.remoteFolderPathEscapedString)

        drone = New Uploader(credentials.ftp_server, credentials.username, credentials.password)
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



#End Region

End Class
