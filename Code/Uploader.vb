Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Text
Imports System.Threading.Tasks
Imports FluentFTP

Public Class Uploader
    Private Property server As String
    Private Property username As String
    Private Property password As String

    Public Sub New(ByVal ftp_server As String, ByVal username As String, ByVal password As String)
        Me.server = ftp_server
        Me.username = username
        Me.password = password
    End Sub

    Private Function CreateFtpClient() As FtpClient
        Return New FtpClient(Me.server, New NetworkCredential With {
            .UserName = username,
            .Password = password
        })
    End Function

    ''' <summary>
    ''' Uploads content of file.
    ''' </summary>
    ''' <param name="filename_with_extension"></param>
    ''' <param name="fileBytes"></param>
    ''' <param name="remoteFolderPathEscapedString"></param>
    ''' <example>
    ''' new Uploader(...).uploadBytes("1.jpg", PictureFromStream(...), "site.com\wwwroot\images\") //from regular Desktop application
    ''' new Uploader(...).uploadBytes("1.jpg", bannerPic.FileBytes, "site.com\wwwroot\images\") //from ASP.NET
    ''' </example>
    Public Function uploadBytes(ByVal filename_with_extension As String, ByVal fileBytes As Byte(), ByVal Optional remoteFolderPathEscapedString As String = "") As FtpStatus
        Using ftp As FtpClient = CreateFtpClient()
            ftp.UploadBytes(fileBytes, remoteFolderPathEscapedString & filename_with_extension)
        End Using
    End Function

    ''' <summary>
    ''' Called typically from C# desktop application (from OpenFileDialog thereabouts). 
    ''' </summary>
    ''' <param name="info"></param>
    ''' <param name="remoteFolderPathEscapedString"></param>
    Public Sub uploadFile(ByVal info As FileInfo, ByVal Optional remoteFolderPathEscapedString As String = "")
        Using ftp As FtpClient = CreateFtpClient()

            Using stream As FileStream = File.OpenRead(info.FullName)
                ftp.UploadStream(stream, remoteFolderPathEscapedString & info.Name)
            End Using
        End Using
    End Sub

    Public Sub downloadFile(remote_file_path As String, full_path_of_filename_with_extension_to_store_it_on_local_file_system As String)
        Using ftp As FtpClient = CreateFtpClient()
            ftp.Connect()
            Try
                ftp.DownloadFile(full_path_of_filename_with_extension_to_store_it_on_local_file_system, remote_file_path)
            Catch ex As Exception
            End Try
        End Using

    End Sub
End Class
