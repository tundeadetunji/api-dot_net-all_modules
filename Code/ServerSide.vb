Imports System.Net.Http
Imports System.Collections.Specialized
Imports System.Text
Imports System.Net
Imports System.IO
Imports Newtonsoft.Json
Imports FluentFTP

Public Class ServerSide

#Region "Receive"
    ''' <summary>
    ''' Gets string result of endpoint, no authentication.
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <returns></returns>
    Public Function Receive(URL As String)
        Dim response
        Using wb = New WebClient()
            response = wb.DownloadString(URL)
        End Using
        Return response
    End Function
    ''' <summary>
    ''' Same as Receive(url as string), except it returns status code alongside
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    Public Function Peek(url As String) As CustomResponseObject

        Dim httpWebRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        httpWebRequest.ContentType = "text/plain"
        httpWebRequest.Method = "GET"

        Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)

        Dim result = ""
        Using streamReader As New StreamReader(httpResponse.GetResponseStream())
            result = streamReader.ReadToEnd()
        End Using

        Return New CustomResponseObject With {.objectResponse = Nothing, .statusCode = httpResponse.StatusCode, .stringResponse = result}
    End Function

#End Region

#Region "Send"

    ''' <summary>
    ''' Send Http Request with Form Values, no authentication.
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="FormValues"></param>
    ''' <returns></returns>

    Public Function Send(URL As String, FormValues As Dictionary(Of String, String)) As String
        Dim data, response, responseInString
        Using wb = New WebClient()
            data = DictionaryToNameValueCollection(FormValues)
            response = wb.UploadValues(URL, "POST", data)
            responseInString = Encoding.UTF8.GetString(response)
        End Using
        Return responseInString
    End Function

    ''' <summary>
    ''' Http request as string, entity will be serialized as JSON.
    ''' Same as Send(url as string, entity as object) but returns status code alongside.
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="entity"></param>
    ''' <returns></returns>
    Public Function Post(url As String, entity As Object) As CustomResponseObject

        Dim httpWebRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        httpWebRequest.ContentType = "application/json"
        httpWebRequest.Method = "POST"

        Using streamWriter As New StreamWriter(httpWebRequest.GetRequestStream())
            Dim json As String = JsonConvert.SerializeObject(entity)
            streamWriter.Write(json)
        End Using

        Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
        Dim result = ""
        Using streamReader As New StreamReader(httpResponse.GetResponseStream())
            result = streamReader.ReadToEnd()
        End Using

        Return New CustomResponseObject With {.stringResponse = result, .objectResponse = Nothing, .statusCode = httpResponse.StatusCode}
    End Function

    ''' <summary>
    ''' Http request as string, entity will be serialized as JSON.
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="entity"></param>
    ''' <returns></returns>
    Public Function Send(url As String, entity As Object) As String

        Dim httpWebRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        httpWebRequest.ContentType = "application/json"
        httpWebRequest.Method = "POST"

        Using streamWriter As New StreamWriter(httpWebRequest.GetRequestStream())
            Dim json As String = JsonConvert.SerializeObject(entity)
            streamWriter.Write(json)
        End Using

        Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
        Dim result = ""
        Using streamReader As New StreamReader(httpResponse.GetResponseStream())
            result = streamReader.ReadToEnd()
        End Using

        Return result
    End Function


#End Region

#Region "Support"


    Private Function DictionaryToNameValueCollection(dictionary As Dictionary(Of String, String))
        Return dictionary.Aggregate(New NameValueCollection(), Function(seed, current)
                                                                   seed.Add(current.Key, current.Value)
                                                                   Return seed
                                                               End Function)

    End Function
#End Region

#Region "Members"
    Public Structure CustomResponseObject
        Public stringResponse As String
        Public objectResponse As Object
        Public statusCode As String
    End Structure
#End Region


#Region "Ftp"
    Class Ftp
        Private Property server As String
        Private Property username As String
        Private Property password As String

        Public Sub New(ByVal ftp_server As String, ByVal username As String, ByVal password As String)
            Me.server = ftp_server
            Me.username = username
            Me.password = password
        End Sub
        ''' <summary>
        ''' Use this if you want to call downloadFile()
        ''' </summary>
        Public Sub New()

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
        Public Sub uploadBytes(ByVal filename_with_extension As String, ByVal fileBytes As Byte(), ByVal Optional remoteFolderPathEscapedString As String = "") ''As FtpStatus
            Using ftp As FtpClient = CreateFtpClient()
                ftp.UploadBytes(fileBytes, remoteFolderPathEscapedString & filename_with_extension)
            End Using
        End Sub

        ''' <summary>
        ''' Called typically from desktop application (from OpenFileDialog thereabouts). 
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

        Public Sub uploadFile(ByVal file_to_upload As String, ByVal Optional remoteFolderPathEscapedString As String = "", ByVal Optional new_file_name_with_extension As String = Nothing)
            Dim filename As String = If(new_file_name_with_extension IsNot Nothing, new_file_name_with_extension, Path.GetFileName(file_to_upload))
            Dim text() As Byte = File.ReadAllBytes(file_to_upload)
            uploadBytes(filename, text, remoteFolderPathEscapedString)
        End Sub

        ''' <summary>
        ''' Downloads file from remote server.
        ''' </summary>
        ''' <param name="remote_file_path"></param>
        ''' <param name="full_path_of_filename_with_extension_to_store_it_on_local_file_system"></param>
        ''' <param name="attempt_to_read_package_as_text_on_receive"></param>
        ''' <returns>the content of the file (assuming it's text) if attempt_to_read_package_as_text_on_receive = true, otherwise full_path_of_filename_with_extension_to_store_it_on_local_file_system</returns>
        Public Function downloadFile(remote_file_path As String, full_path_of_filename_with_extension_to_store_it_on_local_file_system As String, Optional attempt_to_read_package_as_text_on_receive As Boolean = False) As String

            Dim webClient As New WebClient()
            webClient.DownloadFile(remote_file_path, full_path_of_filename_with_extension_to_store_it_on_local_file_system)

            Return If(attempt_to_read_package_as_text_on_receive, General.ReadText(full_path_of_filename_with_extension_to_store_it_on_local_file_system), full_path_of_filename_with_extension_to_store_it_on_local_file_system)

            'Using ftp As FtpClient = CreateFtpClient()
            '    ftp.Connect()
            '    Try
            '        ftp.DownloadFile(full_path_of_filename_with_extension_to_store_it_on_local_file_system, remote_file_path)
            '    Catch ex As Exception
            '    End Try
            'End Using

        End Function
    End Class

#End Region

End Class
