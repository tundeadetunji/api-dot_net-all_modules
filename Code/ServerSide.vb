Imports System.Net.Http
Imports System.Collections.Specialized
Imports System.Text
Imports System.Net

Public Class ServerSide
    Public Function Receive(URL As String)
        Dim response
        Using wb = New WebClient()
            response = wb.DownloadString(URL)
        End Using
        Return response
    End Function

    Private Function DictionaryToNameValueCollection(dictionary As Dictionary(Of String, String))
        Return dictionary.Aggregate(New NameValueCollection(), Function(seed, current)
                                                                   seed.Add(current.Key, current.Value)
                                                                   Return seed
                                                               End Function)

    End Function

    Public Function Send(URL As String, FormValues As Dictionary(Of String, String))
        Dim data, response, responseInString
        Using wb = New WebClient()
            data = DictionaryToNameValueCollection(FormValues)
            response = wb.UploadValues(URL, "POST", data)
            responseInString = Encoding.UTF8.GetString(response)
        End Using
        Return responseInString
    End Function


End Class
