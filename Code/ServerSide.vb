﻿Imports System.Net.Http
Imports System.Collections.Specialized
Imports System.Text
Imports System.Net
Imports System.IO
Imports Newtonsoft.Json

Public Class ServerSide

#Region "Receive"
    ''' <summary>
    ''' Gets string result of endpoint.
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

#End Region

#Region "Send"

    ''' <summary>
    ''' Send Http Request with Form Values, no authentication.
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="FormValues"></param>
    ''' <returns></returns>

    Public Function Send(URL As String, FormValues As Dictionary(Of String, String))
        Dim data, response, responseInString
        Using wb = New WebClient()
            data = DictionaryToNameValueCollection(FormValues)
            response = wb.UploadValues(URL, "POST", data)
            responseInString = Encoding.UTF8.GetString(response)
        End Using
        Return responseInString
    End Function

    ''' <summary>
    ''' Send Http Request with JSON body, no authentication.
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

End Class
