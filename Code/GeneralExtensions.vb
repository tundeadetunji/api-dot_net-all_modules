Imports System.Runtime.CompilerServices
Imports iNovation.Code.General

''' <summary>
''' This class contains extension methods based on methods from iNovation.Code.General.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: November 2024
''' </remarks>
Public Module GeneralExtensions
    <Extension()>
    Public Function EqualsIgnoreCase(ByVal str1 As String, ByVal str2 As String) As Boolean
        If str1 Is Nothing AndAlso str2 Is Nothing Then
            Return True
        End If

        If str1 Is Nothing OrElse str2 Is Nothing Then
            Return False
        End If

        Return String.Equals(str1, str2, StringComparison.OrdinalIgnoreCase)
    End Function
    <Extension()>
    Public Function ToCurrency(ByVal val_ As Integer)
        Return Desktop.ToCurrency(val_)
    End Function
    <Extension()>
    Public Function ToCurrency(ByVal val_ As Double)
        Return Desktop.ToCurrency(val_)
    End Function
    <Extension()>
    Public Function ToCurrency(ByVal val_ As Decimal)
        Return Desktop.ToCurrency(val_)
    End Function
    <Extension()>
    Public Function ToCurrency(ByVal val_ As Long)
        Return Desktop.ToCurrency(val_)
    End Function
    <Extension()>
    Public Function ToCurrency(ByVal val_ As Short)
        Return Desktop.ToCurrency(val_)
    End Function
    <Extension()>
    Public Function ToCurrency(ByVal val_ As Byte)
        Return Desktop.ToCurrency(val_)
    End Function
    <Extension()>
    Public Function ToTitleCase(input As String) As String
        ' Split the input string into lines
        Dim lines As String() = input.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

        ' Create a list to hold the capitalized lines
        Dim capitalizedLines As New List(Of String)()

        ' Process each line
        For Each line As String In lines
            ' Split the line into words
            Dim words As String() = line.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)

            ' Capitalize the first letter of each word
            For i As Integer = 0 To words.Length - 1
                If words(i).Length > 0 Then
                    words(i) = Char.ToUpper(words(i)(0)) & words(i).Substring(1).ToLower()
                End If
            Next

            ' Join the capitalized words back into a line
            capitalizedLines.Add(String.Join(" ", words))
        Next

        ' Join the capitalized lines back into a single string
        Return String.Join(Environment.NewLine, capitalizedLines)
    End Function

    ''' <summary>
    ''' Turns statement to continuous tense. Works for ~80% scenarios.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="suffix"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ToContinuous(ByVal s As String, Optional ByVal suffix As String = "") As String
        Return General.ToContinuous(s, suffix)
    End Function

    <Extension()>
    Public Function DictionaryKeys(Of K, V)(dict As Dictionary(Of K, V)) As List(Of K)
        Return dict.Keys.ToList
    End Function

    <Extension()>
    Public Function DictionaryValues(Of K, V)(dict As Dictionary(Of K, V)) As List(Of V)
        Return dict.Values.ToList
    End Function

    ''' <summary>
    ''' Returns the key of this value in the dictionary. If not found, returns Null.
    ''' </summary>
    ''' <typeparam name="K"></typeparam>
    ''' <typeparam name="V"></typeparam>
    ''' <param name="str_sought"></param>
    ''' <param name="dict"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function KeyOfValue(Of K, V)(str_sought As K, dict As Dictionary(Of K, V)) As K
        For Each _key As K In dict.Keys
            If dict.Item(_key).Equals(str_sought) Then
                Return _key
            End If
        Next
        Return Nothing
    End Function

    <Extension()>
    Public Function Flatten(Of K, V)(dict As Dictionary(Of K, V)) As IEnumerable(Of Object)
        Dim array As List(Of Object) = New List(Of Object)

        For Each kvp As KeyValuePair(Of K, V) In dict
            array.Add(kvp.Key)
            array.Add(kvp.Value)
        Next

        Return CType(array, IEnumerable(Of Object))
    End Function
    <Extension()>
    Public Function ToList(Of K, V)(dict As Dictionary(Of K, V)) As IEnumerable(Of Object)
        Dim array As List(Of Object) = New List(Of Object)

        For Each kvp As KeyValuePair(Of K, V) In dict
            array.Add(kvp.Key)
            array.Add(kvp.Value)
        Next

        Return CType(array, IEnumerable(Of Object))
    End Function
    <Extension()>
    Public Function ListToString(Of E)(list As List(Of E), Optional delimiter As String = vbCrLf, Optional format_output As Boolean = False) As String
        Return String.Join(delimiter, list)
    End Function
    <Extension()>
    Public Function ArrayToString(Of E)(list As E(), Optional delimiter As String = vbCrLf, Optional format_output As Boolean = False) As String
        Return String.Join(delimiter, list)
    End Function
    <Extension()>
    Public Function StringToList(delimited_string As String, Optional delimiter As String = vbCrLf) As List(Of String)
        Return General.StringToList(delimited_string, delimiter)
    End Function
    <Extension()>
    Public Function RemoveHTML(ByVal html_markup As String) As String
        Return General.RemoveHTMLFromText(html_markup)
    End Function
    <Extension()>
    Public Function Acronym(string_ As String, Optional return_upper_case As Boolean = True, Optional separator As String = " ")
        Return General.Acronym(string_, return_upper_case, separator)
    End Function
    <Extension()>
    Public Sub WriteToFile(s As String, file_path As String, Optional append As Boolean = False, Optional trim As Boolean = True)
        General.WriteText(file_path, s, append, trim)
    End Sub
    <Extension()>
    Public Function IsEmail(s As String) As Boolean
        Return General.IsEmail(s)
    End Function

End Module
