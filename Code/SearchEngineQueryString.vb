Imports System.Collections.ObjectModel
Imports iNovation.Code.General

''' <summary>
''' This class is no longer maintained.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public Class SearchEngineQueryString

    Public Class TermWithVariations
        Private _term As String
        Private _variations As List(Of String) = New List(Of String)
        Private _instance As TermWithVariations
        Private ReadOnly _operator As SearchStringOperator

        Public Sub New(search_term As String)
            _term = search_term
            _instance = Me
            _operator = SearchStringOperator.OR_
        End Sub

        Public Sub New(search_term As String, operator_ As SearchStringOperator)
            _term = search_term
            _instance = Me
            _operator = operator_
        End Sub

        Public Sub AddVariation(variation As String)
            If variation.Length < 1 Then Return
            If Not _variations.Contains(variation) Then
                _variations.Add(variation)
            End If
        End Sub

        Public Sub RemoveVariation(variation As String)
            If _variations.Contains(variation) Then
                _variations.Remove(variation)
            End If
        End Sub

        Public Function Variations() As ReadOnlyCollection(Of String)
            Dim result As ReadOnlyCollection(Of String) = _variations.AsReadOnly
            Return result
        End Function
        Public Sub ClearVariations()
            _variations = New List(Of String)
        End Sub
        Public Overrides Function ToString() As String
            Return New String(_term) 'ToDo make this more efficient
        End Function
        Public Function Name() As String
            Return New String(_term) 'ToDo make this more efficient
        End Function
        Public Function BooleanOperatorIs() As SearchStringOperator
            Return _operator
        End Function
        Public Sub Die()
            Try
                _instance = Nothing
            Catch ex As Exception

            End Try
        End Sub

    End Class

    Public Shared Function constructParameterString(parameters As List(Of String)) As String
        Dim result As String = If(parameters.Count > 1, "(", "")
        For i = 0 To parameters.Count - 1
            result &= If(IsPhraseOrSentence(parameters(i)), "(" & parameters(i) & ")", parameters(i)) & If(i <> parameters.Count - 1, " OR ", "")
        Next
        Return If(parameters.Count > 1, result & ")", result).Trim
    End Function
    Public Shared Function constructParameterString(parameters As List(Of String), boolean_operator As SearchStringOperator) As String
        Dim result As String = If(parameters.Count > 1, "(", "")
        For i = 0 To parameters.Count - 1
            result &= If(IsPhraseOrSentence(parameters(i)), "(" & parameters(i) & ")", parameters(i)) & If(i <> parameters.Count - 1, " " & boolean_operator.ToString.Replace("_", "") & " ", "")
        Next
        Return If(parameters.Count > 1, result & ")", result).Trim
    End Function
    Public Shared Function constructParameterString(parameter As String) As String
        Return "(" & parameter & ")"
    End Function
    Public Shared Function constructSiteString(TheSite As String, Optional Prepend As Boolean = True) As String
        Return If(Prepend, "site:" & TheSite, TheSite)
    End Function
    Public Shared Function constructSiteString(sites As List(Of String), Optional Prepend As Boolean = True) As String
        Dim result As String = If(sites.Count > 1, "(", "")
        For i = 0 To sites.Count - 1
            result &= constructSiteString(sites(i), Prepend) & If(i <> sites.Count - 1, " OR ", "")
        Next
        Return If(sites.Count > 1, result & ")", result).Trim
    End Function

    Public Shared Function constructQueryString(sites As List(Of String), parameters As List(Of String), Optional Prepend As Boolean = True)
        Return constructSiteString(sites, Prepend) & " " & constructParameterString(parameters)
    End Function

    Public Shared Function constructQueryString(sites As List(Of String), terms As List(Of TermWithVariations), Optional Prepend As Boolean = True)
        Dim parameters_string As String = ""

        Dim parameters As List(Of String)
        For i = 0 To terms.Count - 1
            parameters = New List(Of String)
            parameters.Add(terms(i).Name)
            For j = 0 To terms(i).Variations.Count - 1
                parameters.Add(terms(i).Variations(j))
            Next
            parameters_string &= constructParameterString(parameters, terms(i).BooleanOperatorIs) & If(i < terms.Count - 1, " AND ", "")
        Next

        Return constructSiteString(sites, Prepend) & " " & parameters_string ''& ")"
    End Function

    Public Shared Function constructQueryString(site As String, parameter As String, Optional Prepend As Boolean = True) As String
        Return constructSiteString(site, Prepend) & " " & constructParameterString(parameter)
    End Function

End Class
