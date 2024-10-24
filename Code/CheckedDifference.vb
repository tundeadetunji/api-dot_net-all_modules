
Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Text
Imports iNovation.Code.General


''' <summary>
''' This class is used to check similarity and contrast between two strings.
''' <para>
''' Example:
''' String a = "abc";
''' String b = "ade";
''' CheckedDifference checked = CheckedDifference.getInstance(a, b);
''' Terminology revolves around "but" and "and", so
''' start typing checked.a will generate (by intellisense) all methods with and
''' likewise checked.b will generate all methods with but
''' </para>
''' This class is no longer maintained.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public NotInheritable Class CheckedDifference
    Private ReadOnly first As String
    Private ReadOnly second As String

    Private common As String = ""
    Private restOfFirst As String = ""
    Private restOfSecond As String = ""

    Private Shared instance As CheckedDifference

    Private linesCommon As IList(Of Integer) = New List(Of Integer)()

    ''' <summary>
    ''' Instantiates a new CheckedDifference.
    ''' </summary>
    ''' <param name="first">  the initial string </param>
    ''' <param name="second"> the string to compare with first </param>
    Private Sub New(ByVal first As String, ByVal second As String)
        If (first.Trim().Length = 0 OrElse first.Length = 0) OrElse (second.Trim().Length = 0 OrElse second.Length = 0) Then
            Throw New System.Exception("The string or the one to be compared with it not supplied!")
        End If

        Me.first = first
        Me.second = second
    End Sub

    Public Shared Function getInstance(ByVal first As String, ByVal second As String) As CheckedDifference
        If instance Is Nothing Then
            instance = New CheckedDifference(first, second)
        End If
        Return instance
    End Function

    Private alreadyCalledGetCharacterDifference As Boolean = False

    Private Sub getCharacterDifference()
        If andTheyAreTheSame() Then
            Return
        End If

        Dim commonBuilder As New StringBuilder()
        Dim thisIsShorter As String = If(butFirstIsLonger() AndAlso Not butSecondIsLonger(), second, first)

        For i As Integer = 0 To thisIsShorter.Length - 1
            If first.Chars(i) = second.Chars(i) Then
                commonBuilder.Append(first.Chars(i))
            ElseIf i = thisIsShorter.Length Then
                common = commonBuilder.ToString()
                If butFirstIsLonger() Then
                    restOfFirst = first.Substring(i)
                Else
                    restOfSecond = second.Substring(i)
                End If
                Exit For
            Else
                common = commonBuilder.ToString()
                restOfFirst = first.Substring(i)
                restOfSecond = second.Substring(i)
                Exit For
            End If
        Next i
        common = commonBuilder.ToString()
        alreadyCalledGetCharacterDifference = True
    End Sub

    Private alreadyCalledGetLineDifference As Boolean = False

    Private Sub getLineDifference()
        If andTheyAreTheSame() Then
            Return
        End If

        Dim linesInFirst As IList(Of String) = StringToList(first)
        Dim linesInSecond As IList(Of String) = StringToList(second)

        Dim thisIsShorter As IList(Of String) = If(linesInFirst.Count > linesInSecond.Count, linesInSecond, linesInFirst)

        For i As Integer = 0 To thisIsShorter.Count - 1
            If linesInFirst(i).Equals(linesInSecond(i), StringComparison.OrdinalIgnoreCase) Then
                linesCommon.Add(i + 1)
            End If
        Next i
        alreadyCalledGetLineDifference = True
    End Sub

    Public Function andTheyAreTheSame() As Boolean
        Return first.Equals(second, StringComparison.OrdinalIgnoreCase)
    End Function

    Public Function butFirstIsLonger() As Boolean
        Return first.Length > second.Length
    End Function

    Public Function butSecondIsLonger() As Boolean
        Return second.Length > first.Length
    End Function

    ''' <summary>
    ''' Checks if any text is common and captures the corresponding line numbers.
    ''' </summary>
    ''' <returns> the list of line numbers that contain common text </returns>
    Public Function andFoundTheseLinesInCommon() As ReadOnlyCollection(Of Integer)
        If Not alreadyCalledGetLineDifference Then
            getLineDifference()
        End If
        Dim result As New ReadOnlyCollection(Of Integer)(linesCommon) '' = linesCommon.

        Return result
    End Function

    Public Function butFoundNothingInCommon() As Boolean
        If andTheyAreTheSame() Then
            Return False
        End If
        If Not alreadyCalledGetCharacterDifference Then
            getCharacterDifference()
        End If
        Return common.Length < 1
    End Function

    ''' <summary>
    ''' Finds common text, starting from the beginning.
    ''' If at any point it sees difference, it concludes.
    ''' </summary>
    ''' <returns> text that is found common </returns>
    Public Function andFoundThisInCommon() As String
        If Not alreadyCalledGetCharacterDifference Then
            getCharacterDifference()
        End If
        Return (New StringBuilder()).Append(common).ToString()
    End Function

    Public Function andFoundThisUniqueToFirst() As String
        If Not alreadyCalledGetCharacterDifference Then
            getCharacterDifference()
        End If
        Return (New StringBuilder()).Append(restOfFirst).ToString()
    End Function

    Public Function andFoundThisUniqueToSecond() As String
        If Not alreadyCalledGetCharacterDifference Then
            getCharacterDifference()
        End If
        Return (New StringBuilder()).Append(restOfSecond).ToString()
    End Function


End Class

