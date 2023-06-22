'from iNovation
Imports iNovation.Code.General

Public Class Language
#Region "Members"
    Public Enum Languages
        yoruba_ng = 0
        hausa_ng = 1
        igbo_ng = 2
        english_uk = 3
        english_us = 4
        french_fr = 5
    End Enum

#End Region



#Region "Words"

    Private Shared ReadOnly Property english_us As Dictionary(Of Integer, String)
        Get
            Dim words As New Dictionary(Of Integer, String)
            With words
                .Add(1, "Welcome")
                .Add(2, "Open")
            End With

        End Get
    End Property

    Private Shared ReadOnly Property english_uk As Dictionary(Of Integer, String)
        Get
            Dim words As New Dictionary(Of Integer, String)
            With words
                .Add(1, "Welcome")
                .Add(2, "Open")
            End With

        End Get
    End Property

    Private Shared ReadOnly Property yoruba_ng As Dictionary(Of Integer, String)
        Get
            Dim words As New Dictionary(Of Integer, String)
            With words
                .Add(1, "Káàbọ")
                .Add(2, "Ṣí")
            End With

        End Get
    End Property

    Private Shared ReadOnly Property french_fr As Dictionary(Of Integer, String)
        Get
            Dim words As New Dictionary(Of Integer, String)
            With words
                .Add(1, "welcome")
                .Add(2, "Open")
            End With

        End Get
    End Property
    Private Shared ReadOnly Property hausa_ng As Dictionary(Of Integer, String)
        Get
            Dim words As New Dictionary(Of Integer, String)
            With words
                .Add(1, "welcome")
                .Add(2, "Open")
            End With

        End Get
    End Property
    Private Shared ReadOnly Property igbo_ng As Dictionary(Of Integer, String)
        Get
            Dim words As New Dictionary(Of Integer, String)
            With words
                .Add(1, "welcome")
                .Add(2, "Open")
            End With

        End Get
    End Property
#End Region

#Region "Conversion"

    Public Shared Function ToLanguage(word_or_phrase As String, which_language As Language, Optional case__ As TextCase = TextCase.None) As String
        'get source
        Dim key As Integer

        'yoruba_ng
        If DictionaryContains(yoruba_ng, word_or_phrase) Then
            key = DictionaryKey(word_or_phrase, yoruba_ng)
        End If

        'igbo_ng
        If DictionaryContains(igbo_ng, word_or_phrase) Then
            key = DictionaryKey(word_or_phrase, igbo_ng)
        End If

        'hausa_ng
        If DictionaryContains(hausa_ng, word_or_phrase) Then
            key = DictionaryKey(word_or_phrase, hausa_ng)
        End If

        'english_us
        If DictionaryContains(english_us, word_or_phrase) Then
            key = DictionaryKey(word_or_phrase, english_us)
        End If

        'english_uk
        If DictionaryContains(english_uk, word_or_phrase) Then
            key = DictionaryKey(word_or_phrase, english_uk)
        End If

        'french_fr
        If DictionaryContains(french_fr, word_or_phrase) Then
            key = DictionaryKey(word_or_phrase, french_fr)
        End If



        'convert
        Dim r As String
        'yoruba_ng
        If which_language.Equals(Languages.yoruba_ng) Or which_language.Equals(Languages.yoruba_ng.ToString) Then
            r = DictionaryValue(yoruba_ng, key)
        End If
        'igbo_ng
        If which_language.Equals(Languages.igbo_ng) Or which_language.Equals(Languages.igbo_ng.ToString) Then
            r = DictionaryValue(igbo_ng, key)
        End If
        'hausa_ng
        If which_language.Equals(Languages.hausa_ng) Or which_language.Equals(Languages.hausa_ng.ToString) Then
            r = DictionaryValue(hausa_ng, key)
        End If
        'english_us
        If which_language.Equals(Languages.english_us) Or which_language.Equals(Languages.english_us.ToString) Then
            r = DictionaryValue(english_us, key)
        End If
        'english_uk
        If which_language.Equals(Languages.english_uk) Or which_language.Equals(Languages.english_uk.ToString) Then
            r = DictionaryValue(english_uk, key)
        End If
        'french_fr
        If which_language.Equals(Languages.french_fr) Or which_language.Equals(Languages.french_fr.ToString) Then
            r = DictionaryValue(french_fr, key)
        End If

        Return TransformText(r, case__)
    End Function

#End Region


End Class
