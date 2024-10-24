'from iNovation
Imports iNovation.Code.General

'from Assemblies
Imports System.Web.UI.HtmlControls
Public Class Bootstrap

    ''' <summary>
    ''' Based mostly on Boostrap 4.1.
    ''' </summary>
    ''' <remarks>
    ''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
    ''' Date: October 24, 2024
    ''' </remarks>
    Public Class Four

#Region "Enums"

        Public Enum BootstrapColor
            __default = 0
            primary = 1
            secondary = 2
            success = 3
            danger = 4
            warning = 5
            info = 6
            light = 7
            dark = 8
        End Enum
#End Region

        ''' <summary>
        ''' Creates div's innerHTML. Same as Div() except it includes footer's hr tag by default.
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="heading"></param>
        ''' <param name="additional_content"></param>
        ''' <param name="color"></param>
        ''' <param name="additional_classes"></param>
        ''' <param name="dismissable"></param>
        ''' <param name="replacement_links_texts"></param>
        ''' <param name="remove_footer_hr"></param>
        ''' <param name="links_are_clickable"></param>
        ''' <returns></returns>
        Public Shared Function Alert(s As String, Optional heading As String = Nothing, Optional additional_content As String = Nothing, Optional color As BootstrapColor = BootstrapColor.warning, Optional additional_classes As String = Nothing, Optional dismissable As Boolean = True, Optional replacement_links_texts As Dictionary(Of String, String) = Nothing, Optional remove_footer_hr As Boolean = False, Optional links_are_clickable As Boolean = True) As String
            If s.Trim.Length < 1 Then
                Return ""
            End If

            If links_are_clickable Then
                s = FormatTextForAlert(s, replacement_links_texts, links_are_clickable)
                heading = FormatTextForAlert(heading, replacement_links_texts, links_are_clickable)
                additional_content = FormatTextForAlert(additional_content, replacement_links_texts, links_are_clickable)
            End If

            Dim additional_classes__ As String = ""
            If additional_classes IsNot Nothing Then additional_classes__ = additional_classes

            Dim r As String = "<div class=""" & additional_classes__ & " alert alert-" & EnumToDrop(color.ToString) & """ role=""alert"">"
            If dismissable Then r = "<div class=""" & additional_classes__ & " alert alert-" & EnumToDrop(color.ToString) & " alert-dismissible fade show"" role=""alert"">"
            If heading IsNot Nothing Then r &= "<h4 class=""alert-heading"">" & heading & "</h4>"
            r &= "<p>" & s & "</p>"
            If additional_content IsNot Nothing And additional_content.Length > 0 Then
                If remove_footer_hr = False Then
                    r &= "<hr>"
                End If
                r &= "<p class=""mb-0"">" & additional_content & "</p>"
            End If

            If dismissable Then r &= "<button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">
    <span aria-hidden=""true"">&times;</span>
  </button>"

            r &= "</div>"
            Return r

        End Function
        Public Shared Function Alert(s As String, Optional heading As String = Nothing, Optional color As BootstrapColor = BootstrapColor.warning, Optional additional_classes As String = Nothing, Optional dismissable As Boolean = True, Optional replacement_links_texts As Dictionary(Of String, String) = Nothing, Optional remove_footer_hr As Boolean = False, Optional links_are_clickable As Boolean = True) As String
            If s.Trim.Length < 1 Then
                Return ""
            End If

            If links_are_clickable Then
                s = FormatTextForAlert(s, replacement_links_texts, links_are_clickable)
                heading = FormatTextForAlert(heading, replacement_links_texts, links_are_clickable)
            End If

            Dim additional_classes__ As String = ""
            If additional_classes IsNot Nothing Then additional_classes__ = additional_classes

            Dim r As String = "<div class=""" & additional_classes__ & " alert alert-" & EnumToDrop(color.ToString) & """ role=""alert"">"
            If dismissable Then r = "<div class=""" & additional_classes__ & " alert alert-" & EnumToDrop(color.ToString) & " alert-dismissible fade show"" role=""alert"">"
            If heading IsNot Nothing Then r &= "<strong>" & heading & "</strong>"
            r &= " " & s

            If dismissable Then r &= "<button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">
    <span aria-hidden=""true"">&times;</span>
  </button>"

            r &= "</div>"

            Return r

        End Function

        ''' <summary>
        ''' Creates div's innerHTML, and places it on HtmlGenericControl. Same as Div() except it includes footer's hr tag by default.
        ''' </summary>
        ''' <param name="div"></param>
        ''' <param name="s"></param>
        ''' <param name="heading"></param>
        ''' <param name="color"></param>
        ''' <param name="additional_classes"></param>
        ''' <param name="dismissable"></param>
        ''' <param name="replacement_links_texts"></param>
        ''' <param name="remove_footer_hr"></param>
        ''' <param name="links_are_clickable"></param>
        Public Shared Sub Alert(div As HtmlGenericControl, s As String, Optional heading As String = Nothing, Optional color As BootstrapColor = BootstrapColor.warning, Optional additional_classes As String = Nothing, Optional dismissable As Boolean = True, Optional replacement_links_texts As Dictionary(Of String, String) = Nothing, Optional remove_footer_hr As Boolean = False, Optional links_are_clickable As Boolean = True)
            div.InnerHtml = Alert(s, heading, color, additional_classes, dismissable, replacement_links_texts, remove_footer_hr, links_are_clickable)
        End Sub

        Public Shared Sub Alert(div As HtmlGenericControl, s As String, Optional heading As String = Nothing, Optional additional_content As String = Nothing, Optional color As BootstrapColor = BootstrapColor.warning, Optional additional_classes As String = Nothing, Optional dismissable As Boolean = True, Optional replacement_links_texts As Dictionary(Of String, String) = Nothing, Optional links_are_clickable As Boolean = True)
            div.InnerHtml = Alert(s, heading, additional_content, color, additional_classes, dismissable, replacement_links_texts, links_are_clickable)
        End Sub
        ''' <summary>
        ''' Creates div's innerHTML. Same as Alert() except it omits footer's hr tag by default.
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="heading"></param>
        ''' <param name="additional_content"></param>
        ''' <param name="color"></param>
        ''' <param name="additional_classes"></param>
        ''' <param name="dismissable"></param>
        ''' <param name="replacement_links_texts"></param>
        ''' <param name="remove_footer_hr"></param>
        ''' <param name="links_are_clickable"></param>
        ''' <returns></returns>
        Public Shared Function Div(s As String, Optional heading As String = Nothing, Optional additional_content As String = Nothing, Optional color As BootstrapColor = BootstrapColor.warning, Optional additional_classes As String = Nothing, Optional dismissable As Boolean = True, Optional replacement_links_texts As Dictionary(Of String, String) = Nothing, Optional remove_footer_hr As Boolean = True, Optional links_are_clickable As Boolean = True) As String
            If s.Trim.Length < 1 Then
                Return ""
            End If

            If links_are_clickable Then
                s = FormatTextForAlert(s, replacement_links_texts, links_are_clickable)
                heading = FormatTextForAlert(heading, replacement_links_texts, links_are_clickable)
                additional_content = FormatTextForAlert(additional_content, replacement_links_texts, links_are_clickable)
            End If

            Dim additional_classes__ As String = ""
            If additional_classes IsNot Nothing Then additional_classes__ = additional_classes

            Dim r As String = "<div class=""" & additional_classes__ & " alert alert-" & EnumToDrop(color.ToString) & """ role=""alert"">"
            If dismissable Then r = "<div class=""" & additional_classes__ & " alert alert-" & EnumToDrop(color.ToString) & " alert-dismissible fade show"" role=""alert"">"
            If heading IsNot Nothing Then r &= "<h4 class=""alert-heading"">" & heading & "</h4>"
            r &= "<p>" & s & "</p>"
            If additional_content IsNot Nothing And additional_content.Length > 0 Then
                If remove_footer_hr = False Then
                    r &= "<hr>"
                End If
                r &= "<p class=""mb-0"">" & additional_content & "</p>"
            End If

            If dismissable Then r &= "<button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">
    <span aria-hidden=""true"">&times;</span>
  </button>"

            r &= "</div>"

            Return r

        End Function
        ''' <summary>
        ''' Creates div's innerHTML, and places it on HtmlGenericControl. Same as Alert() except it omits footer's hr tag by default.
        ''' </summary>
        ''' <param name="div__"></param>
        ''' <param name="s"></param>
        ''' <param name="heading"></param>
        ''' <param name="additional_content"></param>
        ''' <param name="color"></param>
        ''' <param name="additional_classes"></param>
        ''' <param name="dismissable"></param>
        ''' <param name="replacement_links_texts"></param>
        ''' <param name="remove_footer_hr"></param>
        ''' <param name="links_are_clickable"></param>
        Public Shared Sub Div(div__ As HtmlGenericControl, s As String, Optional heading As String = Nothing, Optional additional_content As String = Nothing, Optional color As BootstrapColor = BootstrapColor.warning, Optional additional_classes As String = Nothing, Optional dismissable As Boolean = True, Optional replacement_links_texts As Dictionary(Of String, String) = Nothing, Optional remove_footer_hr As Boolean = True, Optional links_are_clickable As Boolean = True)
            Dim r As String = Div(s, heading, additional_content, color, additional_classes, dismissable, replacement_links_texts, remove_footer_hr, links_are_clickable)
            div__.InnerHtml = r
        End Sub
        ''' <summary>
        ''' Places string on HtmlGenericControl as innerHTML.
        ''' </summary>
        ''' <param name="div__"></param>
        ''' <param name="s"></param>
        Public Shared Sub Div(div__ As HtmlGenericControl, s As String)
            div__.InnerHtml = s
        End Sub

        Public Shared Function FormatTextForAlert(s As String, Optional replacement_links_texts As Dictionary(Of String, String) = Nothing, Optional links_are_clickable As Boolean = True) As String
            If s Is Nothing Then Return ""
            If s IsNot Nothing And s.Length < 1 Then Return ""

            If replacement_links_texts Is Nothing Then Return ""
            If replacement_links_texts.Count < 1 Then Return ""

            Dim u As List(Of String) = ExtractURLs(s, replacement_links_texts, links_are_clickable)
            Dim t As String = ""
            For i = 0 To u.Count - 1
                t &= u(i) & " "
            Next
            Return t.Trim
        End Function

        Public Shared Function ExtractURLs(s As String, Optional replacement_links_texts As Dictionary(Of String, String) = Nothing, Optional links_are_clickable As Boolean = True) As List(Of String)
            Dim r As New List(Of String)
            If s.Trim.Length < 1 Then Return r
            If replacement_links_texts Is Nothing Then Return r
            If replacement_links_texts.Count < 1 Then Return r
            Dim l As List(Of String) = s.Split({" "}, StringSplitOptions.None).ToList
            For i = 0 To l.Count - 1
                If l(i).ToLower.StartsWith("http") Then
                    Dim link__ As Array = separate_link_from_suffix_part_that_is_not_with_link(l(i))
                    If replacement_links_texts IsNot Nothing And replacement_links_texts.Keys.Contains(link__(0)) Then
                        'r.Add(ToAlertLink(l(i), replacement_links_texts.Item(l(i).Trim), links_are_clickable))
                        r.Add(ToAlertLink(link__(0), replacement_links_texts.Item(link__(0)), link__(1), links_are_clickable))
                    Else
                        'r.Add(ToAlertLink(l(i), l(i), links_are_clickable))
                        r.Add(ToAlertLink(link__(0), l(i), link__(0), links_are_clickable))
                    End If
                Else
                    If replacement_links_texts IsNot Nothing And replacement_links_texts.Keys.Contains(l(i).Trim) Then
                        r.Add(replacement_links_texts.Item(l(i).Trim))
                    Else 'If replacement_links_texts Is Nothing Then
                        r.Add(l(i))
                    End If
                End If
            Next
            Return r
        End Function

        Private Shared Function separate_link_from_suffix_part_that_is_not_with_link(text As String) As Array
            Dim l As String = text.Substring(0, text.Length - 1)
            Dim r__ As String = text.Substring(text.Length - 1, 1)
            If isAlphabet(r__) Then
                Return {text, ""}
            ElseIf isNotAlphabet(r__) Then
                Return {l, r__}
            End If
        End Function
        Private Shared Function isNotAlphabet(s As String) As Boolean
            Return isAlphabet(s) = False
        End Function
        Private Shared Function isAlphabet(s As String) As Boolean
            Dim t As String = s.ToLower
            Dim alphabets As String() = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}
            Return alphabets.Contains(t)
        End Function

        Public Shared Function ToAlertLink(link As String, replacement_links_text As String, suffix_part_that_is_not_with_link As String, Optional links_are_clickable As Boolean = True) As String
            If link.Trim.Length < 1 Then Return ""
            Dim r As String = ""
            If links_are_clickable Then
                If replacement_links_text.Length > 0 Then
                    r = "<a target=""_blank"" href=""" & link & """ class=""alert-link"">" & replacement_links_text & "</a>"
                Else
                    r = "<a target=""_blank"" href=""" & link & """ class=""alert-link"">" & link & "</a>"
                End If
            Else
                If replacement_links_text.Length > 0 Then
                    r = replacement_links_text
                Else
                    r = link
                End If
            End If
            Return r & suffix_part_that_is_not_with_link '"<a href=""#"" class=""alert-link""></a>"
        End Function


    End Class
End Class
