Imports iNovation.Code.General
Imports iNovation.Code.Web
Imports iNovation.Code.Sequel
Imports System.Web.UI.HtmlControls

''' <summary>
''' Uses shoppy's CSS/JS. Use to style the dashboard UI viewed directly by the end-user.
''' https://w3layouts.com/template/shoppy-e-commerce-admin-panel-responsive-web-template/
''' https://creativecommons.org/licenses/by/3.0/
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: October 24, 2024
''' </remarks>
Public Class Charts

#Region "Pie"

    Public Shared Function PieChart(l_values As List(Of String), Optional l_colors As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Return Chart_Pie_String(l_values, l_colors, width, height)
    End Function
    Public Shared Function PieChart(l_values As List(Of String), l_labels As List(Of String), div As HtmlGenericControl, Optional legend As HtmlGenericControl = Nothing, Optional l_colors As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Dim s As String = Chart_Pie_String(l_values, l_colors, width, height)
        If div IsNot Nothing Then div.InnerHtml = s
        Return s
    End Function

    Private Shared Function Chart_Pie_String(l_values As List(Of String), Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Dim id = NewGUID()
        Dim s As String = "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """ style=""width: " & width & "px; height: " & height & "px""></canvas>
			<script>
				var pieData = ["
        s &= Chart_Pie_Variable(l_values, COLORS)
        s &= "];
					new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Pie(pieData);
			  </script>"
        Return s
    End Function

    Private Shared Function Chart_Pie_Variable(l_values As List(Of String), COLORS As List(Of String))
        Dim l As List(Of String) = l_values
        Dim s As String = "", color As String

        For i = 0 To l.Count - 1
            If COLORS IsNot Nothing Then
                color = COLORS(i)
            Else
                color = RandomColor(New List(Of Integer))
            End If
            s &= "{
						value: " & l(i) & ",
						color: """ & color & """
				  }"
            If i <> l.Count - 1 Then s &= ","
        Next
        Return s
    End Function
#End Region

#Region "Doughnut"

    Public Shared Function DoughnutChart(l_values As List(Of String), Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Return Chart_Doughnut_String(l_values, COLORS, width, height)
    End Function
    Public Shared Function DoughnutChart(l_values As List(Of String), div As HtmlGenericControl, Optional legend As HtmlGenericControl = Nothing, Optional l_labels As List(Of String) = Nothing, Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Dim s As String = Chart_Doughnut_String(l_values, COLORS, width, height)
        If div IsNot Nothing Then div.InnerHtml = s
        Return s
    End Function

    Private Shared Function Chart_Doughnut_String(l_values As List(Of String), Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
        Dim id = NewGUID()
        Dim s As String = "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """ style=""width: " & width & "px; height: " & height & "px""></canvas>
			<script>
				var doughnutData = ["
        s &= Chart_Doughnut_Variable(l_values, COLORS)
        s &= "];
					new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Doughnut(doughnutData);
			  </script>"
        Return s
    End Function
    Private Shared Function Chart_Doughnut_Variable(l_values As List(Of String), COLORS As List(Of String))
        Dim l As List(Of String) = l_values
        Dim s As String = "", color As String

        For i = 0 To l.Count - 1
            If COLORS IsNot Nothing Then
                color = COLORS(i)
            Else
                color = RandomColor(New List(Of Integer))
            End If
            s &= "{
						value: " & l(i) & ",
						color: """ & color & """
				  }"
            If i <> l.Count - 1 Then s &= ","
        Next
        Return s
    End Function
#End Region

#Region "Line"

    Public Shared Function LineChart(l_values As List(Of String), l_labels As List(Of String), Optional width As Integer = 400, Optional height As Integer = 300) As String
        Return Chart_Line_String(l_values, l_labels, width, height)
    End Function
    Public Shared Function LineChart(l_values As List(Of String), l_labels As List(Of String), div As HtmlGenericControl, Optional width As Integer = 400, Optional height As Integer = 300) As String
        Dim s As String = Chart_Line_String(l_values, l_labels, width, height)
        div.InnerHtml = s
        Return s
    End Function

    Private Shared Function Chart_Line_String(l_values As List(Of String), l_labels As List(Of String), Optional width As Integer = 400, Optional height As Integer = 300) As String
        Dim id = NewGUID()
        Dim s As String = "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """ style=""width: " & width & "px; height: " & height & "px""></canvas>
			<script>
				var lineChartData = {"
        s &= Chart_Line_Variable(l_values, l_labels)
        s &= "new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Line(lineChartData);
			  </script>"
        Return s
    End Function
    Private Shared Function Chart_Line_Variable(l_values As List(Of String), l_labels As List(Of String))
        Dim l_string As String = "labels : ["
        For i = 0 To l_labels.Count - 1
            l_string &= """" & l_labels(i) & """"
            If i <> l_labels.Count - 1 Then l_string &= ","
        Next
        Dim fc = RandomColor(New List(Of Integer))
        Dim c = RandomColor(New List(Of Integer))
        l_string &= "],"

        Dim v_string As String = "datasets: [
										{
											fillColor: """ & fc & """,
											strokeColor: """ & c & """,
											pointColor: """ & c & """,
											pointStrokeColor: ""#fff"",
											data: ["
        For j = 0 To l_values.Count - 1
            v_string &= l_values(j)
            If j <> l_values.Count - 1 Then v_string &= ","
        Next
        v_string &= "]}]};"
        Return l_string & v_string
    End Function

#End Region

#Region "Bar"

    Public Shared Function BarChart(labels As List(Of String), values As List(Of String), Optional width As Integer = 400, Optional height As Integer = 300) As String
        Dim id As String = NewGUID()
        Return "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """
            style=""width: " & width & "px; height: " & height & "px""></canvas>
        <script>
            var barChartData = {
                labels: [" & bar_labels_string(labels) & "],
                datasets: [{ fillColor: """ & Web.RandomColor(New List(Of Integer)) & """, data: [" & bar_dataset_string(values) & "] }]
            };
            new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Bar(barChartData);
        </script>"
    End Function
    Public Shared Sub BarChart(labels As List(Of String), values As List(Of String), div_for_chart As HtmlGenericControl, Optional width As Integer = 400, Optional height As Integer = 300)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = BarChart(labels, values, width, height)
    End Sub
    Private Shared Function bar_labels_string(labels As List(Of String)) As String
        Dim result As String = ""
        For i = 0 To labels.Count - 1
            result &= """" & labels(i) & """" & If(i < labels.Count - 1, ", ", "")
        Next
        Return result
    End Function

    Private Shared Function bar_dataset_string(dataset As List(Of String)) As String
        Dim result As String = ""
        For i = 0 To dataset.Count - 1
            result &= dataset(i) & If(i < dataset.Count - 1, ", ", "")
        Next
        Return result
    End Function


    ''' <summary>
    ''' Discontinued. Prefer BarChart(List(Of String), List(Of String), Integer, Integer)
    ''' l_dataset.Length must be equal to x_labels.Length.
    ''' </summary>
    ''' <param name="x_labels"></param>
    ''' <param name="l_dataset"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <returns></returns>
    Public Shared Function BarChart(x_labels As List(Of String), l_dataset As List(Of BarChartDataSet), Optional width As Integer = 400, Optional height As Integer = 300)
        Return Chart_Bar_String(x_labels, l_dataset, width, height)
    End Function
    ''' <summary>
    ''' Discontinued. Prefer BarChart(List(Of String),List(Of String), HtmlGenericControl, Integer, Integer).
    ''' </summary>
    ''' <param name="x_labels"></param>
    ''' <param name="l_dataset"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <returns></returns>
    Public Shared Function BarChart(x_labels As List(Of String), l_dataset As List(Of BarChartDataSet), div_for_chart As HtmlGenericControl, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square, Optional width As Integer = 400, Optional height As Integer = 300)
        Dim s = Chart_Bar_String(x_labels, l_dataset, width, height)

        Dim colors As New List(Of String)
        For i = 0 To l_dataset.Count - 1
            colors.Add(l_dataset(i).color_)
        Next

        If div_for_legend IsNot Nothing Then
            Dim legend_values As New List(Of String)
            For i = 0 To l_dataset.Count - 1
                legend_values.Add(l_dataset(i).legend_value)
            Next

            Legend(legend_values, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing And chart_title IsNot Nothing Then
            div_for_title.InnerText = chart_title
        End If

        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s
        Return s
    End Function

    Private Shared Function Chart_Bar_String(x_labels As List(Of String), l_dataset As List(Of BarChartDataSet), Optional width As Integer = 400, Optional height As Integer = 300) As String
        Dim id = NewGUID()
        Dim s As String = "<canvas id=""" & id & """ width=""" & width & """ height=""" & height & """ style=""width: " & width & "px; height: " & height & "px""></canvas>
        <script>
        	var barChartData = {"
        s &= Chart_Bar_Variable(x_labels, l_dataset)
        s &= "};
        		new Chart(document.getElementById(""" & id & """).getContext(""2d"")).Bar(barChartData);
          </script>"

        Return s
    End Function
    Private Shared Function Chart_Bar_Variable(x_labels As List(Of String), datasets As List(Of BarChartDataSet))
        Dim s As String = "labels: ["
        For i = 0 To x_labels.Count - 1
            s &= """" & x_labels(i) & """"
            If i <> x_labels.Count - 1 Then
                s &= ","
            End If
        Next
        s &= "],
                            datasets: ["

        For i = 0 To datasets.Count - 1
            s &= "{fillColor: """ & datasets(i).color_ & """, data: ["
            For j = 0 To datasets(i).x_values_for_each_x_label.Count - 1
                s &= datasets(i).x_values_for_each_x_label(j)
                If j <> datasets(i).x_values_for_each_x_label.Count - 1 Then
                    s &= ","
                End If
            Next
            s &= "]}"
            If i <> datasets.Count - 1 Then
                s &= ","
            End If
        Next
        s &= "]"
        Return s
    End Function


#End Region

#Region "Progress Bars"

    '-------------------
    'Bars
    '-------------------
    'Private Enum OperationType
    '    Count
    '    MIN
    '    MAX
    '    Sum
    '    AVG

    'End Enum

    'Public Shared Function ProgressBar(label As String, value As Object, div As System.Web.UI.HtmlControls.HtmlGenericControl, Optional title_ As String = Nothing) As String
    '    Dim title__ As String = ""
    '    If title_ IsNot Nothing Then title__ = title_

    '    Dim r As String = "<div class=""home-progres-main"">
    '					<h3>" & title__ & "</h3>
    '				</div>
    '				<div class='bar_group'>"

    '    r &= "<div class='bar_group__bar thin' label='" & label & "' show_values='true' tooltip='true' value='" & value & "'></div>"
    '    r &= "</div>"
    '    WriteContent(r, div)
    '    Return r
    'End Function

    'Public Shared Function ProgressBar(labels_values As List(Of Object), div As System.Web.UI.HtmlControls.HtmlGenericControl, Optional title_ As String = Nothing) As String
    '    Dim title__ As String = ""
    '    If title_ IsNot Nothing Then title__ = title_

    '    Dim r As String = "<div class=""home-progres-main"">
    '					<h3>" & title__ & "</h3>
    '				</div>
    '				<div class='bar_group'>"

    '    With labels_values
    '        If .Count > 0 Then
    '            For i As Integer = 0 To .Count - 1 Step 2
    '                r &= "<div class='bar_group__bar thin' label='" & labels_values(i) & "' show_values='true' tooltip='true' value='" & labels_values(i + 1) & "'></div>"
    '            Next
    '        End If
    '    End With
    '    r &= "</div>"
    '    WriteContent(r, div)
    '    Return r
    'End Function

    'Private Shared Function BarsO(g As GridView, div As System.Web.UI.HtmlControls.HtmlGenericControl, table_ As String, connection_string As String, field_to_group As String, field_to_apply_function_on As String, function_ As OperationType, where_keys As Array, where_keys_values As Array, Optional title_ As String = Nothing) As String
    '    Dim title__ As String = ""
    '    If title_ IsNot Nothing Then title__ = title_

    '    Dim r As String = "<div class=""home-progres-main"">
    '					<h3>" & title__ & "</h3>
    '				</div>
    '				<div class='bar_group'>"

    '    Dim q As String
    '    Select Case function_
    '        Case OperationType.AVG
    '            q = BuildAVGString_GROUPED(table_, field_to_group, field_to_apply_function_on, where_keys)
    '        Case OperationType.Count
    '            q = BuildCountString_GROUPED(table_, field_to_apply_function_on, field_to_group, where_keys)
    '        Case OperationType.Sum
    '            q = BuildSumString_GROUPED(table_, field_to_group, field_to_apply_function_on, where_keys)
    '    End Select
    '    Display(g, q, connection_string, where_keys_values)
    '    With g
    '        If .Rows.Count > 0 Then
    '            .Visible = True
    '            For i As Integer = 0 To .Rows.Count - 1
    '                r &= "<div class='bar_group__bar thin' label='" & .Rows(i).Cells(0).Text & "' show_values='true' tooltip='true' value='" & ToCurrency(.Rows(i).Cells(1).Text) & "'></div>"
    '            Next
    '        End If
    '    End With
    '    r &= "</div>"
    '    WriteContent(r, div)
    '    Return r
    'End Function
#End Region

#Region "Pie From Queries"
    ''' <summary>
    ''' counts the number of rows where each distinct item occur
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_count"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_count = quater: first(2), second(2), third(3)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function PieCount(table As String, field_to_count As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim field_to_group As String = field_to_count
        Dim query = BuildCountString_GROUPED(table, field_to_count, field_to_group, keys_)
        Dim g As DataTable = QDataTable(query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Counts of " & field_to_count
                Catch ex As Exception

                End Try

            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieCountFromExcel(table As String, field_to_count As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim field_to_group As String = field_to_count
        Dim query = BuildCountString_GROUPED_Excel(table, field_to_count, field_to_group, where_keys, where_values)
        Dim g As DataTable = QDataTableFromExcel(query, connection_string)

        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Counts of " & field_to_count
                Catch ex As Exception

                End Try

            End If
        End If

        ''g.Dispose()

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function

    ''' <summary>
    ''' counts the number of rows where each distinct item occur, but between a period specified by field_for_interval (from interval_from to interval_to)
    ''' ToDo: Update BuildCountString_GROUPED_BETWEEN to be able to use keys_ and/or keys_values (i.e. include Where clause)
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_count"></param>
    ''' <param name="field_for_interval"></param>
    ''' <param name="interval_from"></param>
    ''' <param name="interval_to"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_count = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(1), second(2), third(2)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function PieCountAcrossInterval(table As String, field_to_count As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim field_to_group As String = field_to_count
        Dim query = BuildCountString_GROUPED_BETWEEN(table, field_to_group, field_to_count, {field_for_interval})
        Dim g As DataTable = QDataTable(query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Counts of " & field_to_count & " between " & Date.Parse(interval_from).ToShortDateString & " and " & Date.Parse(interval_to).ToShortDateString

                Catch ex As Exception
                    div_for_title.InnerText = "Counts of " & field_to_count & " between " & interval_from & " and " & interval_to
                End Try

            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    ''' <summary>
    ''' WIP
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_count"></param>
    ''' <param name="field_for_interval"></param>
    ''' <param name="interval_from"></param>
    ''' <param name="interval_to"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="where_keys"></param>
    ''' <param name="where_values"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <returns></returns>
    'Private Shared Function PieCountAcrossIntervalFromExcel(table As String, field_to_count As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
    '    Dim field_to_group As String = field_to_count
    '    Dim query = BuildCountString_GROUPED_BETWEEN_Excel(table, field_to_group, field_to_count, interval_from, interval_to) 'ToDo/WIP should be able to use where_keys as well in the query or otherwise
    '    Dim g As DataTable = QDataTableFromExcel(query, connection_string)
    '    Dim list_labels As New List(Of String)
    '    Dim list_values As New List(Of String)
    '    With g
    '        For i = 0 To .Rows.Count - 1
    '            list_labels.Add(.Rows(i).Item(0).ToString)
    '            list_values.Add(.Rows(i).Item(1).ToString)
    '        Next
    '    End With

    '    Dim colors As New List(Of String)
    '    With colors
    '        For i = 0 To g.Rows.Count - 1
    '            .Add(RandomColor(New List(Of Integer)))
    '        Next
    '    End With

    '    If div_for_legend IsNot Nothing Then
    '        Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
    '    End If

    '    If div_for_title IsNot Nothing Then
    '        If chart_title IsNot Nothing Then
    '            div_for_title.InnerText = chart_title
    '        Else
    '            Try
    '                div_for_title.InnerText = "Counts of " & field_to_count & " between " & Date.Parse(interval_from).ToShortDateString & " and " & Date.Parse(interval_to).ToShortDateString

    '            Catch ex As Exception
    '                div_for_title.InnerText = "Counts of " & field_to_count & " between " & interval_from & " and " & interval_to
    '            End Try

    '        End If
    '    End If

    '    Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    'End Function

    ''' <summary>
    ''' sums the number of rows where each distinct item occur
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_group"></param>
    ''' <param name="field_to_sum"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_group = quater, field_to_sum = drink: first(75), second(155), third(95)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function PieSum(table As String, field_to_group As String, field_to_sum As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildSumString_GROUPED(table, field_to_group, field_to_sum, keys_)
        Dim g As DataTable = QDataTable(query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Sums of " & field_to_sum

                Catch ex As Exception

                End Try


            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieSumFromExcel(table As String, field_to_group As String, field_to_sum As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildSumString_GROUPED_Excel(table, field_to_group, field_to_sum, where_keys, where_values)
        Dim g As DataTable = QDataTableFromExcel(query, connection_string)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Sums of " & field_to_sum

                Catch ex As Exception

                End Try


            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function

    ''' <summary>
    ''' sums the number of rows where each distinct item occur, but between a period specified by field_for_interval (from interval_from to interval_to)
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_sum"></param>
    ''' <param name="field_to_group"></param>
    ''' <param name="field_for_interval"></param>
    ''' <param name="interval_from"></param>
    ''' <param name="interval_to"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_sum = drink, field_to_group = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(25), second(155), third(65)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function PieSumAcrossInterval(table As String, field_to_sum As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query As String = BuildSumString_GROUPED_BETWEEN(table, field_to_group, field_to_sum, {field_for_interval})
        Dim g As DataTable = QDataTable(query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Sums of " & field_to_sum & " between " & Date.Parse(interval_from).ToShortDateString & " and " & Date.Parse(interval_to).ToShortDateString

                Catch ex As Exception
                    div_for_title.InnerText = "Sums of " & field_to_sum & " between " & interval_from & " and " & interval_to

                End Try

            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function

    ''' <summary>
    ''' averages the number of rows where each distinct item occur
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_average"></param>
    ''' <param name="field_to_group"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_group = quater, field_to_average = drink: first(37), second(77), third(31)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function PieAverage(table As String, field_to_average As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildAVGString_GROUPED(table, field_to_group, field_to_average, keys_)
        Dim g As DataTable = QDataTable(query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Averages of " & field_to_average

                Catch ex As Exception

                End Try

            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function

    Public Shared Function PieAverageFromExcel(table As String, field_to_average As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildAVGString_GROUPED_Excel(table, field_to_group, field_to_average, where_keys, where_values)
        Dim g As DataTable = QDataTableFromExcel(query, connection_string)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Averages of " & field_to_average

                Catch ex As Exception

                End Try

            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function

    ''' <summary>
    ''' averages the number of rows where each distinct item occur, but between a period specified by field_for_interval (from interval_from to interval_to)
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_average"></param>
    ''' <param name="field_to_group"></param>
    ''' <param name="field_for_interval"></param>
    ''' <param name="interval_from"></param>
    ''' <param name="interval_to"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_average = drink, field_to_group = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(25), second(77), third(32)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function PieAverageAcrossInterval(table As String, field_to_average As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query As String = BuildAVGString_GROUPED_BETWEEN(table, field_to_group, field_to_average, {field_for_interval})
        Dim g As DataTable = QDataTable(query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Averages of " & field_to_average & " between " & Date.Parse(interval_from).ToShortDateString & " and " & Date.Parse(interval_to).ToShortDateString

                Catch ex As Exception
                    div_for_title.InnerText = "Averages of " & field_to_average & " between " & interval_from & " and " & interval_to

                End Try

            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function

    Private Shared Function PieTop(table As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional rows_to_select As Long = 10, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.DESC, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildTopString_GROUPED(table, field_to_group, keys_, keys_values, rows_to_select, OrderByField, order_by)
        Dim g As DataTable = QDataTable(query, connection_string, keys_values)

        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Top " & rows_to_select & " " & field_to_group

                Catch ex As Exception

                End Try

            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
#End Region

#Region "Doughnut From Queries"
    ''' <summary>
    ''' counts the number of rows where each distinct item occur
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_count"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_count = quater: first(2), second(2), third(3)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function DoughnutCount(table As String, field_to_count As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim field_to_group As String = field_to_count
        Dim query = BuildCountString_GROUPED(table, field_to_count, field_to_group, keys_)
        Dim g As DataTable = QDataTable(query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Counts of " & field_to_count

                Catch ex As Exception

                End Try

            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    Public Shared Function DoughnutCountFromExcel(table As String, field_to_count As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim field_to_group As String = field_to_count
        Dim query = BuildCountString_GROUPED_Excel(table, field_to_count, field_to_group, where_keys, where_values)
        Dim g As DataTable = QDataTableFromExcel(query, connection_string)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Counts of " & field_to_count

                Catch ex As Exception

                End Try

            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    ''' <summary>
    ''' counts the number of rows where each distinct item occur, but between a period specified by field_for_interval (from interval_from to interval_to)
    ''' ToDo: Update BuildCountString_GROUPED_BETWEEN to be able to use keys_ and/or keys_values (i.e. include Where clause)
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_count"></param>
    ''' <param name="field_for_interval"></param>
    ''' <param name="interval_from"></param>
    ''' <param name="interval_to"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_count = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(1), second(2), third(2)
    ''' </example>
    ''' <returns></returns>

    Public Shared Function DoughnutCountAcrossInterval(table As String, field_to_count As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim field_to_group As String = field_to_count
        Dim query = BuildCountString_GROUPED_BETWEEN(table, field_to_group, field_to_count, {field_for_interval})
        Dim g As DataTable = QDataTable(query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Counts of " & field_to_count & " between " & Date.Parse(interval_from).ToShortDateString & " and " & Date.Parse(interval_to).ToShortDateString

                Catch ex As Exception
                    div_for_title.InnerText = "Counts of " & field_to_count & " between " & interval_from & " and " & interval_to

                End Try

            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    ''' <summary>
    ''' sums the number of rows where each distinct item occur
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_group"></param>
    ''' <param name="field_to_sum"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_group = quater, field_to_sum = drink: first(75), second(155), third(95)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function DoughnutSum(table As String, field_to_group As String, field_to_sum As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildSumString_GROUPED(table, field_to_group, field_to_sum, keys_)
        Dim g As DataTable = QDataTable(query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Sums of " & field_to_sum

                Catch ex As Exception

                End Try

            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    Public Shared Function DoughnutSumFromExcel(table As String, field_to_group As String, field_to_sum As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildSumString_GROUPED_Excel(table, field_to_group, field_to_sum, where_keys, where_values)
        Dim g As DataTable = QDataTableFromExcel(query, connection_string)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Sums of " & field_to_sum

                Catch ex As Exception

                End Try

            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    ''' <summary>
    ''' sums the number of rows where each distinct item occur, but between a period specified by field_for_interval (from interval_from to interval_to)
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_sum"></param>
    ''' <param name="field_to_group"></param>
    ''' <param name="field_for_interval"></param>
    ''' <param name="interval_from"></param>
    ''' <param name="interval_to"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_sum = drink, field_to_group = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(25), second(155), third(65)
    ''' </example>
    ''' <returns></returns>

    Public Shared Function DoughnutSumAcrossInterval(table As String, field_to_sum As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query As String = BuildSumString_GROUPED_BETWEEN(table, field_to_group, field_to_sum, {field_for_interval})
        Dim g As DataTable = QDataTable(query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With
        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Sums of " & field_to_sum & " between " & Date.Parse(interval_from).ToShortDateString & " and " & Date.Parse(interval_to).ToShortDateString
                Catch ex As Exception
                    div_for_title.InnerText = "Sums of " & field_to_sum & " between " & interval_from & " and " & interval_to

                End Try
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    ''' <summary>
    ''' averages the number of rows where each distinct item occur
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_average"></param>
    ''' <param name="field_to_group"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_group = quater, field_to_average = drink: first(37), second(77), third(31)
    ''' </example>
    ''' <returns></returns>

    Public Shared Function DoughnutAverage(table As String, field_to_average As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildAVGString_GROUPED(table, field_to_group, field_to_average, keys_)
        Dim g As DataTable = QDataTable(query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Averages of " & field_to_average

                Catch ex As Exception

                End Try

            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    Public Shared Function DoughnutAverageFromExcel(table As String, field_to_average As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query = BuildAVGString_GROUPED_Excel(table, field_to_group, field_to_average, where_keys, where_values)
        Dim g As DataTable = QDataTableFromExcel(query, connection_string)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Averages of " & field_to_average

                Catch ex As Exception

                End Try

            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    ''' <summary>
    ''' averages the number of rows where each distinct item occur, but between a period specified by field_for_interval (from interval_from to interval_to)
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="field_to_average"></param>
    ''' <param name="field_to_group"></param>
    ''' <param name="field_for_interval"></param>
    ''' <param name="interval_from"></param>
    ''' <param name="interval_to"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="div_for_legend"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="chart_title"></param>
    ''' <param name="LegendMarkerStyle_"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	25	first
    ''' 3	75	second
    ''' 4	80	second
    ''' 5	45	third
    ''' 6	20	third
    ''' 7	30	third
    ''' field_to_average = drink, field_to_group = quater, field_for_interval = id, interval_from = 2, interval_to = 6 : first(25), second(77), third(32)
    ''' </example>
    ''' <returns></returns>
    Public Shared Function DoughnutAverageAcrossInterval(table As String, field_to_average As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim query As String = BuildAVGString_GROUPED_BETWEEN(table, field_to_group, field_to_average, {field_for_interval})
        Dim g As DataTable = QDataTable(query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Item(0).ToString)
                list_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With

        Dim colors As New List(Of String)
        With colors
            For i = 0 To g.Rows.Count - 1
                .Add(RandomColor(New List(Of Integer)))
            Next
        End With

        If div_for_legend IsNot Nothing Then
            Legend(list_labels, colors, div_for_legend, LegendMarkerStyle_)
        End If

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Averages of " & field_to_average & " between " & Date.Parse(interval_from).ToShortDateString & " and " & Date.Parse(interval_to).ToShortDateString

                Catch ex As Exception
                    div_for_title.InnerText = "Averages of " & field_to_average & " between " & interval_from & " and " & interval_to

                End Try

            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

#End Region

#Region "Line From Queries"
    ''' <summary>
    ''' Line chart, x axis is for labels, y axis is for values
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="x_label_field"></param>
    ''' <param name="y_value_field"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="x_label_field_caption"></param>
    ''' <param name="chart_title"></param>
    ''' <example>
    ''' id	drink	quater
    ''' 1	50	first
    ''' 2	75	second
    ''' 3	25	third
    ''' 4	45	fourth
    ''' 5	80	fifth
    ''' 6	30	sixth
    ''' 7	20	seventh
    ''' x_label_field: quater, y_value_field: drink, x_label_field_caption: what should be denoted as the 'topic/subject' in the title, e.g. the column name of y_value_field (in this case, "drink")
    ''' </example>
    ''' <returns></returns>

    Public Shared Function Line(table As String, x_label_field As String, y_value_field As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional y_values_field_name_caption As String = Nothing, Optional chart_title As String = Nothing)
        Dim x_label_field_caption As String = x_label_field
        Dim query As String = BuildSelectString(table, {x_label_field, y_value_field}, keys_)
        Dim g As DataTable = QDataTable(query, connection_string, keys_values)

        Dim l_values As New List(Of String), l_labels As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                l_labels.Add(.Rows(i).Item(0).ToString)
                l_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With
        Dim s = Chart_Line_String(l_values, l_labels)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s


        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    If y_values_field_name_caption IsNot Nothing Then
                        div_for_title.InnerText = "Progression of " & y_values_field_name_caption & " through " & x_label_field_caption
                    Else
                        div_for_title.InnerText = "Progression through " & x_label_field_caption
                    End If
                Catch ex As Exception
                    div_for_title.InnerText = ""
                End Try
            End If

        End If

        Return s
    End Function

    Public Shared Function LineFromExcel(table As String, x_label_field As String, y_value_field As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional y_values_field_name_caption As String = Nothing, Optional chart_title As String = Nothing)
        Dim x_label_field_caption As String = x_label_field
        Dim query As String = BuildSelectString_Excel(table, {x_label_field, y_value_field}, where_keys, where_values)
        Dim g As DataTable = QDataTableFromExcel(query, connection_string)

        Dim l_values As New List(Of String), l_labels As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                l_labels.Add(.Rows(i).Item(0).ToString)
                l_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With
        Dim s = Chart_Line_String(l_values, l_labels)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s


        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    If y_values_field_name_caption IsNot Nothing Then
                        div_for_title.InnerText = "Progression of " & y_values_field_name_caption & " through " & x_label_field_caption
                    Else
                        div_for_title.InnerText = "Progression through " & x_label_field_caption
                    End If
                Catch ex As Exception
                    div_for_title.InnerText = ""
                End Try
            End If

        End If

        Return s
    End Function

    ''' <summary>
    ''' WIP
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="x_label_field"></param>
    ''' <param name="y_value_field"></param>
    ''' <param name="field_for_interval"></param>
    ''' <param name="interval_from"></param>
    ''' <param name="interval_to"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="div_for_chart"></param>
    ''' <param name="keys_"></param>
    ''' <param name="keys_values"></param>
    ''' <param name="div_for_title"></param>
    ''' <param name="x_label_field_caption"></param>
    ''' <param name="chart_title"></param>
    ''' <returns></returns>
    Private Shared Function LineAcrossInterval(table As String, x_label_field As String, y_value_field As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional x_label_field_caption As String = Nothing, Optional chart_title As String = Nothing)
        Dim query As String = BuildSelectString_BETWEEN(table, {x_label_field, y_value_field})
        'Dim g As DataTable = QDataTable(query, connection_string, keys_values)
        Dim g As DataTable = QDataTable(query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})


        Dim l_values As New List(Of String), l_labels As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                l_labels.Add(.Rows(i).Item(0).ToString)
                l_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With
        Dim s = Chart_Line_String(l_values, l_labels)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s


        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Progression of " & x_label_field_caption & " between " & Date.Parse(interval_from).ToShortDateString & " and " & Date.Parse(interval_to).ToShortDateString
                Catch ex As Exception
                    div_for_title.InnerText = "Progression of " & x_label_field_caption & " between " & interval_from & " and " & interval_to
                End Try
            End If

        End If

        Return s

    End Function

    Public Shared Function Line(g As DataTable, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional x_label_field_caption As String = Nothing, Optional chart_title As String = Nothing)
        Dim l_values As New List(Of String), l_labels As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                l_labels.Add(.Rows(i).Item(0).ToString)
                l_values.Add(.Rows(i).Item(1).ToString)
            Next
        End With
        Dim s = Chart_Line_String(l_values, l_labels)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s

        If div_for_title IsNot Nothing Then
            If chart_title IsNot Nothing Then
                div_for_title.InnerText = chart_title
            Else
                Try
                    div_for_title.InnerText = "Progression of " & x_label_field_caption
                Catch ex As Exception
                    div_for_title.InnerText = ""
                End Try
            End If

        End If

        'If div_for_title IsNot Nothing And chart_title IsNot Nothing Then
        '    div_for_title.InnerText = chart_title
        'Else
        '    Try
        '        div_for_title.InnerText = "Progression of " & x_label_field_caption
        '    Catch ex As Exception
        '        div_for_title.InnerText = ""
        '    End Try
        'End If

        Return s
    End Function
#End Region

#Region "Members"
    Public Enum LegendMarkerStyle
        Circle
        Square
    End Enum
    Public Structure BarChartDataSet
        Public x_values_for_each_x_label As List(Of String)
        Public color_ As String
        Public legend_value As String
    End Structure

#End Region

#Region "Support Functions"
    Private Shared Function Legend(Labels As List(Of String), legend_colors As List(Of String), LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        Dim legend_ As String = ""
        With legend_colors
            For i As Integer = 0 To .Count - 1
                If LegendMarkerStyle_ = LegendMarkerStyle.Circle Then
                    legend_ &= "<span style=""border-radius:50%; background-color:" & legend_colors(i).Replace("""", "") & """>&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;&nbsp;" & Labels(i)
                ElseIf LegendMarkerStyle_ = LegendMarkerStyle.Square Then
                    legend_ &= "<span style=""border-radius:20%; background-color:" & legend_colors(i).Replace("""", "") & """>&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;&nbsp;" & Labels(i)
                End If
                legend_ &= "<br />"
            Next
        End With
        legend_ &= "</div>"
        If LegendControl IsNot Nothing Then WriteContent(legend_, LegendControl)
        Return legend_
    End Function

#End Region

End Class
