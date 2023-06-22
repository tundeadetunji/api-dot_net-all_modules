Imports Bank.Transactions
Imports Web_Module.DataConnectionWeb
Imports Web_Module.Functions
Imports General_Module.FormatWindow
Imports NModule.NFunctions
Imports NModule.W
Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Web.UI.WebControls
Imports Web_Module.DW
Imports System.Web.UI.HtmlControls

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
    'Private Shared Function DoughnutChart(grid_with_values As GridView, Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
    '	Dim g As GridView = grid_with_values
    '	Dim l As New List(Of String)
    '	With g
    '		If .Rows.Count < 1 Then Return ""
    '		For i = 0 To .Rows.Count - 1
    '			l.Add(g.Rows(i).Cells(0).Text)
    '		Next
    '	End With
    '	Dim s As String = Chart_Doughnut_String(l, COLORS, width, height)
    '	Return s
    'End Function
    'Private Shared Function DoughnutChart(grid_with_values As GridView, div As HtmlGenericControl, Optional COLORS As List(Of String) = Nothing, Optional width As Integer = 470, Optional height As Integer = 300) As String
    '	Dim g As GridView = grid_with_values
    '	Dim l As New List(Of String)
    '	With g
    '		For i = 0 To .Rows.Count - 1
    '			l.Add(g.Rows(i).Cells(0).Text)
    '		Next
    '	End With
    '	Dim s As String = Chart_Doughnut_String(l, COLORS, width, height)
    '	If div IsNot Nothing Then div.InnerHtml = s
    '	Return s
    'End Function

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

    'Public Shared Function Line(grid_with_labels_values As GridView, Optional width As Integer = 400, Optional height As Integer = 300) As String
    '    Dim l_values As New List(Of String), l_labels As New List(Of String)
    '    Dim g As GridView = grid_with_labels_values
    '    With g
    '        For i = 0 To .Rows.Count - 1
    '            l_values.Add(.Rows(i).Cells(1).Text)
    '            l_labels.Add(.Rows(i).Cells(0).Text)
    '        Next
    '    End With
    '    Return Chart_Line_String(l_values, l_labels, width, height)
    'End Function
    'Public Shared Function Line(grid_with_labels_values As GridView, div As HtmlGenericControl, Optional width As Integer = 400, Optional height As Integer = 300) As String
    '    Dim l_values As New List(Of String), l_labels As New List(Of String)
    '    Dim g As GridView = grid_with_labels_values
    '    With g
    '        For i = 0 To .Rows.Count - 1
    '            l_values.Add(.Rows(i).Cells(1).Text)
    '            l_labels.Add(.Rows(i).Cells(0).Text)
    '        Next
    '    End With
    '    Dim s = Chart_Line_String(l_values, l_labels, width, height)
    '    If div IsNot Nothing Then div.InnerHtml = s
    '    Return s
    'End Function

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

    ''' <summary>
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
    Public Shared Function PieCount(g As GridView, table As String, field_to_count As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildCountString_GROUPED(table, field_to_count, field_to_group, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Counts Of " & field_to_count
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieCountOverTime(g As GridView, table As String, field_to_count As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildCountString_GROUPED_BETWEEN(table, field_to_group, field_to_count, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Counts Of " & field_to_count & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieSum(g As GridView, table As String, field_to_group As String, field_to_sum As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildSumString_GROUPED(table, field_to_group, field_to_sum, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Sums Of " & field_to_sum
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieSumOverTime(g As GridView, table As String, field_to_sum As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query As String = BuildSumString_GROUPED_BETWEEN(table, field_to_group, field_to_sum, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Sums Of " & field_to_sum & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieAverage(g As GridView, table As String, field_to_apply_avg_on As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildAVGString_GROUPED(table, field_to_group, field_to_apply_avg_on, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Averages Of " & field_to_apply_avg_on
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Public Shared Function PieAverageOverTime(g As GridView, table As String, field_to_apply_avg_on As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query As String = BuildAVGString_GROUPED_BETWEEN(table, field_to_group, field_to_apply_avg_on, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Averages Of " & field_to_apply_avg_on & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
    Private Shared Function PieTop(g As GridView, table As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional rows_to_select As Long = 10, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.DESC, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildTopString_GROUPED(table, field_to_group, keys_, keys_values, rows_to_select, OrderByField, order_by)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Top " & rows_to_select & " " & field_to_group
            End If
        End If

        Return PieChart(list_values, list_labels, div_for_chart, div_for_legend, colors)
    End Function
#End Region

#Region "Doughnut From Queries"
    Public Shared Function DoughnutCount(g As GridView, table As String, field_to_count As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildCountString_GROUPED(table, field_to_count, field_to_group, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Counts Of " & field_to_count
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function
    Public Shared Function DoughnutCountOverTime(g As GridView, table As String, field_to_count As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildCountString_GROUPED_BETWEEN(table, field_to_group, field_to_count, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Counts Of " & field_to_count & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

    Public Shared Function DoughnutSum(g As GridView, table As String, field_to_group As String, field_to_sum As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildSumString_GROUPED(table, field_to_group, field_to_sum, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Sums Of " & field_to_sum
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function
    Public Shared Function DoughnutSumOverTime(g As GridView, table As String, field_to_sum As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query As String = BuildSumString_GROUPED_BETWEEN(table, field_to_group, field_to_sum, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Sums Of " & field_to_sum & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function
    Public Shared Function DoughnutAverage(g As GridView, table As String, field_to_apply_avg_on As String, field_to_group As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query = BuildAVGString_GROUPED(table, field_to_group, field_to_apply_avg_on, keys_)
        Display(g, query, connection_string, keys_values)
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Averages Of " & field_to_apply_avg_on
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function
    Public Shared Function DoughnutAverageOverTime(g As GridView, table As String, field_to_apply_avg_on As String, field_to_group As String, field_for_interval As String, interval_from As String, interval_to As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_legend As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing, Optional LegendMarkerStyle_ As LegendMarkerStyle = LegendMarkerStyle.Square)
        g.Visible = False
        Dim query As String = BuildAVGString_GROUPED_BETWEEN(table, field_to_group, field_to_apply_avg_on, {field_for_interval})
        Display(g, query, connection_string, {field_for_interval & "_From", interval_from, field_for_interval & "_To", interval_to})
        Dim list_labels As New List(Of String)
        Dim list_values As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                list_labels.Add(.Rows(i).Cells(0).Text)
                list_values.Add(.Rows(i).Cells(1).Text)
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
                div_for_title.InnerText = "Averages Of " & field_to_apply_avg_on & " Between " & Date.Parse(interval_from).ToShortDateString & " And " & Date.Parse(interval_to).ToShortDateString
            End If
        End If

        Return DoughnutChart(list_values, div_for_chart, div_for_legend, list_labels, colors)
    End Function

#End Region

#Region "Line From Queries"
    Public Shared Function Line(g As GridView, table As String, x_label_field As String, y_value_field As String, connection_string As String, Optional div_for_chart As HtmlGenericControl = Nothing, Optional keys_ As Array = Nothing, Optional keys_values As Array = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing)
        g.Visible = False
        Dim query As String = BuildSelectString(table, {x_label_field, y_value_field}, keys_)
        Display(g, query, connection_string, keys_values)

        Dim l_values As New List(Of String), l_labels As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                l_labels.Add(.Rows(i).Cells(0).Text)
                l_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With
        Dim s = Chart_Line_String(l_values, l_labels)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s

        If div_for_title IsNot Nothing And chart_title IsNot Nothing Then
            div_for_title.InnerText = chart_title
        End If

        Return s
    End Function

    Public Shared Function Line(g As GridView, Optional div_for_chart As HtmlGenericControl = Nothing, Optional div_for_title As HtmlGenericControl = Nothing, Optional chart_title As String = Nothing)
        g.Visible = False
        Dim l_values As New List(Of String), l_labels As New List(Of String)
        With g
            For i = 0 To .Rows.Count - 1
                l_labels.Add(.Rows(i).Cells(0).Text)
                l_values.Add(.Rows(i).Cells(1).Text)
            Next
        End With
        Dim s = Chart_Line_String(l_values, l_labels)
        If div_for_chart IsNot Nothing Then div_for_chart.InnerHtml = s
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
