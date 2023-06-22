
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;

class x
{
    public static string PieChart(List<string> l_values, List<string> l_colors = null, int width = 470, int height = 300)
    {
        return Chart_Pie_String(l_values, l_colors, width, height);
    }

    private static string Chart_Pie_String(List<string> l_values, List<string> COLORS = null, int width = 470, int height = 300)
    {
        var id = NewGUID();
        string s = "<canvas id=\"" + id + "\" width=\"" + width + "\" height=\"" + height + "\" style=\"width: " + width + "px; height: " + height + @"px""></canvas>
			<script>
				var pieData = [";
        s += Chart_Pie_Variable(l_values, COLORS);
        s += @"];
					new Chart(document.getElementById(""" + id + @""").getContext(""2d"")).Pie(pieData);
			  </script>";
        return s;
    }

    private static void Chart_Pie_Variable(List<string> l_values, List<string> COLORS)
    {
        List<string> l = l_values;
        string s = "";
        string color;

        for (var i = 0; i <= l.Count - 1; i++)
        {
            if (COLORS != null)
                color = COLORS[i];
            else
                color = RandomColor(new List<int>());
            s += @"{
						value: " + l[i] + @",
						color: """ + color + @"""
				  }";
            if (i != l.Count - 1)
                s += ",";
        }
        return s;
    }


    public static string DoughnutChart(List<string> l_values, List<string> COLORS = null, int width = 470, int height = 300)
    {
        return Chart_Doughnut_String(l_values, COLORS, width, height);
    }
    public static string DoughnutChart(List<string> l_values, HtmlGenericControl div, HtmlGenericControl legend = null/* TODO Change to default(_) if this is not a reference type */, List<string> l_labels = null, List<string> COLORS = null, int width = 470, int height = 300)
    {
        string s = Chart_Doughnut_String(l_values, COLORS, width, height);
        if (div != null)
            div.InnerHtml = s;
        return s;
    }

    private static string Chart_Doughnut_String(List<string> l_values, List<string> COLORS = null, int width = 470, int height = 300)
    {
        var id = NewGUID();
        string s = "<canvas id=\"" + id + "\" width=\"" + width + "\" height=\"" + height + "\" style=\"width: " + width + "px; height: " + height + @"px""></canvas>
			<script>
				var doughnutData = [";
        s += Chart_Doughnut_Variable(l_values, COLORS);
        s += @"];
					new Chart(document.getElementById(""" + id + @""").getContext(""2d"")).Doughnut(doughnutData);
			  </script>";
        return s;
    }
    private static void Chart_Doughnut_Variable(List<string> l_values, List<string> COLORS)
    {
        List<string> l = l_values;
        string s = "";
        string color;

        for (var i = 0; i <= l.Count - 1; i++)
        {
            if (COLORS != null)
                color = COLORS[i];
            else
                color = RandomColor(new List<int>());
            s += @"{
						value: " + l[i] + @",
						color: """ + color + @"""
				  }";
            if (i != l.Count - 1)
                s += ",";
        }
        return s;
    }
    public static string LineChart(List<string> l_values, List<string> l_labels, int width = 400, int height = 300)
    {
        return Chart_Line_String(l_values, l_labels, width, height);
    }
    public static string LineChart(List<string> l_values, List<string> l_labels, HtmlGenericControl div, int width = 400, int height = 300)
    {
        string s = Chart_Line_String(l_values, l_labels, width, height);
        div.InnerHtml = s;
        return s;
    }

    private static string Chart_Line_String(List<string> l_values, List<string> l_labels, int width = 400, int height = 300)
    {
        var id = NewGUID();
        string s = "<canvas id=\"" + id + "\" width=\"" + width + "\" height=\"" + height + "\" style=\"width: " + width + "px; height: " + height + @"px""></canvas>
			<script>
				var lineChartData = {";
        s += Chart_Line_Variable(l_values, l_labels);
        s += "new Chart(document.getElementById(\"" + id + @""").getContext(""2d"")).Line(lineChartData);
			  </script>";
        return s;
    }
    private static void Chart_Line_Variable(List<string> l_values, List<string> l_labels)
    {
        string l_string = "labels : [";
        for (var i = 0; i <= l_labels.Count - 1; i++)
        {
            l_string += "\"" + l_labels[i] + "\"";
            if (i != l_labels.Count - 1)
                l_string += ",";
        }
        var fc = RandomColor(new List<int>());
        var c = RandomColor(new List<int>());
        l_string += "],";

        string v_string = @"datasets: [
										{
											fillColor: """ + fc + @""",
											strokeColor: """ + c + @""",
											pointColor: """ + c + @""",
											pointStrokeColor: ""#fff"",
											data: [";
        for (var j = 0; j <= l_values.Count - 1; j++)
        {
            v_string += l_values[j];
            if (j != l_values.Count - 1)
                v_string += ",";
        }
        v_string += "]}]};";
        return l_string + v_string;
    }

    public static void BarChart(List<string> x_labels, List<BarChartDataSet> l_dataset, int width = 400, int height = 300)
    {
        return Chart_Bar_String(x_labels, l_dataset, width, height);
    }
    public static void BarChart(List<string> x_labels, List<BarChartDataSet> l_dataset, HtmlGenericControl div_for_chart, HtmlGenericControl div_for_legend = null/* TODO Change to default(_) if this is not a reference type */, HtmlGenericControl div_for_title = null/* TODO Change to default(_) if this is not a reference type */, string chart_title = null, LegendMarkerStyle LegendMarkerStyle_ = LegendMarkerStyle.Square, int width = 400, int height = 300)
    {
        var s = Chart_Bar_String(x_labels, l_dataset, width, height);

        List<string> colors = new List<string>();
        for (var i = 0; i <= l_dataset.Count - 1; i++)
            colors.Add(l_dataset[i].color_);

        if (div_for_legend != null)
        {
            List<string> legend_values = new List<string>();
            for (var i = 0; i <= l_dataset.Count - 1; i++)
                legend_values.Add(l_dataset[i].legend_value);

            Legend(legend_values, colors, div_for_legend, LegendMarkerStyle_);
        }

        if (div_for_title != null & chart_title != null)
            div_for_title.InnerText = chart_title;

        if (div_for_chart != null)
            div_for_chart.InnerHtml = s;
        return s;
    }

    private static string Chart_Bar_String(List<string> x_labels, List<BarChartDataSet> l_dataset, int width = 400, int height = 300)
    {
        var id = NewGUID();
        string s = "<canvas id=\"" + id + "\" width=\"" + width + "\" height=\"" + height + "\" style=\"width: " + width + "px; height: " + height + @"px""></canvas>
        <script>
        	var barChartData = {";
        s += Chart_Bar_Variable(x_labels, l_dataset);
        s += @"};
        		new Chart(document.getElementById(""" + id + @""").getContext(""2d"")).Bar(barChartData);
          </script>";

        return s;
    }
    private static void Chart_Bar_Variable(List<string> x_labels, List<BarChartDataSet> datasets)
    {
        string s = "labels: [";
        for (var i = 0; i <= x_labels.Count - 1; i++)
        {
            s += "\"" + x_labels[i] + "\"";
            if (i != x_labels.Count - 1)
                s += ",";
        }
        s += @"],
                            datasets: [";

        for (var i = 0; i <= datasets.Count - 1; i++)
        {
            s += "{fillColor: \"" + datasets[i].color_ + "\", data: [";
            for (var j = 0; j <= datasets[i].x_values_for_each_x_label.Count - 1; j++)
            {
                s += datasets[i].x_values_for_each_x_label(j);
                if (j != datasets[i].x_values_for_each_x_label.Count - 1)
                    s += ",";
            }
            s += "]}";
            if (i != datasets.Count - 1)
                s += ",";
        }
        s += "]";
        return s;
    }









}
