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
Public Class DBCharts
#Region "Fields"
	Private l_ri
	Private l_pie
	Private l_line
	Private list_of_values_PIE
	Private list_of_values_DOUGHNUT
	Private list_of_values_LINE
	Private legend_colors_PIE
	Private legend_colors_DOUGHNUT
	Private hasLegend_DOUGHNUT As Boolean
	Private hasLegend_PIE As Boolean

	Public Property legend_title_size__ As Byte = 3
	Public Property colors_control__ As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing
	Public Property colors_temp_list As New List(Of String)
	Public Property legend_marker_style__ As LegendMarkerStyle = LegendMarkerStyle.Circle
#End 

#Region "Enums"
	Public Enum RangeType
		DateInterval
		Numeric
	End Enum

	Public Enum ChartType
		Pie
		Doughnut
		Line
	End Enum
	Public Enum LegendMarkerStyle
		Circle
		Square
	End Enum

#End Region

#Region "New"
	Public Sub New(Optional legend_title_size As Byte = 3, Optional legend_marker_style As LegendMarkerStyle = LegendMarkerStyle.Circle)
		legend_marker_style__ = legend_marker_style

		list_of_values_PIE = New List(Of String)
		list_of_values_DOUGHNUT = New List(Of String)
		list_of_values_LINE = New List(Of Object)

		l_ri = New List(Of Integer)
		l_ri.Add(0)
		l_ri.Clear()

		l_pie = New List(Of String)
		l_pie.add(0)
		l_pie.clear

		l_line = New List(Of Object)
		l_line.add(0)
		l_line.clear

		legend_colors_DOUGHNUT = New List(Of String)
		legend_colors_DOUGHNUT.add(0)
		legend_colors_DOUGHNUT.clear

		legend_colors_PIE = New List(Of String)
		legend_colors_PIE.add(0)
		legend_colors_PIE.clear

		If legend_title_size.ToString.Length > 0 Then
			If legend_title_size < 1 Or legend_title_size > 7 Then
				legend_title_size__ = 7
			Else
				legend_title_size__ = legend_title_size
			End If
		End If
	End Sub
#End Region

#Region "Helper Functions"
	Public Function WriteColors(colors_ As Array) As String
		Dim r As String = ListToString(colors_)
		If colors_control__ IsNot Nothing Then Write(r, colors_control__)
		Return r
	End Function

	''' <summary>
	''' Use in place of LineLabels (pass same list into it and use it as arg instead) if the labels are to be treated as database column headers.
	''' </summary>
	''' <param name="select_keys"></param>
	''' <returns></returns>
	Public Function LineLabelsAsSelectKeys(select_keys As String()) As List(Of String)
		If select_keys IsNot Nothing Then
			Return ExtractHeader(select_keys)
		End If
	End Function

	Private Sub WriteLegend(Labels As List(Of String), legend_colors As List(Of String), chartType As ChartType, LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl, Optional title_ As String = Nothing, Optional slices_values As List(Of String) = Nothing, Optional showValues As Boolean = True)
		If chartType = ChartType.Doughnut Then
			legend_colors = legend_colors_DOUGHNUT
		ElseIf chartType = ChartType.Pie Then
			legend_colors = legend_colors_PIE
		End If
		Dim legend_ As String = ""
		If title_ <> Nothing Then legend_ &= "<h" & legend_title_size__ & ">" & title_ & "</h" & legend_title_size__ & "><br />"
		With legend_colors
			For i As Integer = 0 To .Count - 1
				If legend_marker_style__ = LegendMarkerStyle.Circle Then
					legend_ &= "<span style=""border-radius:50%; background-color:" & legend_colors(i).Replace("""", "") & """>&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;&nbsp;" & Labels(i)
				ElseIf legend_marker_style__ = LegendMarkerStyle.Square Then
					legend_ &= "<span style=""border-radius:20%; background-color:" & legend_colors(i).Replace("""", "") & """>&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;&nbsp;" & Labels(i)
				End If
				If showValues = True And slices_values IsNot Nothing Then legend_ &= " - " & slices_values(i)
				legend_ &= "<br />"
			Next
		End With
		legend_ &= "</div>"
		WriteContent(legend_, LegendControl)
	End Sub

	Private Function GetTitle(operationType As OperationType, labels_ As List(Of String), field_to_apply_function_on As String) As String
		Dim suffx As String = " " & labels_(0) '& " occurred"
		If labels_.Count > 1 Then suffx = " each of " & labels_(0) & ", " & labels_(1) ' & " etc. occurred"

		Dim suffx2 As String = ""
		If labels_.Count > 2 Then
			suffx2 = "..."
		End If

		Select Case operationType
			Case OperationType.AVG
				Return "Average of " & field_to_apply_function_on & " for" & suffx & suffx2
			Case OperationType.Count
				Dim r As String = "Number of times" & suffx
				If labels_.Count < 3 Then
					r &= " occurred"
				Else
					r &= " etc. occurred"
				End If
				Return r
			Case OperationType.MIN
				Return "Minimum value of " & field_to_apply_function_on & " for" & suffx & suffx2
			Case OperationType.MAX
				Return "Maximum value of " & field_to_apply_function_on & " for" & suffx & suffx2
			Case OperationType.Sum
				Return "Total of " & field_to_apply_function_on & " for" & suffx & suffx2
		End Select
	End Function

	Private Function GetTitle_Line(operationType As OperationType, field_to_apply_function_on As String, range_min As Object, range_max As Object) As String
		Dim r As String

		Select Case operationType
			Case OperationType.AVG
				r = "Averages of " & field_to_apply_function_on & " between " & NumOrDateToString(range_min) & " and " & NumOrDateToString(range_max)
			Case OperationType.Count
				r = "Number of occurrences of " & field_to_apply_function_on & " between " & NumOrDateToString(range_min) & " and " & NumOrDateToString(range_max)
			'Case OperationType.MIN
			'	r = "Minimum value of " & field_to_apply_function_on & " between " & NumOrDateToString(range_min) & " and " & NumOrDateToString(range_max)
			'Case OperationType.MAX
			'	r = "Maximum value of " & field_to_apply_function_on & " between " & NumOrDateToString(range_min) & " and " & NumOrDateToString(range_max)
			Case OperationType.Sum
				r = "Totals of " & field_to_apply_function_on & " between " & NumOrDateToString(range_min) & " and " & NumOrDateToString(range_max)
				'Case OperationType.None
				'	r = "Values of " & field_to_apply_function_on & " between " & NumOrDateToString(range_min) & " and " & NumOrDateToString(range_max)
		End Select

		Return "<h" & legend_title_size__ & ">" & r & "</h" & legend_title_size__ & "><br />"
	End Function
	Private Function GetTitle_Bank(RangeFrom As String, RangeTo As String) As String
		Dim r As String = "Total Volume between " & RangeFrom & " and " & RangeTo

		Return "<h" & legend_title_size__ & ">" & r & "</h" & legend_title_size__ & "><br />"
	End Function

	Private Function NumOrDateToString(num_or_date As Object) As String
		If TypeOf num_or_date Is Date Or TypeOf num_or_date Is DateTime Or TypeOf num_or_date Is Date Then
			Return Date.Parse(Date.Parse(num_or_date).ToShortDateString())
		Else
			Return num_or_date
		End If
		'Try
		'	Return Date.Parse(num_or_date).ToShortDateString
		'Catch ex As Exception
		'	Return num_or_date
		'End Try
	End Function
#End Region

#Region "Pie"
	Private Function PieChartRandomColor() As String
		Dim l As List(Of Integer) = l_ri
		Dim r As Integer = RandomList(0, 256, l)
		Dim g As Integer = RandomList(0, 256, l)
		Dim b As Integer = RandomList(0, 256, l)
		Dim return_ As String = "rgb(" & r & ", " & g & ", " & b & ")"
		Return return_
	End Function
	Public Structure PieChartData
		Public Labels As List(Of String)
		Public slices_values_ As List(Of String)
		Public sender_ As Object
		Public page_or_updatePanel As Object
		Public instance_of_script_manager As Object
		Public id_of_canvas As String
		Public LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl
		Public title_ As String
		Public ShowValuesInLegend As Boolean
	End Structure
	Public Sub PieChart(PieChartData_ As PieChartData)
		If PieChartData_.LegendControl IsNot Nothing And PieChartData_.Labels IsNot Nothing Then
			hasLegend_PIE = True
		Else
			hasLegend_PIE = False
		End If
		ConstructPieChart(PieChartData_.id_of_canvas, BuildPieChartData(PieChartData_.slices_values_, PieChartData_.Labels, PieChartData_.LegendControl, PieChartData_.ShowValuesInLegend, PieChartData_.title_), PieChartData_.sender_, PieChartData_.page_or_updatePanel, PieChartData_.instance_of_script_manager)
	End Sub
	Public Sub PieChart(slices_values_ As List(Of String), id_of_canvas As String, sender_ As Object, page_or_updatePanel As Object, instance_of_script_manager As Object, Optional Labels As List(Of String) = Nothing, Optional LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing, Optional ShowValuesInLegend As Boolean = True, Optional title_ As String = Nothing)

		Dim pieChartData As New PieChartData
		With pieChartData
			.slices_values_ = slices_values_
			.sender_ = sender_
			.page_or_updatePanel = page_or_updatePanel
			.instance_of_script_manager = instance_of_script_manager
			.id_of_canvas = id_of_canvas
			.Labels = Labels
			.LegendControl = LegendControl
			.title_ = title_
			.ShowValuesInLegend = ShowValuesInLegend
		End With
		PieChart(pieChartData)
	End Sub

	Private Function BuildPieChartData(slices_values_ As List(Of String), Labels As List(Of String), LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl, Optional ShowValuesInLegend As Boolean = True, Optional title_ As String = Nothing) As String

		Dim l As New List(Of String) '= values_colors
		With l
			For i As Integer = 0 To slices_values_.Count - 1
				l.Add(slices_values_(i))
				Dim color_ As String = PieChartRandomColor()
				l.Add(color_)
				colors_temp_list.Add(color_)
				If hasLegend_PIE Then legend_colors_PIE.add(color_)
			Next
		End With

		WriteColors(colors_temp_list.ToArray)

		If hasLegend_PIE Then
			WriteLegend(Labels, legend_colors_PIE, ChartType.Pie, LegendControl, title_, slices_values_, ShowValuesInLegend)
		End If

		Dim return_ As String = "["
		For i As Integer = 0 To l.Count - 1 Step 2
			return_ &= "{ value: " & l(i) & ", color: """ & l(i + 1) & """}"
			If i + 1 < l.Count - 1 Then return_ &= ","
		Next
		Return return_ & "]"
	End Function
	Private Sub ConstructPieChart(canvas_ As String, arg_ As Object, sender_ As Object, page_or_updatePanel As Object, instance_of_script_manager As Object, Optional title_control As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing, Optional title_string As String = "", Optional placeholder_for_ref_to_js_file As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing, Optional relative_path_to_js_file As String = Nothing)
		If title_control IsNot Nothing Then WriteContent(title_string, title_control)
		'		script_manager.RegisterStartupScript(control_, control_.GetType(), "text", "CallFunction('" & str_.Replace("'", "") & "')", True)
		Dim function_ As String = "PieChart"
		Dim param_
		If TypeOf (arg_) Is String Then
			param_ = arg_.ToString.Replace("'", "")
		Else
			param_ = arg_
		End If
		If placeholder_for_ref_to_js_file IsNot Nothing And relative_path_to_js_file IsNot Nothing Then
			Dim placeholder_text As String = "<script type=""text/javascript"" src=" & relative_path_to_js_file & "></script>"
			placeholder_for_ref_to_js_file.InnerHtml = placeholder_text
			placeholder_for_ref_to_js_file.Visible = True
		End If
		instance_of_script_manager.RegisterStartupScript(sender_, page_or_updatePanel.GetType(), "text", function_ + "('" + canvas_ + "', " + param_ + ")", True)
	End Sub

#End Region

#Region "Doughnut"
	Private Function DoughnutChartRandomColor() As String
		Dim l As List(Of Integer) = l_ri
		Dim r As Integer = RandomList(0, 256, l)
		Dim g As Integer = RandomList(0, 256, l)
		Dim b As Integer = RandomList(0, 256, l)
		Dim return_ As String = "rgb(" & r & ", " & g & ", " & b & ")"
		Return return_
	End Function
	Public Structure DoughnutChartData
		Public Labels As List(Of String)
		Public slices_values_ As List(Of String)
		Public sender_ As Object
		Public page_or_updatePanel As Object
		Public instance_of_script_manager As Object
		Public id_of_canvas As String
		Public LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl
		Public title_ As String
		Public ShowValuesInLegend As Boolean
	End Structure
	Public Sub DoughnutChart(DoughnutChartData_ As DoughnutChartData)

		If DoughnutChartData_.LegendControl IsNot Nothing And DoughnutChartData_.Labels IsNot Nothing Then
			hasLegend_DOUGHNUT = True
		Else
			hasLegend_DOUGHNUT = False
		End If

		ConstructDoughnutChart(DoughnutChartData_.id_of_canvas, BuildDoughnutChartData(DoughnutChartData_.slices_values_, DoughnutChartData_.Labels, DoughnutChartData_.LegendControl, DoughnutChartData_.ShowValuesInLegend, DoughnutChartData_.title_), DoughnutChartData_.sender_, DoughnutChartData_.page_or_updatePanel, DoughnutChartData_.instance_of_script_manager)
	End Sub
	Public Sub DoughnutChart(slices_values_ As List(Of String), id_of_canvas As String, sender_ As Object, page_or_updatePanel As Object, instance_of_script_manager As Object, Optional Labels As List(Of String) = Nothing, Optional LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing, Optional ShowValuesInLegend As Boolean = True, Optional title_ As String = Nothing)
		Dim doughnutChartData As New DoughnutChartData
		With doughnutChartData
			.slices_values_ = slices_values_
			.sender_ = sender_
			.page_or_updatePanel = page_or_updatePanel
			.instance_of_script_manager = instance_of_script_manager
			.id_of_canvas = id_of_canvas
			.Labels = Labels
			.LegendControl = LegendControl
			.title_ = title_
			.ShowValuesInLegend = ShowValuesInLegend
		End With
		DoughnutChart(doughnutChartData)
	End Sub

	Private Function BuildDoughnutChartData(slices_values_ As List(Of String), Labels As List(Of String), LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl, Optional ShowValuesInLegend As Boolean = True, Optional title_ As String = Nothing) As String

		Dim l As New List(Of String) ' = value_color
		With l
			For i As Integer = 0 To slices_values_.Count - 1
				l.Add(slices_values_(i))
				Dim color_ As String = DoughnutChartRandomColor()
				l.Add(color_)
				colors_temp_list.Add(color_)
				If hasLegend_DOUGHNUT Then legend_colors_DOUGHNUT.add(color_)
			Next
		End With

		WriteColors(colors_temp_list.ToArray)

		If hasLegend_DOUGHNUT Then
			WriteLegend(Labels, legend_colors_DOUGHNUT, ChartType.Doughnut, LegendControl, title_, slices_values_, ShowValuesInLegend)
		End If

		Dim return_ As String = "["
		For i As Integer = 0 To l.Count - 1 Step 2
			return_ &= "{ value: " & l(i) & ", color: """ & l(i + 1) & """}"
			If i + 1 < l.Count - 1 Then return_ &= ","
		Next
		Return return_ & "]"
	End Function
	Private Sub ConstructDoughnutChart(canvas_ As String, arg_ As Object, sender_ As Object, page_or_updatePanel As Object, instance_of_script_manager As Object, Optional title_control As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing, Optional title_string As String = "", Optional placeholder_for_ref_to_js_file As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing, Optional relative_path_to_js_file As String = Nothing)
		If title_control IsNot Nothing Then WriteContent(title_string, title_control)
		Dim function_ As String = "DoughnutChart" '		script_manager.RegisterStartupScript(control_, control_.GetType(), "text", "CallFunction('" & str_.Replace("'", "") & "')", True)
		Dim param_
		If TypeOf (arg_) Is String Then
			param_ = arg_.ToString.Replace("'", "")
		Else
			param_ = arg_
		End If
		If placeholder_for_ref_to_js_file IsNot Nothing And relative_path_to_js_file IsNot Nothing Then
			Dim placeholder_text As String = "<script type=""text/javascript"" src=" & relative_path_to_js_file & "></script>"
			placeholder_for_ref_to_js_file.InnerHtml = placeholder_text
			placeholder_for_ref_to_js_file.Visible = True
		End If
		instance_of_script_manager.RegisterStartupScript(sender_, page_or_updatePanel.GetType(), "text", function_ + "('" + canvas_ + "', " + param_ + ")", True)
	End Sub

#End Region

#Region "Line"


	'Private Function LineYCoordinates(YCoordinates As List(Of String)) As List(Of Object)
	'	Return LineData(YCoordinates)
	'End Function
	Private Function LineChartRandomColor() As String()
		Dim l As List(Of Integer) = l_ri
		Dim r As Integer = RandomList(0, 256, l)
		Dim g As Integer = RandomList(0, 256, l)
		Dim b As Integer = RandomList(0, 256, l)
		Dim s As Integer = Random_(1, 25)
		'		Dim s2 As Integer = Random_(0, 10)

		Dim r2 As Integer = r + s
		If r2 > 255 Then r2 = 255
		Dim g2 As Integer = g + s
		If g2 > 255 Then g2 = 255
		Dim b2 As Integer = b + s
		If b2 > 255 Then b2 = 255

		'Dim r3 As Integer = r2 + s2
		'If r3 > 255 Then r3 = 255
		'Dim g3 As Integer = g2 + s2
		'If g3 > 255 Then g3 = 255
		'Dim b3 As Integer = b2 + s2
		'If b3 > 255 Then b3 = 255

		Dim index__ = 1 '0.74
		'If index_ = 1 Then
		'	index__ = 0.44
		'Else
		'	index__ = 0.44
		'End If

		Dim return_ As String = "rgba(" & r & ", " & g & ", " & b & ", " & index__ & ")"
		Dim return_2 As String = "rgb(" & r2 & ", " & g2 & ", " & b2 & ")"
		Dim return_3 As String = "#fff"
		WriteColors({return_, return_2, return_3})
		Return {return_, return_2, return_3}
	End Function
	Public Structure LineChartData
		Public Labels As List(Of String)
		Public YCoordinates As List(Of String)
		Public sender_ As Object
		Public page_or_updatePanel As Object
		Public instance_of_script_manager As Object
		Public id_of_canvas As String
	End Structure
	Public Sub LineChart(LineChartData_ As LineChartData)
		ConstructLineChart(LineChartData_.id_of_canvas, BuildLineChartData(LineChartData_.Labels, LineData(LineChartData_.YCoordinates)), LineChartData_.sender_, LineChartData_.page_or_updatePanel, LineChartData_.instance_of_script_manager)
	End Sub

	Public Sub LineChart(Labels As List(Of String), YCoordinates As List(Of String), id_of_canvas As String, sender_ As Object, page_or_updatePanel As Object, instance_of_script_manager As Object)
		Dim lineChartData As New LineChartData
		With lineChartData
			.Labels = Labels
			.YCoordinates = YCoordinates
			.sender_ = sender_
			.page_or_updatePanel = page_or_updatePanel
			.instance_of_script_manager = instance_of_script_manager
			.id_of_canvas = id_of_canvas
		End With
		LineChart(lineChartData)
	End Sub

	Private Function BuildLineChartData(labels_ As List(Of String), datasets_fillColor_strokeColor_pointStrokeColor_data As List(Of Object)) As String
		Dim data_ As List(Of Object) = datasets_fillColor_strokeColor_pointStrokeColor_data
		Dim result As String = "{labels: ["
		For i As Integer = 0 To labels_.Count - 1
			result &= """" & labels_(i) & """"
			If i < labels_.Count - 1 Then
				result &= ", "
			End If
		Next
		result &= "], datasets: ["
		For j As Integer = 0 To data_.Count - 1 Step 4
			result &= "{fillColor: """ & data_(j) & """, strokeColor: """ & data_(j + 1) & """, pointColor: """ & data_(j + 1) & """, pointStrokeColor: """ & data_(j + 2) & """, data: ["
			Dim vals As List(Of String) = data_(j + 3)
			For k As Integer = 0 To vals.Count - 1
				result &= vals(k)
				If k < vals.Count - 1 Then
					result &= ", "
				End If
			Next
			result &= "]}"
			If j + 3 < data_.Count - 1 Then
				result &= ", "
			End If
		Next
		Return result & "]}"
	End Function
	Private Function LineData(YCoordinates As List(Of String)) As List(Of Object)
		'		Dim list_of_all_values As New List(Of Object)
		'		Dim list_of_all_values As New List(Of Object) ' = list_of_values_LINE
		list_of_values_LINE = New List(Of Object)

		Dim l_c As String() = LineChartRandomColor()
		list_of_values_LINE.Add(l_c(0))
		list_of_values_LINE.Add(l_c(1))
		list_of_values_LINE.Add(l_c(2))

		Dim list_of_dataset_values As New List(Of String)
		For i As Integer = 0 To YCoordinates.Count - 1
			list_of_dataset_values.Add(YCoordinates(i))
		Next
		list_of_values_LINE.Add(list_of_dataset_values)
		Return list_of_values_LINE
	End Function

	Private Sub ConstructLineChart(canvas_ As String, arg_ As Object, sender_ As Object, page_or_updatePanel As Object, instance_of_script_manager As Object, Optional placeholder_for_ref_to_js_file As System.Web.UI.HtmlControls.HtmlGenericControl = Nothing, Optional relative_path_to_js_file As String = Nothing)
		'		script_manager.RegisterStartupScript(control_, control_.GetType(), "text", "CallFunction('" & str_.Replace("'", "") & "')", True)
		Dim function_ As String = "LineChart"
		Dim param_
		If TypeOf (arg_) Is String Then
			param_ = arg_.ToString.Replace("'", "")
		Else
			param_ = arg_
		End If
		If placeholder_for_ref_to_js_file IsNot Nothing And relative_path_to_js_file IsNot Nothing Then
			Dim placeholder_text As String = "<script type=""text/javascript"" src=" & relative_path_to_js_file & "></script>"
			placeholder_for_ref_to_js_file.InnerHtml = placeholder_text
			placeholder_for_ref_to_js_file.Visible = True
		End If
		instance_of_script_manager.RegisterStartupScript(sender_, page_or_updatePanel.GetType(), "text", function_ + "('" + canvas_ + "', " + param_ + ")", True)
	End Sub

#End Region

#Region "From Bank"
	Public Structure LineChartDataFromBank
		Public key As String
		Public RangeFrom As String
		Public RangeTo As String
		Public sender As Object
		Public page_or_updatePanel As Object
		Public instance_of_script_manager As Object
		Public id_of_canvas As String
		Public TitleControl As System.Web.UI.HtmlControls.HtmlGenericControl
		Public title As String
		Public use_default_title As Boolean

	End Structure
	Public Async Sub LineChartFromBank(chart_data As LineChartDataFromBank)
		Dim labels_ As New List(Of String)
		Dim ycoords As New List(Of String)
		Try
			For i As Integer = chart_data.RangeFrom To chart_data.RangeTo
				labels_.Add(i)
				ycoords.Add(VolumeTotal(chart_data.key, i, i).Result)
			Next

			Dim cd As New LineChartData
			Dim chart_title As String = chart_data.title
			If chart_data.use_default_title Then chart_title = GetTitle_Bank(chart_data.RangeFrom, chart_data.RangeTo)
			If chart_data.TitleControl IsNot Nothing Then WriteContent(chart_title, chart_data.TitleControl)

			With cd
				.Labels = labels_
				.YCoordinates = ycoords
				.sender_ = chart_data.sender
				.page_or_updatePanel = chart_data.page_or_updatePanel
				.instance_of_script_manager = chart_data.instance_of_script_manager
				.id_of_canvas = chart_data.id_of_canvas
			End With
			LineChart(cd)
		Catch ex As Exception

		End Try

	End Sub

#End Region

#Region "From DB"

	Public Structure PieChartDataFromDB
		Public table As String
		Public where_keys As String()
		Public field_to_apply_function_on As String
		Public field_to_group As String
		Public grid As GridView
		Public connection_string As String
		Public select_parameter_keys_values_ As String()
		Public sender As Object
		Public page_or_updatePanel As Object
		Public instance_of_script_manager As Object
		Public id_of_canvas As String
		Public LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl
		Public title As String
		Public use_default_title As Boolean
		Public ShowValuesInLegend As Boolean
		Public operationType As OperationType
	End Structure
	Public Structure DoughnutChartDataFromDB
		Public table As String
		Public where_keys As String()
		Public field_to_apply_function_on As String
		Public field_to_group As String
		Public grid As GridView
		Public connection_string As String
		Public select_parameter_keys_values_ As String()
		Public sender As Object
		Public page_or_updatePanel As Object
		Public instance_of_script_manager As Object
		Public id_of_canvas As String
		Public LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl
		Public title As String
		Public use_default_title As Boolean
		Public ShowValuesInLegend As Boolean
		Public operationType As OperationType
	End Structure
	Public Structure LineChartDataFromDB
		Public RangeField As String

		Public RangeFrom As Object
		Public RangeTo As Object
		Public RangeType_ As RangeType
		Public table As String

		'		Public where_keys As String()
		Public field_to_apply_function_on As String
		Public field_to_group As String
		Public grid As GridView
		Public connection_string As String
		'Public select_parameter_keys_values_ As String()
		Public sender As Object
		Public page_or_updatePanel As Object
		Public instance_of_script_manager As Object
		Public id_of_canvas As String
		Public TitleControl As System.Web.UI.HtmlControls.HtmlGenericControl
		Public title As String
		Public use_default_title As Boolean
		Public operationType As OperationType

	End Structure
	Public Sub LineChartFromDB(chart_data As LineChartDataFromDB)
		Dim labels_ As New List(Of String)
		Dim ycoords As New List(Of String)
		'Dim where_keys As String() = 
		Dim query_ As String, title_ As String
		Select Case chart_data.operationType
			Case OperationType.AVG
				query_ = BuildAVGString_GROUPED_BETWEEN(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, {chart_data.RangeField})
			Case OperationType.Sum
				query_ = BuildSumString_GROUPED_BETWEEN(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, {chart_data.RangeField})
			Case OperationType.Count
				query_ = BuildCountString_GROUPED_BETWEEN(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, {chart_data.RangeField})
		End Select

		Dim where_keys_vals As String()
		'		If TypeOf chart_data.RangeFrom Is Date Or TypeOf chart_data.RangeFrom Is DateTime Or IsDate(Date.Parse(chart_data.RangeFrom)) Then
		'If IsNumeric(chart_data.RangeFrom) = False Or IsDate(chart_data.RangeFrom) Then ' IsDate(Date.Parse(chart_data.RangeFrom)) And IsNumeric(chart_data.RangeFrom) = False Then
		'	where_keys_vals = {chart_data.RangeField & "_FROM", DateTimeToSQL(chart_data.RangeFrom), chart_data.RangeField & "_TO", DateTimeToSQL(chart_data.RangeTo)}
		'Else
		'	where_keys_vals = {chart_data.RangeField & "_FROM", chart_data.RangeFrom, chart_data.RangeField & "_TO", chart_data.RangeTo}
		'End If
		If chart_data.RangeType_ = RangeType.DateInterval Then
			where_keys_vals = {chart_data.RangeField & "_FROM", DateTimeToSQL(chart_data.RangeFrom), chart_data.RangeField & "_TO", DateTimeToSQL(chart_data.RangeTo)}
		Else
			where_keys_vals = {chart_data.RangeField & "_FROM", chart_data.RangeFrom, chart_data.RangeField & "_TO", chart_data.RangeTo}
		End If
		'		Write(ListToString(where_keys_vals), chart_data.TitleControl)
		'		Exit Sub

		Try
			Display(chart_data.grid, query_, chart_data.connection_string, where_keys_vals)
			With chart_data.grid
				If .Rows.Count > 0 Then
					.Visible = False
					For i As Integer = 0 To .Rows.Count - 1
						If chart_data.RangeType_ = RangeType.DateInterval Then
							labels_.Add(DateToShort(Date.Parse(.Rows(i).Cells(0).Text)))
						Else
							labels_.Add(.Rows(i).Cells(0).Text)
						End If
						ycoords.Add(.Rows(i).Cells(1).Text)
					Next
				End If
			End With
			'NModule.W.Write(ListToString(labels_), chart_data.TitleControl)
			'Exit Sub


			'convert labels
			'Dim temp_list As New List(Of String)
			'With labels_
			'	For i As Integer = 0 To .Count - 1
			'		temp_list.Add(labels_(i))
			'	Next
			'End With
			'labels_.Clear()
			'With temp_list
			'	For i As Integer = 0 To .Count - 1
			'		labels_.Add(NumOrDateToString(temp_list(i)))
			'	Next
			'End With
			'temp_list.Clear()

			Dim cd As New LineChartData
			Dim chart_title As String = chart_data.title
			If chart_data.use_default_title Then chart_title = GetTitle_Line(chart_data.operationType, chart_data.field_to_apply_function_on, chart_data.RangeFrom, chart_data.RangeTo)
			If chart_data.TitleControl IsNot Nothing Then WriteContent(chart_title, chart_data.TitleControl)

			With cd
				.Labels = labels_
				.YCoordinates = ycoords
				.sender_ = chart_data.sender
				.page_or_updatePanel = chart_data.page_or_updatePanel
				.instance_of_script_manager = chart_data.instance_of_script_manager
				.id_of_canvas = chart_data.id_of_canvas
			End With
			LineChart(cd)
		Catch ex As Exception
			Write(ex.ToString, chart_data.TitleControl)
		End Try

	End Sub
	Public Sub ChartFromDB(Pie_OR_Doughnut_OR_Line_ChartDataFromDB As Object)
		If TypeOf Pie_OR_Doughnut_OR_Line_ChartDataFromDB Is PieChartDataFromDB Then
			PieChartFromDB(Pie_OR_Doughnut_OR_Line_ChartDataFromDB)
		ElseIf TypeOf Pie_OR_Doughnut_OR_Line_ChartDataFromDB Is DoughnutChartDataFromDB Then '
			DoughnutChartFromDB(Pie_OR_Doughnut_OR_Line_ChartDataFromDB)
		Else
			LineChartFromDB(Pie_OR_Doughnut_OR_Line_ChartDataFromDB)
		End If
	End Sub

	Public Sub PieChartFromDB(table As String, where_keys As String(), field_to_apply_function_on As String, field_to_group As String, connection_string As String, select_parameter_keys_values_ As String(), operationType As OperationType, grid As GridView, id_of_canvas As String, sender As Object, page_or_updatePanel As Object, instance_of_script_manager As Object, chartType As ChartType, LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl, Optional ShowValuesInLegend As Boolean = True, Optional title As String = Nothing, Optional use_default_title As Boolean = True)
		Dim chart_data As New PieChartDataFromDB
		With chart_data
			.table = table
			.where_keys = where_keys
			.field_to_apply_function_on = field_to_apply_function_on
			.field_to_group = field_to_group
			.grid = grid
			.connection_string = connection_string
			.select_parameter_keys_values_ = select_parameter_keys_values_
			.sender = sender
			.page_or_updatePanel = page_or_updatePanel
			.instance_of_script_manager = instance_of_script_manager
			.id_of_canvas = id_of_canvas
			'			.chartType = chartType
			.LegendControl = LegendControl
			.title = title
			.use_default_title = use_default_title
			.ShowValuesInLegend = ShowValuesInLegend
			.operationType = operationType
		End With
		PieChartFromDB(chart_data)
	End Sub
	Public Sub PieChartFromDB(chart_data As PieChartDataFromDB)
		Dim query_ As String, title_ As String
		Select Case chart_data.operationType
			Case OperationType.Count
				query_ = BuildCountString_GROUPED(chart_data.table, chart_data.field_to_apply_function_on, chart_data.where_keys)
			Case OperationType.AVG
				query_ = BuildAVGString_GROUPED(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, chart_data.where_keys)
			Case OperationType.Sum
				query_ = BuildSumString_GROUPED(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, chart_data.where_keys)
			Case OperationType.MIN
				query_ = BuildMinString_GROUPED(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, chart_data.where_keys)
			Case OperationType.MAX
				query_ = BuildMaxString_GROUPED(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, chart_data.where_keys)
		End Select

		Try
			Display(chart_data.grid, query_, chart_data.connection_string, chart_data.select_parameter_keys_values_)
			Dim l As New List(Of String)
			Dim labels_ As New List(Of String)
			With chart_data.grid
				If .Rows.Count > 0 Then
					.Visible = False
					For i As Integer = 0 To .Rows.Count - 1
						l.Add(.Rows(i).Cells(1).Text)
						labels_.Add(.Rows(i).Cells(0).Text)
					Next
				End If
			End With

			Dim pcd As New PieChartData ', dcd As New DoughnutChartData
			Dim chart_title As String = chart_data.title
			If chart_data.use_default_title Then chart_title = GetTitle(chart_data.operationType, labels_, chart_data.field_to_apply_function_on)

			'If chart_data.chartType = ChartType.Doughnut Then
			'	With dcd
			'		.slices_values_ = l
			'		.sender_ = chart_data.sender
			'		.page_or_updatePanel = chart_data.page_or_updatePanel
			'		.instance_of_script_manager = chart_data.instance_of_script_manager
			'		.id_of_canvas = chart_data.id_of_canvas
			'		.Labels = labels_
			'		.LegendControl = chart_data.LegendControl
			'		.title_ = chart_title
			'		.ShowValuesInLegend = chart_data.ShowValuesInLegend
			'	End With
			'	DoughnutChart(dcd)
			'Else
			With pcd
				.slices_values_ = l
				.sender_ = chart_data.sender
				.page_or_updatePanel = chart_data.page_or_updatePanel
				.instance_of_script_manager = chart_data.instance_of_script_manager
				.id_of_canvas = chart_data.id_of_canvas
				.Labels = labels_
				.LegendControl = chart_data.LegendControl
				.title_ = chart_title
				.ShowValuesInLegend = chart_data.ShowValuesInLegend
			End With
			PieChart(pcd)
			'End If
		Catch ex As Exception

		End Try

	End Sub

	Public Sub DoughnutChartFromDB(table As String, where_keys As String(), field_to_apply_function_on As String, field_to_group As String, connection_string As String, select_parameter_keys_values_ As String(), operationType As OperationType, grid As GridView, id_of_canvas As String, sender As Object, page_or_updatePanel As Object, instance_of_script_manager As Object, chartType As ChartType, LegendControl As System.Web.UI.HtmlControls.HtmlGenericControl, Optional ShowValuesInLegend As Boolean = True, Optional title As String = Nothing, Optional use_default_title As Boolean = True)
		Dim chart_data As New DoughnutChartDataFromDB
		With chart_data
			.table = table
			.where_keys = where_keys
			.field_to_apply_function_on = field_to_apply_function_on
			.field_to_group = field_to_group
			.grid = grid
			.connection_string = connection_string
			.select_parameter_keys_values_ = select_parameter_keys_values_
			.sender = sender
			.page_or_updatePanel = page_or_updatePanel
			.instance_of_script_manager = instance_of_script_manager
			.id_of_canvas = id_of_canvas
			'.chartType = chartType
			.LegendControl = LegendControl
			.title = title
			.use_default_title = use_default_title
			.ShowValuesInLegend = ShowValuesInLegend
			.operationType = operationType
		End With
		DoughnutChartFromDB(chart_data)
	End Sub

	Public Sub DoughnutChartFromDB(chart_data As DoughnutChartDataFromDB)
		Dim query_ As String, title_ As String
		Select Case chart_data.operationType
			Case OperationType.Count
				query_ = BuildCountString_GROUPED(chart_data.table, chart_data.field_to_apply_function_on, chart_data.where_keys)
			Case OperationType.AVG
				query_ = BuildAVGString_GROUPED(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, chart_data.where_keys)
			Case OperationType.Sum
				query_ = BuildSumString_GROUPED(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, chart_data.where_keys)
			Case OperationType.MIN
				query_ = BuildMinString_GROUPED(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, chart_data.where_keys)
			Case OperationType.MAX
				query_ = BuildMaxString_GROUPED(chart_data.table, chart_data.field_to_group, chart_data.field_to_apply_function_on, chart_data.where_keys)
		End Select

		Try
			Display(chart_data.grid, query_, chart_data.connection_string, chart_data.select_parameter_keys_values_)
			Dim l As New List(Of String)
			Dim labels_ As New List(Of String)
			With chart_data.grid
				If .Rows.Count > 0 Then
					.Visible = False
					For i As Integer = 0 To .Rows.Count - 1
						l.Add(.Rows(i).Cells(1).Text)
						labels_.Add(.Rows(i).Cells(0).Text)
					Next
				End If
			End With

			Dim dcd As New DoughnutChartData
			Dim chart_title As String = chart_data.title
			If chart_data.use_default_title Then chart_title = GetTitle(chart_data.operationType, labels_, chart_data.field_to_apply_function_on)

			'			If chart_data.chartType = ChartType.Doughnut Then
			With dcd
				.slices_values_ = l
				.sender_ = chart_data.sender
				.page_or_updatePanel = chart_data.page_or_updatePanel
				.instance_of_script_manager = chart_data.instance_of_script_manager
				.id_of_canvas = chart_data.id_of_canvas
				.Labels = labels_
				.LegendControl = chart_data.LegendControl
				.title_ = chart_title
				.ShowValuesInLegend = chart_data.ShowValuesInLegend
			End With
			DoughnutChart(dcd)
			'Else
			'	With pcd
			'		.slices_values_ = l
			'		.sender_ = chart_data.sender
			'		.page_or_updatePanel = chart_data.page_or_updatePanel
			'		.instance_of_script_manager = chart_data.instance_of_script_manager
			'		.id_of_canvas = chart_data.id_of_canvas
			'		.Labels = labels_
			'		.LegendControl = chart_data.LegendControl
			'		.title_ = chart_title
			'		.ShowValuesInLegend = chart_data.ShowValuesInLegend
			'	End With
			'	PieChart(pcd)
			'End If
		Catch ex As Exception

		End Try

	End Sub

#End Region

	'-------------------
	'From Scratch
	'-------------------

#Region "Enums"
	Public Enum LegendMarkerType
		Circle
		Square
	End Enum

	Public Enum TitleSize
		Default_
		Large
		Gigantic
	End Enum

	Public Enum ChartAs
		Pie
		Doughnut
		Line
	End Enum

	Public Enum StatIs
		Count
		'MIN
		'MAX
		Sum
		Average
	End Enum

#End Region

#Region "Helper Functions"
	Private Sub WriteLabels()
		'		Write("Chart", chart_type_drop_label__)
		Write("Statistic", operation_type_drop_label__)
		Write("Statistic On", functions_drop_label__)
		'		Write("Time Range", range_drop_label__)
		Write("From", range_from_text_label__)
		Write("To", range_to_text_label__)
		'		Write("Title Size", title_text_size_drop_label__)
		Write("Notify When Statistic Reaches", reminder_text_label__)
		Write("Set", reminder_button__)
	End Sub

	Public Function OperationType__(text_from_drop As String) As OperationType
		Select Case text_from_drop
			Case StatIs.Average.ToString
				Return OperationType.AVG
			Case StatIs.Count.ToString
				Return OperationType.Count
			Case StatIs.Sum.ToString
				Return OperationType.Sum
		End Select
	End Function

	Private Function title_size__(text_from_drop)
		Select Case text_from_drop
			Case TitleSize.Default_.ToString
				Return 3
			Case TitleSize.Large.ToString
				Return 2
			Case TitleSize.Gigantic.ToString
				Return 1
		End Select
	End Function
#End Region

#Region "Fields"
	Private Property operation_type_drop__ As DropDownList
	Private Property functions_drop__ As DropDownList
	Private Property range_from_text__ As TextBox
	Private Property range_to_text__ As TextBox
	Private Property operation_type_drop_label__ As Object
	Private Property functions_drop_label__ As Object
	Private Property range_from_text_label__ As Object
	Private Property range_to_text_label__ As Object
	Private Property retrieve_button__ As Button
	Private Property reminder_text_label__ As Object
	Private Property reminder_button__ As Button
	Private Property reminder_text__ As TextBox
	Private Property id__ As String
	Private Property con_string__ As String

#End Region
#Region "Main"
	''' <summary>
	''' Prepares to build chart from scratch. Format of FieldsOfTables is ? from ?, e.g. Field from Table.
	''' </summary>
	''' <param name="con_string"></param>
	''' <param name="id_"></param>
	''' <param name="OperationTypeDrop"></param>
	''' <param name="OperationTypeDropLabel"></param>
	''' <param name="FunctionFieldsDrop"></param>
	''' <param name="FunctionFieldsDropLabel"></param>
	''' <param name="RangeFromTextBox"></param>
	''' <param name="RangeFromTextBoxLabel"></param>
	''' <param name="RangeToTextBox"></param>
	''' <param name="RangeToTextBoxLabel"></param>
	''' <param name="RetrieveButton"></param>
	''' <param name="ReminderTextBox"></param>
	''' <param name="ReminderTextBoxLabel"></param>
	''' <param name="ReminderButton"></param>
	Public Sub ChartInitializer(con_string As String, id_ As String, OperationTypeDrop As DropDownList, OperationTypeDropLabel As Object, FunctionFieldsDrop As DropDownList, FunctionFieldsDropLabel As Object, RangeFromTextBox As TextBox, RangeFromTextBoxLabel As Object, RangeToTextBox As TextBox, RangeToTextBoxLabel As Object, RetrieveButton As Button, ReminderTextBox As TextBox, ReminderTextBoxLabel As Object, ReminderButton As Button)
		con_string__ = con_string
		operation_type_drop__ = OperationTypeDrop
		functions_drop__ = FunctionFieldsDrop
		range_from_text__ = RangeFromTextBox
		range_to_text__ = RangeToTextBox
		operation_type_drop_label__ = OperationTypeDropLabel
		functions_drop_label__ = FunctionFieldsDropLabel
		range_from_text_label__ = RangeFromTextBoxLabel
		range_to_text_label__ = RangeToTextBoxLabel
		retrieve_button__ = RetrieveButton
		reminder_text_label__ = ReminderTextBoxLabel
		reminder_button__ = ReminderButton
		reminder_text__ = ReminderTextBox
		id__ = id_
		InitializeControls(con_string, id__)
	End Sub

	Private Sub InitializeControls(con_string As String, id_ As String)
		'reminder_text__
		'reminder_text__.TextMode = TextBoxMode.Number

		'retrieve_button__
		retrieve_button__.Text = "Chart"
		'chart type
		'With chart_type_drop__
		'	With .Items
		'		If .Count < 1 Then
		'			.Add(ChartAs.Line.ToString)
		'			'.Add(ChartAs.Doughnut.ToString)
		'			'.Add(ChartAs.Pie.ToString)
		'		End If
		'	End With
		'End With
		'time range
		'		DData(range_drop__, BuildSelectString_DISTINCT("Chart", {"FieldFriendlyName"}, {"Category"}), con_string, "FieldFriendlyName", "FieldFriendlyName", {"Category", "Range"})

		'function
		DData(functions_drop__, BuildSelectString_DISTINCT("Chart", {"FieldFriendlyName"}, {"Category", "ID"}), con_string, "FieldFriendlyName", "FieldFriendlyName", {"Category", "Function", "ID", id_})

		'operation type
		With operation_type_drop__
			With .Items
				If .Count < 1 Then
					.Add(StatIs.Sum.ToString)
					.Add(StatIs.Average.ToString)
					.Add(StatIs.Count.ToString)
				End If
			End With
		End With
		'title
		'With title_text_size_drop__
		'	With .Items
		'		If .Count < 1 Then
		'			.Add(TitleSize.Default_.ToString.Replace("_", ""))
		'			.Add(TitleSize.Large.ToString)
		'			.Add(TitleSize.Gigantic.ToString)
		'		End If
		'	End With
		'End With
		'range from
		range_from_text__.TextMode = TextBoxMode.Date
		'range to
		range_to_text__.TextMode = TextBoxMode.Date
		'labels
		WriteLabels()
	End Sub

	Public Function Chart(id_ As String, con_string__ As String, functions_drop__ As DropDownList, operation_type_drop__ As DropDownList, range_from_text__ As TextBox, range_to_text__ As TextBox, id_of_canvas As String, g As GridView, sender As Object, page_OR_updatePanel As Object, instance_of_script_manager As Object, TitleControl As System.Web.UI.HtmlControls.HtmlGenericControl) As Boolean

		If range_from_text__.Text.Trim.Length < 1 Or range_to_text__.Text.Trim.Length < 1 Or NModule.W.Content(functions_drop__).Length < 1 Then Return False
		Dim table_ As String = QData(BuildSelectString("Chart", {"TableName"}, {"FieldFriendlyName", "ID"}), con_string__, {"FieldFriendlyName", NModule.W.Content(functions_drop__), "ID", id_})
		Dim group_ As String = QData(BuildSelectString("Chart", {"FieldName"}, {"TableName", "Category", "ID"}), con_string__, {"TableName", table_, "Category", "Range", "ID", id_})
		Dim cd As New LineChartDataFromDB
		With cd
			.connection_string = con_string__
			.field_to_apply_function_on = QData(BuildSelectString("Chart", {"FieldName"}, {"FieldFriendlyName", "ID"}), con_string__, {"FieldFriendlyName", NModule.W.Content(functions_drop__), "ID", id_})
			.field_to_group = group_
			.grid = g
			.id_of_canvas = id_of_canvas
			.instance_of_script_manager = instance_of_script_manager
			.operationType = OperationType__(NModule.W.Content(operation_type_drop__))
			.page_or_updatePanel = page_OR_updatePanel
			.RangeField = group_
			.RangeFrom = Date.Parse(NModule.W.Content(range_from_text__))
			.RangeTo = Date.Parse(NModule.W.Content(range_to_text__))
			.RangeType_ = RangeType.DateInterval
			.sender = sender
			.table = table_
			.title = ""
			.TitleControl = TitleControl
			.use_default_title = True
		End With
		LineChartFromDB(cd)
		Return True
	End Function
#End Region




End Class
