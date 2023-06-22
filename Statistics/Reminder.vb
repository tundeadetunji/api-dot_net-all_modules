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
Imports Statistics.Charts
Public Class Reminder

#Region "Fields"
    Private Property con_string__ As String
    Private table__ As String = "Reminder"
    Private table_Charts__ As String = "Chart"
#End Region

    Public Sub New(connection_string As String)
        con_string__ = connection_string
    End Sub

    Public Function SetReminder(id_ As String, OperationType As String, functionField As String, FieldValue As Object, sender As Object, page_OR_update_panel As Object, instance_of_script_manager As Object) As Boolean
        If FieldValue.ToString.Length < 1 Then Return False

        Dim field_name As String = QData(BuildSelectString(table_Charts__, {"FieldName"}, {"FieldFriendlyName", "ID"}), con_string__, {"FieldFriendlyName", functionField, "ID", id_})
        Dim table_name As String = QData(BuildSelectString(table_Charts__, {"TableName"}, {"FieldFriendlyName", "ID"}), con_string__, {"FieldFriendlyName", functionField, "ID", id_})

        If CommitSequel(BuildInsertString(table__, {"FieldName", "TableName", "OperationType", "FieldValue", "ID"}), con_string__, {"FieldName", field_name, "TableName", table_name, "OperationType", OperationType, "FieldValue", FieldValue, "ID", id_}) = True Then
            JavaScript("Alert", "Reminder has been set.", sender, page_OR_update_panel, instance_of_script_manager)
        Else
            JavaScript("Alert", "General Error. Reminder has not been set.", sender, page_OR_update_panel, instance_of_script_manager)
        End If
    End Function

    Public Function CheckForUpdate(id__ As String, g As GridView, g_ As GridView, sender As Object, page_OR_update_panel As Object, instance_of_script_manager As Object)
        Dim operation_type As String
        Dim field_name As String
        Dim table_name As String
        Dim field_value
        Dim sum = 0
        Dim row_count = 0
        Dim r = 0
        Dim r_ As String = ""
        Dim id_
		Display(g, BuildTopString(table__, Nothing, {"ID"}, 1), con_string__, {"ID", id__})
		With g
            If .Rows.Count < 1 Then Return False
            .Visible = False
            id_ = .Rows(0).Cells(0).Text
            field_name = .Rows(0).Cells(3).Text
            table_name = .Rows(0).Cells(4).Text
            operation_type = .Rows(0).Cells(5).Text
            field_value = .Rows(0).Cells(6).Text
        End With
        Display(g_, BuildSelectString(table_name, {field_name}, {"ID", id__}), con_string__)
        With g_
            If .Rows.Count > 0 Then
                .Visible = False
                row_count = .Rows.Count
                For i As Integer = 0 To .Rows.Count - 1
                    sum += Val(.Rows(i).Cells(0).Text)
                Next
            End If
        End With

        Select Case operation_type
			Case DBCharts.OperationType.Sum.ToString
				r = sum
			Case DBCharts.StatIs.Sum.ToString
				r = sum
        End Select

        If sum <> 0 And row_count <> 0 Then
            Select Case operation_type
				Case DBCharts.OperationType.AVG.ToString
					r = sum / row_count
				Case DBCharts.StatIs.Average.ToString
					r = sum / row_count

				Case DBCharts.OperationType.Count.ToString
					r = row_count
				Case DBCharts.StatIs.Count.ToString
					r = row_count
            End Select
        End If
        If r >= field_value Then
            Select Case operation_type
				Case DBCharts.OperationType.Sum.ToString
					r_ = "Total of all " & field_name & " has reached " & field_value & ". It is now " & r & "."
				Case DBCharts.StatIs.Sum.ToString
					r_ = "Total of all " & field_name & " has reached " & field_value & ". It is now " & r & "."
				Case DBCharts.OperationType.AVG.ToString
					r_ = "Average of all " & field_name & " has reached " & field_value & ". It is now " & r & "."
				Case DBCharts.StatIs.Average.ToString
					r_ = "Average of all " & field_name & " has reached " & field_value & ". It is now " & r & "."
				Case DBCharts.OperationType.Count.ToString
					r_ = "Number of occurrences of all " & field_name & " has reached " & field_value & ". It is now " & r & "."
				Case DBCharts.StatIs.Count.ToString
					r_ = "Number of occurrences of all " & field_name & " has reached " & field_value & ". It is now " & r & "."
            End Select
            Try
                CommitSequel(DeleteString_CONDITIONAL(table__, {"RecordSerial", "=", "ID", "="}), con_string__, {"RecordSerial", id_, "ID", id__})
            Catch
            End Try
            Return r_
        Else
            Return False
        End If

    End Function

End Class
