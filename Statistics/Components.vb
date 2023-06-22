Imports HModule.H
Imports General_Module.FormatWindow
Imports NModule.InternalTypes
Imports NModule.NFunctions
Imports NModule.SJ
Imports System.IO
Imports Web_Module.DW
Imports Web_Module.Methods
Imports Web_Module.DataConnectionAPI
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Syncfusion.DocIO.DLS
Imports JModule.J
Imports Miscellaneous.GeneralModule
Imports Statistics.Bank
Imports Statistics.Charts
Imports Statistics.Cookies
Imports Statistics.Reminder
Imports Statistics.Traffic
Imports NModule.W
Imports Web_Module.DataConnectionWeb
Imports Web_Module.Functions
Imports HModule.F
Imports System.Web.Profile
Imports Microsoft.AspNet.Identity
Imports HModule.N
Imports JModule.W
Imports System.Web.UI.WebControls
Imports Web_Module.Bootstrap
Imports System.Web.UI.HtmlControls

Public Class Components
    'Public Shared Function InvoiceCount(gridCountInvoice As GridView, connection_string As String, divCountInvoice As HtmlGenericControl, Optional date_ As Date = Nothing, Optional glyph_class As String = "fa fa-file-text-o", Optional date_as_long_not_short As Boolean = True)
    '    Dim date__ As Date
    '    If date_ = Nothing Then
    '        date__ = Now
    '    Else
    '        date__ = date_
    '    End If

    '    Try
    '        gridCountInvoice.Visible = False
    '        Dim count_invoice_query As String = BuildSelectString_DISTINCT("Sales", {"InvoiceNumber"}, {"RecordDate"})
    '        Dim count_invoice_query_kv As Array = {"RecordDate", date__}
    '        Dim value = Display(gridCountInvoice, count_invoice_query, connection_string, count_invoice_query_kv).Rows.Count
    '        Return InformationCard(value, ToPlural(value, SingularWord.Invoice, "", False, TextCase.Capitalize), getDateString(date_), ColorCode.Success, glyph_class, divCountInvoice)
    '    Catch
    '    End Try
    'End Function
    'Public Shared Function SaleCount(gridCountInvoice As GridView, connection_string As String, divCountInvoice As HtmlGenericControl, Optional date_ As Date = Nothing, Optional glyph_class As String = "fa fa-file-text-o", Optional date_as_long_not_short As Boolean = True)
    '    Dim date__ As Date
    '    If date_ = Nothing Then
    '        date__ = Now
    '    Else
    '        date__ = date_
    '    End If

    '    Try
    '        gridCountInvoice.Visible = False
    '        Dim count_invoice_query As String = BuildSelectString_DISTINCT("Sales", {"InvoiceNumber"}, {"RecordDate"})
    '        Dim count_invoice_query_kv As Array = {"RecordDate", DateToSQL(date__)}
    '        Dim value = Display(gridCountInvoice, count_invoice_query, connection_string, count_invoice_query_kv).Rows.Count
    '        Return InformationCard(value, ToPlural(value, SingularWord.Sale, "", False, TextCase.Capitalize), getDateString(date_), ColorCode.Success, glyph_class, divCountInvoice)
    '    Catch
    '    End Try
    'End Function
    'Public Shared Function ReceiptCount(gridCountInvoice As GridView, connection_string As String, divCountInvoice As HtmlGenericControl, Optional date_ As Date = Nothing, Optional glyph_class As String = "fa fa-file-text-o", Optional date_as_long_not_short As Boolean = True)
    '    Dim date__ As Date
    '    If date_ = Nothing Then
    '        date__ = Now
    '    Else
    '        date__ = date_
    '    End If

    '    Try
    '        gridCountInvoice.Visible = False
    '        Dim count_invoice_query As String = BuildSelectString_DISTINCT("Sales", {"InvoiceNumber"}, {"RecordDate"})
    '        Dim count_invoice_query_kv As Array = {"RecordDate", DateToSQL(date__)}
    '        Dim value = Display(gridCountInvoice, count_invoice_query, connection_string, count_invoice_query_kv).Rows.Count
    '        Return InformationCard(value, ToPlural(value, SingularWord.Receipt, "", False, TextCase.Capitalize), getDateString(date_), ColorCode.Success, glyph_class, divCountInvoice)
    '    Catch
    '    End Try
    'End Function

    'Public Shared Function QuantitySold(connection_string As String, divQuantitySoldToday As HtmlGenericControl, Optional date_ As Date = Nothing, Optional glyph_class As String = "fa fa-file-text-o", Optional date_as_long_not_short As Boolean = True)
    '    Dim date__ As Date
    '    If date_ = Nothing Then
    '        date__ = Now
    '    Else
    '        date__ = date_
    '    End If

    '    Try
    '        Dim qty_sold_query As String = BuildSumString_UNGROUPED("Sales", "Quantity", {"RecordDate"})
    '        Dim qty_sold_query_kv As Array = {"RecordDate", DateToSQL(date__)}
    '        Dim value = QData(qty_sold_query, connection_string, qty_sold_query_kv)
    '        Return InformationCard(value, ToPlural(value, SingularWord.Product, " Sold", False, TextCase.Capitalize), getDateString(date_), ColorCode.Success, glyph_class, divQuantitySoldToday)
    '    Catch
    '    End Try
    'End Function

    Public Shared Function ChartsDrop(drop As DropDownList, chart_pattern As ChartPattern) As DropDownList
        Select Case chart_pattern
            Case ChartPattern.Bar
                BindProperty(drop, GetEnum(New ChartFromPatternIsBar))
            Case ChartPattern.BarLine
                BindProperty(drop, GetEnum(New ChartFromPatternIsBarLine))
            Case ChartPattern.PieDoughnut
                BindProperty(drop, GetEnum(New ChartFromPatternIsPieDoughnut))
            Case ChartPattern.All
                BindProperty(drop, GetEnum(New ChartIs))
        End Select
        Return drop
    End Function
    Public Shared Function ChartsDropBar(drop As DropDownList)
        Return ChartsDrop(drop, ChartPattern.Bar)
    End Function
    Public Shared Function ChartsDropBarLine(drop As DropDownList)
        Return ChartsDrop(drop, ChartPattern.BarLine)
    End Function
    Public Shared Function ChartsDropPieDoughnut(drop As DropDownList)
        Return ChartsDrop(drop, ChartPattern.PieDoughnut)
    End Function
    Public Shared Function ChartsDrop(drops As Array, chart_pattern As ChartPattern) As Array
        Try
            For i = 0 To drops.Length - 1
                ChartsDrop(drops(i), chart_pattern)
            Next
        Catch ex As Exception

        End Try
        Return drops
    End Function

#Region "Support Functions"

    'Private Shared Function getDateString(Optional date_ As Date = Nothing) As String
    '    If date_ = Nothing Then
    '        Return "Today, as at " & Now.ToShortTimeString
    '    Else
    '        Return "On " And Date.Parse(date_).ToLongDateString
    '    End If

    'End Function

#End Region
End Class
