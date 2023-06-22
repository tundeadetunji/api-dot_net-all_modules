Imports Web_Module.Functions
Imports System.Web.UI.HtmlControls
Imports Bank.Transactions
Imports NModule.W
Imports NModule.NFunctions
Imports System.Threading.Tasks
Imports General_Module.FormatWindow
Public Class Bank

	Public Structure PayInfo
		Public Operation As Operation '= Operation.Volume
		Public Key As String '= ""
		Public From As String '= ""
		Public To_ As String '= ""
		Public Customer_Code_OR_Email As String '= ""
		Public PreferredOutput As PreferredReturnIs '= PreferredReturnIs.Regular
	End Structure
	Public Shared Async Function ReceiptCards(Key As String, Operation As Operation, div As HtmlGenericControl, Optional Customer_Code_OR_Email As String = Nothing, Optional From As String = Nothing, Optional To_ As String = Nothing, Optional PreferredOutput As PreferredReturnIs = PreferredReturnIs.Regular, Optional show_customer_email As Boolean = True) As Task(Of String)
		Dim code_ As String = "", From_ As String = "", To__ As String = ""
		If Customer_Code_OR_Email IsNot Nothing Then code_ = Customer_Code_OR_Email
		If From IsNot Nothing Then From_ = From
		If To_ IsNot Nothing Then To__ = To_

		Dim p As New PayInfo With {.Operation = Operation, .Key = Key, .From = From_, .To_ = To__, .Customer_Code_OR_Email = code_, .PreferredOutput = PreferredOutput}
		Return ReceiptCards(p, div, show_customer_email).Result

	End Function

	Public Shared Async Function ReceiptCards(ref As PayInfo, div As HtmlGenericControl, Optional show_customer_email As Boolean = True) As Task(Of String)
		Dim id As PayInfo = ref
		Dim l_headlines As New List(Of String)
		Dim l_details As New List(Of String)
		Dim l_dates As New List(Of String)

		Dim Operation As Operation = id.Operation
		Dim Key As String = id.Key
		Dim From As String = id.From
		Dim To_ As String = id.To_
		Dim Customer_Code_OR_Email As String = id.Customer_Code_OR_Email
		Dim AsList As Boolean = True ' id.AsList
		Dim PreferredOutput As PreferredReturnIs = id.PreferredOutput

		'Dim r As String
		'Dim a As Array
		Select Case Operation
			Case Operation.Volume
				l_headlines.Add("As Of " & Now.ToLongDateString)
				'r = 
				l_details.Add(ToIO(Volume(Key, From, To_).Result))
				l_dates.Add("")
			Case Operation.VolumesFromCustomer
				Dim a As Array = VolumesFromCustomer(Key, Customer_Code_OR_Email, AsList, PreferredOutput).Result
				With a
					For i As Integer = 0 To .Length - 1
						l_headlines.Add("")
						'l_details.Add(a(i))
						If show_customer_email Then
							l_dates.Add("from " & Customer_Code_OR_Email)
						Else
							l_dates.Add("")
						End If
					Next
				End With
				l_details = ArrayToList(a, ListIsOf.String_)
			Case Operation.VolumeTotal
				l_headlines.Add("As Of " & Now.ToLongDateString)
				'r = 
				l_details.Add(VolumeTotal(Key, From, To_).Result)
				l_dates.Add("")
			Case Operation.VolumeTotalFromCustomer
				l_headlines.Add("As Of " & Now.ToLongDateString)
				'r = 
				l_details.Add(VolumeTotalFromCustomer(Key, Customer_Code_OR_Email).Result)
				If show_customer_email Then
					l_dates.Add("from " & Customer_Code_OR_Email)
				Else
					l_dates.Add("")
				End If
		End Select

		Dim cards_ As String = "<div class=""card-columns"">"
		Dim card_ As String
		'Dim num As Long = 
		'If Operation = Operation.VolumesFromCustomer Then num = l_headlines.Count
		With l_headlines
			For j As Integer = 0 To l_headlines.Count - 1
				card_ = "<div class=""card text-center"">
					<div class=""card-body"">
					<h5 class=""card-title"">" & l_headlines(j) & "</h5>
					<p class=""card-text"">" & l_details(j) & "</p>
					<p class=""card-text""><small class=""text-muted"">" & l_dates(j) & "</small></p>
					</div>
					</div>"
				cards_ &= card_
			Next

		End With
		cards_ &= "</div>"

		WriteContent(cards_, div)
		Return cards_
	End Function


End Class
