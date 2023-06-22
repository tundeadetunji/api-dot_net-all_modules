Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Public Class JSON
#Region "J"
	Public Shared Function ToJArray(list_or_array As Object) As JArray
		Return JsonConvert.SerializeObject(list_or_array)
	End Function
	Public Shared Function ToJObject(dict As Dictionary(Of String, Object), Optional format As Formatting = Formatting.Indented)
		Return JsonConvert.SerializeObject(dict, format)
	End Function
	Public Shared Function JSONSerializeIt(list_or_dict As Object, Optional format As Formatting = Formatting.Indented)
		Return JsonConvert.SerializeObject(list_or_dict, format)

		'Dim a As New Dictionary(Of String, Object)
		'a.Add("d1", "v1")
		'a.Add("d2", "v1")
		't.Text = JSONSerializeIt(a)


		'Dim p As New List(Of String)
		'p.Add("d1")
		'p.Add("d2")
		't.Text = JSONSerializeIt(p)


		'Dim l As New List(Of Object)
		'l.Add("l1")
		'l.Add("l2")
		'Dim d As New Dictionary(Of String, Object)
		'd.Add("k1", l)
		'Dim jarray = {"a", d}
		't.Text = JSONSerializeIt(jarray)

	End Function
	Public Shared Function JSONDeserializeIt(serialized_)
		Return JsonConvert.DeserializeObject(Of Object)(serialized_)


		'Dim p As JObject = JSONDeserializeIt(t.Text)


		'Dim l As JArray = JSONDeserializeIt(t.Text)

	End Function

#End Region

#Region "JRetrieval"
	Public Shared Function JSONProperties(o As JObject) As List(Of Object)
		Dim p As New List(Of Object)

		For i = 0 To o.Properties.Count - 1
			'If toReturn = JSONToReturn.Properties Or toReturn = JSONToReturn.PropertiesAndValues Then
			p.Add(o.Properties(i).Name)
			'ElseIf toReturn = JSONToReturn.Values Or toReturn = JSONToReturn.PropertiesAndValues
			'	v.Add(o.PropertyValues(i).ToString)
			'End If
		Next
		Return p
	End Function
	Public Shared Function JSONValues(a As JArray) As List(Of Object)
		Dim p As New List(Of Object)

		For i = 0 To a.Count - 1
			p.Add(a.Item(i))
		Next
		Return p


		'u.Text = ListToString(JSONValues(JArray.Parse(ReadText(file__))))

	End Function

	Public Shared Function JSONValue(o As JObject, property_ As String) As Object
		Return o.Item(property_) '.ToString
	End Function

#End Region

End Class
