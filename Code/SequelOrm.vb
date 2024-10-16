Imports System.Reflection
Imports Newtonsoft.Json.Linq
Public Class SequelOrm

    Public Function MapDictionaryToObject(Of T As Class)(dictionary As Dictionary(Of String, Object), obj As T) As T
        If dictionary Is Nothing Then
            Throw New ArgumentNullException(NameOf(dictionary))
        End If

        If obj Is Nothing Then
            Throw New ArgumentNullException(NameOf(obj))
        End If

        Dim type = GetType(T)
        Dim properties = type.GetProperties(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        Dim value As Object = Nothing

        For Each prop In properties
            If dictionary.TryGetValue(prop.Name, value) Then
                If prop.CanWrite Then
                    prop.SetValue(obj, value)
                Else
                    Dim field = type.GetField(prop.Name, BindingFlags.Instance Or BindingFlags.NonPublic)
                    If field IsNot Nothing Then
                        field.SetValue(obj, value)
                    End If
                End If
            End If
        Next

        ' Handle private fields
        Dim fields = type.GetFields(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)

        For Each field In fields
            If dictionary.TryGetValue(field.Name, value) Then
                field.SetValue(obj, value)
            End If
        Next

        Return obj
    End Function


End Class

'Save
'DeleteById
'DeleteAll
'SaveAll
'FindAll
'FindById
'ExistsById
'Count
