<AttributeUsage(AttributeTargets.Property)>
Public Class JoinTableAttribute
    Inherits Attribute

    Public ReadOnly Property TableName As String
    Public ReadOnly Property ThisKey As String
    Public ReadOnly Property OtherKey As String

    Public Sub New(tableName As String, thisKey As String, otherKey As String)
        Me.TableName = tableName
        Me.ThisKey = thisKey
        Me.OtherKey = otherKey
    End Sub
End Class
