<AttributeUsage(AttributeTargets.Property)>
Public Class ForeignKeyAttribute
    Inherits Attribute

    Public ReadOnly Property ColumnName As String

    Public Sub New(columnName As String)
        Me.ColumnName = columnName
    End Sub
End Class
