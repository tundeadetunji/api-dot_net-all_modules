
''' <summary>
''' Allows explicit SQL type specification for a property (e.g. "TEXT", "NVARCHAR(MAX)").
''' </summary>
<AttributeUsage(AttributeTargets.Property)>
Public Class SqlTypeAttribute
    Inherits Attribute
    Public Property SqlType As String
    Public Sub New(sqlType As String)
        Me.SqlType = sqlType
    End Sub
End Class