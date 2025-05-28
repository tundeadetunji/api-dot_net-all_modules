<AttributeUsage(AttributeTargets.Property)>
Public Class EnumStorageAttribute
    Inherits Attribute
    Public Property StorageType As EnumStorageType

    Public Sub New(storageType As EnumStorageType)
        Me.StorageType = storageType
    End Sub
End Class