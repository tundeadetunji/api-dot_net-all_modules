Imports System.Reflection

Public Class OrmMetadata
    Public Property EntityType As Type
    Public Property TableName As String
    Public Property Properties As List(Of PropertyInfo)
    Public Property EnumStorageTypes As Dictionary(Of PropertyInfo, EnumStorageType) ' <- updated
    Public Property ChildCollections As List(Of PropertyInfo)
    Public Property NavigationProperties As List(Of PropertyInfo)
    Public Property IdProperty As PropertyInfo
    Public Property RowVersionProperty As PropertyInfo
End Class

