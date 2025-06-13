Imports System.Collections.Concurrent

Public Class DefaultOrmMetadataCache
    Implements IOrmMetadataCache

    Private ReadOnly _cache As New ConcurrentDictionary(Of Type, OrmMetadata)

    Public Function GetMetadata(entityType As Type) As OrmMetadata Implements IOrmMetadataCache.GetMetadata
        Return _cache.GetOrAdd(entityType, Function(t) OrmMetadataBuilder.Build(t))
    End Function

    Public Sub Clear() Implements IOrmMetadataCache.Clear
        _cache.Clear()
    End Sub
End Class
