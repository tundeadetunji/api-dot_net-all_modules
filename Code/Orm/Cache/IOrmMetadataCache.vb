Public Interface IOrmMetadataCache
    Function GetMetadata(entityType As Type) As OrmMetadata
    Sub Clear() ' Optional for test environments
End Interface
