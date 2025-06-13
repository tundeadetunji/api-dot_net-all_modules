Imports System.Reflection

Public NotInheritable Class OrmMetadataBuilder
    Public Shared Function Build(entityType As Type) As OrmMetadata
        Dim props = entityType.GetProperties(BindingFlags.Public Or BindingFlags.Instance).ToList()

        Dim metadata As New OrmMetadata With {
            .EntityType = entityType,
            .TableName = entityType.Name, ' or use your GetTableName if implemented
            .Properties = props,
            .EnumStorageTypes = props.
                Where(Function(p) p.PropertyType.IsEnum OrElse Nullable.GetUnderlyingType(p.PropertyType)?.IsEnum).
                ToDictionary(Function(p) p, Function(p) OrmUtils.GetEnumStorageType(p)),
            .ChildCollections = props.
                Where(Function(p) IsGenericList(p.PropertyType)).
                ToList(),
            .NavigationProperties = props.
                Where(Function(p) IsClassButNotString(p.PropertyType) AndAlso Not IsGenericList(p.PropertyType)).
                ToList(),
            .IdProperty = OrmUtils.FindIdProperty(entityType, "Id"), ' replace with configurable idColumn if needed
            .RowVersionProperty = props.FirstOrDefault(Function(p) String.Equals(p.Name, "RowVersion", StringComparison.OrdinalIgnoreCase))
        }

        Return metadata
    End Function

    Private Shared Function IsGenericList(type As Type) As Boolean
        Return type.IsGenericType AndAlso type.GetGenericTypeDefinition() = GetType(List(Of))
    End Function

    Private Shared Function IsClassButNotString(type As Type) As Boolean
        Return type.IsClass AndAlso type IsNot GetType(String)
    End Function
End Class
