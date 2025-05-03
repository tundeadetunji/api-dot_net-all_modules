Imports System.Threading
Public Class SequelOrm
    Implements IOrm

#Region "init"
    Private Const Id As String = "Id"

    Private Shared _instance As Lazy(Of SequelOrm)
    Private ReadOnly strategy As IOrm

    Private Sub New(provider As SupportedDbProvider, connectionString As String)
        Select Case provider
            Case SupportedDbProvider.SqlServer
                strategy = SqlServerOrm.Instance(connectionString)
                'Case SupportedDbProvider.MySQL
                '    strategy = MySqlOrm.Instance(connectionString)
                'Case SupportedDbProvider.PostgreSQL
                '    strategy = PostgreSqlOrm.Instance(connectionString)
            Case Else
                Throw New NotSupportedException($"Provider '{provider}' is not supported.")
        End Select
    End Sub

    Public Shared Function GetInstance(provider As SupportedDbProvider, connectionString As String) As SequelOrm
        If _instance Is Nothing Then
            _instance = New Lazy(Of SequelOrm)(Function() New SequelOrm(provider, connectionString), LazyThreadSafetyMode.ExecutionAndPublication)
        End If
        Return _instance.Value
    End Function
#End Region

#Region "exported"

    Public Sub DeleteAll(Of T)(Optional cascade As Boolean = True) Implements IOrm.DeleteAll
        strategy.DeleteAll(Of T)(cascade)
    End Sub

    Public Sub DeleteAllInTable(Of T)(tableName As String, Optional cascade As Boolean = True) Implements IOrm.DeleteAllInTable
        strategy.DeleteAllInTable(Of T)(tableName, cascade)
    End Sub

    Public Sub DeleteWhereId(Of T)(id As Object, Optional idColumn As String = "Id", Optional cascade As Boolean = True) Implements IOrm.DeleteWhereId
        strategy.DeleteWhereId(Of T)(id, idColumn, cascade)
    End Sub

    Public Sub DeleteWhereIdInTable(Of T)(id As Object, tableName As String, Optional idColumn As String = "Id", Optional cascade As Boolean = True) Implements IOrm.DeleteWhereIdInTable
        strategy.DeleteWhereIdInTable(Of T)(id, tableName, idColumn, cascade)
    End Sub

    Public Sub DeleteBy(Of T)(conditions As List(Of Condition), Optional cascade As Boolean = True) Implements IOrm.DeleteBy
        strategy.DeleteBy(Of T)(conditions, cascade)
    End Sub

    Public Sub DeleteByInTable(Of T)(conditions As List(Of Condition), tableName As String, Optional cascade As Boolean = True) Implements IOrm.DeleteByInTable
        strategy.DeleteByInTable(Of T)(conditions, tableName, cascade)
    End Sub

    Public Function CountBy(Of T)(conditions As List(Of Condition)) As Long Implements IOrm.CountBy
        Return strategy.CountBy(Of T)(conditions)
    End Function

    Public Function CountByInTable(Of T)(conditions As List(Of Condition), tableName As String) As Long Implements IOrm.CountByInTable
        Return strategy.CountByInTable(Of T)(conditions, tableName)
    End Function

    Public Function Count(Of T)() As Long Implements IOrm.Count
        Return strategy.Count(Of T)
    End Function

    Public Function CountInTable(Of T)(tableName As String) As Long Implements IOrm.CountInTable
        Return strategy.CountInTable(Of T)(tableName)
    End Function

    Public Function Create(Of T)(objs As List(Of T), Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As List(Of T) Implements IOrm.Create
        Return strategy.Create(Of T)(objs, idColumn, IdWillAutoIncrement)
    End Function

    Public Function Create(Of T)(obj As T, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As T Implements IOrm.Create
        Return strategy.Create(Of T)(obj, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateInTable(Of T)(objs As List(Of T), tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As List(Of T) Implements IOrm.CreateInTable
        Return strategy.CreateInTable(Of T)(objs, tableName, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As T Implements IOrm.CreateInTable
        Return strategy.CreateInTable(Of T)(obj, tableName, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateOrUpdate(Of T)(obj As T, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As T Implements IOrm.CreateOrUpdate
        Return strategy.CreateOrUpdate(Of T)(obj, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateOrUpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As T Implements IOrm.CreateOrUpdateInTable
        Return strategy.CreateOrUpdateInTable(Of T)(obj, tableName, idColumn, IdWillAutoIncrement)
    End Function

    Public Function ExistsById(Of T)(id As Object, Optional idColumn As String = "Id") As Boolean Implements IOrm.ExistsById
        Return strategy.ExistsById(Of T)(id, idColumn)
    End Function

    Public Function ExistsByIdInTable(id As Object, tableName As String, Optional idColumn As String = "Id") As Boolean Implements IOrm.ExistsByIdInTable
        Return strategy.ExistsByIdInTable(id, tableName, idColumn)
    End Function

    Public Function ExistsBy(Of T)(conditions As List(Of Condition)) As Boolean Implements IOrm.ExistsBy
        Return strategy.ExistsBy(Of T)(conditions)
    End Function

    Public Function ExistsByInTable(Of T)(conditions As List(Of Condition), tableName As String) As Boolean Implements IOrm.ExistsByInTable
        Return strategy.ExistsByInTable(Of T)(conditions, tableName)
    End Function

    Public Function FindAll(Of T)(Optional idColumn As String = "Id", Optional ascending As Boolean = True) As List(Of T) Implements IOrm.FindAll
        Return strategy.FindAll(Of T)(idColumn, ascending)
    End Function

    Public Function FindAllInTable(Of T)(tableName As String, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As List(Of T) Implements IOrm.FindAllInTable
        Return strategy.FindAllInTable(Of T)(tableName, idColumn, ascending)
    End Function

    Public Function FindAllPaged(Of T)(pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Page(Of T) Implements IOrm.FindAllPaged
        Return strategy.FindAllPaged(Of T)(pageNumber, maxPerPage, idColumn, ascending)
    End Function

    Public Function FindAllPagedInTable(Of T)(tableName As String, pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Page(Of T) Implements IOrm.FindAllPagedInTable
        Return strategy.FindAllPagedInTable(Of T)(tableName, pageNumber, maxPerPage, idColumn, ascending)
    End Function

    Public Function FindById(Of T)(id As Object, Optional idColumn As String = "Id") As T Implements IOrm.FindById
        Return strategy.FindById(Of T)(id, idColumn)
    End Function

    Public Function FindByIdInTable(Of T)(id As Object, tableName As String, Optional idColumn As String = "Id") As T Implements IOrm.FindByIdInTable
        Return strategy.FindByIdInTable(Of T)(id, tableName, idColumn)
    End Function

    Public Function FindBy(Of T)(conditions As List(Of Condition)) As List(Of T) Implements IOrm.FindBy
        Return strategy.FindBy(Of T)(conditions)
    End Function

    Public Function FindByInTable(Of T)(conditions As List(Of Condition), tableName As String) As List(Of T) Implements IOrm.FindByInTable
        Return strategy.FindByInTable(Of T)(conditions, tableName)
    End Function

    Public Function FindByPaged(Of T)(conditions As List(Of Condition), pageNumber As Integer, maxPerPage As Integer) As Page(Of T) Implements IOrm.FindByPaged
        Return strategy.FindByPaged(Of T)(conditions, pageNumber, maxPerPage)
    End Function

    Public Function FindByPagedInTable(Of T)(conditions As List(Of Condition), tableName As String, pageNumber As Integer, maxPerPage As Integer) As Page(Of T) Implements IOrm.FindByPagedInTable
        Return strategy.FindByPagedInTable(Of T)(conditions, tableName, pageNumber, maxPerPage)
    End Function

    Public Function Update(Of T)(obj As T, Optional idColumn As String = "Id") As T Implements IOrm.Update
        Return strategy.Update(Of T)(obj, idColumn)
    End Function

    Public Function UpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id") As T Implements IOrm.UpdateInTable
        Return strategy.UpdateInTable(Of T)(obj, tableName, idColumn)
    End Function

    Public Sub PrepareDatabase(mode As DbPrepMode, entities As List(Of Type), Optional idColumn As String = Id) Implements IOrm.PrepareDatabase
        strategy.PrepareDatabase(mode, entities, idColumn)
    End Sub


#End Region

End Class