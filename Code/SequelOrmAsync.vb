Imports System.Threading

Public Class SequelOrmAsync
    Implements IOrmAsync

#Region "init"
    Private Const Id As String = "Id"

    Private Shared _instance As Lazy(Of SequelOrmAsync)
    Private ReadOnly strategy As IOrmAsync

    Private Sub New(provider As SupportedDbProvider, connectionString As String)
        Select Case provider
            Case SupportedDbProvider.SqlServer
                strategy = SqlServerOrmAsync.Instance(connectionString)
            Case Else
                Throw New NotSupportedException($"Provider '{provider}' is not supported.")
        End Select
    End Sub

    Public Shared Function GetInstance(provider As SupportedDbProvider, connectionString As String) As SequelOrmAsync
        If _instance Is Nothing Then
            _instance = New Lazy(Of SequelOrmAsync)(Function() New SequelOrmAsync(provider, connectionString), LazyThreadSafetyMode.ExecutionAndPublication)
        End If
        Return _instance.Value
    End Function
#End Region

#Region "exported"

    Public Sub DeleteAllAsync(Of T)(Optional cascade As Boolean = True) Implements IOrmAsync.DeleteAllAsync
        strategy.DeleteAllAsync(Of T)(cascade)
    End Sub

    Public Sub DeleteAllInTableAsync(Of T)(tableName As String, Optional cascade As Boolean = True) Implements IOrmAsync.DeleteAllInTableAsync
        strategy.DeleteAllInTableAsync(Of T)(tableName, cascade)
    End Sub

    Public Sub DeleteWhereIdAsync(Of T)(id As Object, Optional idColumn As String = "Id", Optional cascade As Boolean = True) Implements IOrmAsync.DeleteWhereIdAsync
        strategy.DeleteWhereIdAsync(Of T)(id, idColumn, cascade)
    End Sub

    Public Sub DeleteWhereIdInTableAsync(Of T)(id As Object, tableName As String, Optional idColumn As String = "Id", Optional cascade As Boolean = True) Implements IOrmAsync.DeleteWhereIdInTableAsync
        strategy.DeleteWhereIdInTableAsync(Of T)(id, tableName, idColumn, cascade)
    End Sub

    Public Sub DeleteByAsync(Of T)(conditions As List(Of Condition), Optional cascade As Boolean = True) Implements IOrmAsync.DeleteByAsync
        strategy.DeleteByAsync(Of T)(conditions, cascade)
    End Sub

    Public Sub DeleteByInTableAsync(Of T)(conditions As List(Of Condition), tableName As String, Optional cascade As Boolean = True) Implements IOrmAsync.DeleteByInTableAsync
        strategy.DeleteByInTableAsync(Of T)(conditions, tableName, cascade)
    End Sub

    Public Function CountByAsync(Of T)(conditions As List(Of Condition)) As Task(Of Long) Implements IOrmAsync.CountByAsync
        Return strategy.CountByAsync(Of T)(conditions)
    End Function

    Public Function CountByInTableAsync(Of T)(conditions As List(Of Condition), tableName As String) As Task(Of Long) Implements IOrmAsync.CountByInTableAsync
        Return strategy.CountByInTableAsync(Of T)(conditions, tableName)
    End Function

    Public Function CountAsync(Of T)() As Task(Of Long) Implements IOrmAsync.CountAsync
        Return strategy.CountAsync(Of T)
    End Function

    Public Function CountInTableAsync(Of T)(tableName As String) As Task(Of Long) Implements IOrmAsync.CountInTableAsync
        Return strategy.CountInTableAsync(Of T)(tableName)
    End Function

    Public Function CreateAsync(Of T)(objs As List(Of T), Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of List(Of T)) Implements IOrmAsync.CreateAsync
        Return strategy.CreateAsync(Of T)(objs, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateAsync(Of T)(obj As T, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of T) Implements IOrmAsync.CreateAsync
        Return strategy.CreateAsync(Of T)(obj, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateInTableAsync(Of T)(objs As List(Of T), tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of List(Of T)) Implements IOrmAsync.CreateInTableAsync
        Return strategy.CreateInTableAsync(Of T)(objs, tableName, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateInTableAsync(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of T) Implements IOrmAsync.CreateInTableAsync
        Return strategy.CreateInTableAsync(Of T)(obj, tableName, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateOrUpdateAsync(Of T)(obj As T, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of T) Implements IOrmAsync.CreateOrUpdateAsync
        Return strategy.CreateOrUpdateAsync(Of T)(obj, idColumn, IdWillAutoIncrement)
    End Function

    Public Function CreateOrUpdateInTableAsync(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As Task(Of T) Implements IOrmAsync.CreateOrUpdateInTableAsync
        Return strategy.CreateOrUpdateInTableAsync(Of T)(obj, tableName, idColumn, IdWillAutoIncrement)
    End Function

    Public Function ExistsByIdAsync(Of T)(id As Object, Optional idColumn As String = "Id") As Task(Of Boolean) Implements IOrmAsync.ExistsByIdAsync
        Return strategy.ExistsByIdAsync(Of T)(id, idColumn)
    End Function

    Public Function ExistsByIdInTableAsync(id As Object, tableName As String, Optional idColumn As String = "Id") As Task(Of Boolean) Implements IOrmAsync.ExistsByIdInTableAsync
        Return strategy.ExistsByIdInTableAsync(id, tableName, idColumn)
    End Function

    Public Function ExistsByAsync(Of T)(conditions As List(Of Condition)) As Task(Of Boolean) Implements IOrmAsync.ExistsByAsync
        Return strategy.ExistsByAsync(Of T)(conditions)
    End Function

    Public Function ExistsByInTableAsync(Of T)(conditions As List(Of Condition), tableName As String) As Task(Of Boolean) Implements IOrmAsync.ExistsByInTableAsync
        Return strategy.ExistsByInTableAsync(Of T)(conditions, tableName)
    End Function

    Public Function FindAllAsync(Of T)(Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Task(Of List(Of T)) Implements IOrmAsync.FindAllAsync
        Return strategy.FindAllAsync(Of T)(idColumn, ascending)
    End Function

    Public Function FindAllInTableAsync(Of T)(tableName As String, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Task(Of List(Of T)) Implements IOrmAsync.FindAllInTableAsync
        Return strategy.FindAllInTableAsync(Of T)(tableName, idColumn, ascending)
    End Function

    Public Function FindAllPagedAsync(Of T)(pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Task(Of Page(Of T)) Implements IOrmAsync.FindAllPagedAsync
        Return strategy.FindAllPagedAsync(Of T)(pageNumber, maxPerPage, idColumn, ascending)
    End Function

    Public Function FindAllPagedInTableAsync(Of T)(tableName As String, pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Task(Of Page(Of T)) Implements IOrmAsync.FindAllPagedInTableAsync
        Return strategy.FindAllPagedInTableAsync(Of T)(tableName, pageNumber, maxPerPage, idColumn, ascending)
    End Function

    Public Function FindByIdAsync(Of T)(id As Object, Optional idColumn As String = "Id") As Task(Of T) Implements IOrmAsync.FindByIdAsync
        Return strategy.FindByIdAsync(Of T)(id, idColumn)
    End Function

    Public Function FindByIdInTableAsync(Of T)(id As Object, tableName As String, Optional idColumn As String = "Id") As Task(Of T) Implements IOrmAsync.FindByIdInTableAsync
        Return strategy.FindByIdInTableAsync(Of T)(id, tableName, idColumn)
    End Function

    Public Function FindByAsync(Of T)(conditions As List(Of Condition)) As Task(Of List(Of T)) Implements IOrmAsync.FindByAsync
        Return strategy.FindByAsync(Of T)(conditions)
    End Function

    Public Function FindByInTableAsync(Of T)(conditions As List(Of Condition), tableName As String) As Task(Of List(Of T)) Implements IOrmAsync.FindByInTableAsync
        Return strategy.FindByInTableAsync(Of T)(conditions, tableName)
    End Function

    Public Function FindByPagedAsync(Of T)(conditions As List(Of Condition), pageNumber As Integer, maxPerPage As Integer) As Task(Of Page(Of T)) Implements IOrmAsync.FindByPagedAsync
        Return strategy.FindByPagedAsync(Of T)(conditions, pageNumber, maxPerPage)
    End Function

    Public Function FindByPagedInTableAsync(Of T)(conditions As List(Of Condition), tableName As String, pageNumber As Integer, maxPerPage As Integer) As Task(Of Page(Of T)) Implements IOrmAsync.FindByPagedInTableAsync
        Return strategy.FindByPagedInTableAsync(Of T)(conditions, tableName, pageNumber, maxPerPage)
    End Function

    Public Function UpdateAsync(Of T)(obj As T, Optional idColumn As String = "Id") As Task(Of T) Implements IOrmAsync.UpdateAsync
        Return strategy.UpdateAsync(Of T)(obj, idColumn)
    End Function

    Public Function UpdateInTableAsync(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id") As Task(Of T) Implements IOrmAsync.UpdateInTableAsync
        Return strategy.UpdateInTableAsync(Of T)(obj, tableName, idColumn)
    End Function

    Public Sub PrepareDatabaseAsync(mode As DbPrepMode, entities As List(Of Type), Optional idColumn As String = Id) Implements IOrmAsync.PrepareDatabaseAsync
        strategy.PrepareDatabaseAsync(mode, entities, idColumn)
    End Sub


#End Region

End Class
