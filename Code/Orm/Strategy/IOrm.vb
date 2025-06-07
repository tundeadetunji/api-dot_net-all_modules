''' <summary>
''' Object Relational Mapping.
''' SqlServer is currently the only supported database provider; this implementation was tested on it.
''' <br />
''' Uses Lazy Intialization for Singletons.
''' <br />
''' Support for PostgreSql and MySql, as well as Logging and Transaction Isolation Levels, is actively <b>WIP</b>>.
''' <br />
''' Parent-to-child foreign key convention is {ParentTypeName}_{idColumn} (e.g. Person {Id, Name, List(Of Phone)} => Phone {Id, Person_Id, PhoneNumber}).
''' <br />
''' Supports 1 : many (e.g. 1 Person can have many Phones); the rest are actively <b>WIP</b>.
''' <br />
''' Picks up non read-only properties only, and ignores RowVersion property if it doesnt exist.
''' <br />
''' You may need to install appropriate Sql Clients or Providers from nuget.org, e.g. System.Data.SqlClient.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: May 30, 2025
''' </remarks>
Public Interface IOrm
    Inherits IOrmDocs

#Region "PrepareDatabase"
    ''' <summary>
    ''' Initializes the database.
    ''' Uses Transaction.
    ''' </summary>
    ''' <param name="mode"></param>
    ''' <param name="entities"></param>
    ''' <param name="idColumn"></param>
    Sub PrepareDatabase(mode As DbPrepMode, entities As List(Of Type), Optional idColumn As String = "Id", Optional isolation As IsolationLevel = IsolationLevel.ReadCommitted)

#End Region

#Region "Create"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' Will insert new record regardlesss. To create only if it doesn't exist, use CreateOrUpdate().
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Function Create(Of T)(obj As T, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As T

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' Will insert new record regardlesss. To create only if it doesn't exist, use CreateOrUpdate().
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Function CreateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As T

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' Will insert new record regardlesss. To create only if it doesn't exist, use CreateOrUpdate().
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="objs"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Function Create(Of T)(objs As List(Of T), Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As List(Of T)

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' Will insert new record regardlesss. To create only if it doesn't exist, use CreateOrUpdate().
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="objs"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Function CreateInTable(Of T)(objs As List(Of T), tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As List(Of T)
#End Region

#Region "Exists"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Function ExistsById(Of T)(id As Object, Optional idColumn As String = "Id") As Boolean

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <param name="id"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Function ExistsByIdInTable(id As Object, tableName As String, Optional idColumn As String = "Id") As Boolean

#End Region

#Region "ExistsBy"
    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <returns></returns>
    Function ExistsBy(Of T)(conditions As List(Of Condition)) As Boolean

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Function ExistsByInTable(Of T)(conditions As List(Of Condition), tableName As String) As Boolean

#End Region

#Region "Delete"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="cascade">When set to true, deletes all child tables when parent table is deleted (may throw exception if database or table structure does not support it, especially when set to false)</param>
    Sub DeleteAll(Of T)(Optional cascade As Boolean = True)

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="tableName"></param>
    ''' <param name="cascade">When set to true, deletes all child tables when parent table is deleted (may throw exception if database or table structure does not support it, especially when set to false)</param>
    Sub DeleteAllInTable(Of T)(tableName As String, Optional cascade As Boolean = True)

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="cascade">When set to true, deletes all child tables when parent table is deleted (may throw exception if database or table structure does not support it, especially when set to false)</param>
    Sub DeleteWhereId(Of T)(id As Object, Optional idColumn As String = "Id", Optional cascade As Boolean = True)

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="cascade">When set to true, deletes all child tables when parent table is deleted (may throw exception if database or table structure does not support it, especially when set to false)</param>
    Sub DeleteWhereIdInTable(Of T)(id As Object, tableName As String, Optional idColumn As String = "Id", Optional cascade As Boolean = True)
#End Region

#Region "DeleteBy"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="cascade">When set to true, deletes all child tables when parent table is deleted (may throw exception if database or table structure does not support it, especially when set to false)</param>
    Sub DeleteBy(Of T)(conditions As List(Of Condition), Optional cascade As Boolean = True)

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="tableName"></param>
    ''' <param name="cascade">When set to true, deletes all child tables when parent table is deleted (may throw exception if database or table structure does not support it, especially when set to false)</param>
    Sub DeleteByInTable(Of T)(conditions As List(Of Condition), tableName As String, Optional cascade As Boolean = True)
#End Region

#Region "CountBy"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <returns></returns>
    Function CountBy(Of T)(conditions As List(Of Condition)) As Long

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Function CountByInTable(Of T)(conditions As List(Of Condition), tableName As String) As Long

#End Region

#Region "Count"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <returns></returns>
    Function Count(Of T)() As Long

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Function CountInTable(Of T)(tableName As String) As Long

#End Region

#Region "FindById"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Function FindById(Of T)(id As Object, Optional idColumn As String = "Id") As T

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="id"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Function FindByIdInTable(Of T)(id As Object, tableName As String, Optional idColumn As String = "Id") As T

#End Region

#Region "FindAll"
    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="idColumn"></param>
    ''' <param name="ascending"></param>
    ''' <returns></returns>
    Function FindAll(Of T)(Optional idColumn As String = "Id", Optional ascending As Boolean = True) As List(Of T)

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Function FindAllInTable(Of T)(tableName As String, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As List(Of T)

#End Region

#Region "FindAllPaged"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="pageNumber"></param>
    ''' <param name="maxPerPage"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="ascending"></param>
    ''' <returns></returns>
    Function FindAllPaged(Of T)(pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Page(Of T)

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="pageNumber"></param>
    ''' <param name="maxPerPage"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="ascending"></param>
    ''' <returns></returns>
    Function FindAllPagedInTable(Of T)(tableName As String, pageNumber As Integer, maxPerPage As Integer, Optional idColumn As String = "Id", Optional ascending As Boolean = True) As Page(Of T)

#End Region

#Region "FindBy"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <returns></returns>
    Function FindBy(Of T)(conditions As List(Of Condition)) As List(Of T)

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Function FindByInTable(Of T)(conditions As List(Of Condition), tableName As String) As List(Of T)

#End Region

#Region "FindByPaged"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="pageNumber"></param>
    ''' <param name="maxPerPage"></param>
    ''' <returns></returns>
    Function FindByPaged(Of T)(conditions As List(Of Condition), pageNumber As Integer, maxPerPage As Integer) As Page(Of T)

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="conditions"></param>
    ''' <param name="tableName"></param>
    ''' <param name="pageNumber"></param>
    ''' <param name="maxPerPage"></param>
    ''' <returns></returns>
    Function FindByPagedInTable(Of T)(conditions As List(Of Condition), tableName As String, pageNumber As Integer, maxPerPage As Integer) As Page(Of T)
#End Region

#Region "Update"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Function Update(Of T)(obj As T, Optional idColumn As String = "Id") As T

    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <returns></returns>
    Function UpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id") As T
#End Region

#Region "CreateOrUpdate"

    ''' <summary>
    ''' Use this if the name of the class is the same as the name of the table.
    ''' Creates record if it doesn't exist, Updates otherwise.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Function CreateOrUpdate(Of T)(obj As T, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As T
    ''' <summary>
    ''' Use this if the name of the class is not the same as the name of the table.
    ''' Creates record if it doesn't exist, Updates otherwise.
    ''' Uses Transaction.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="tableName"></param>
    ''' <param name="idColumn"></param>
    ''' <param name="IdWillAutoIncrement"></param>
    ''' <returns></returns>
    Function CreateOrUpdateInTable(Of T)(obj As T, tableName As String, Optional idColumn As String = "Id", Optional IdWillAutoIncrement As Boolean = True) As T
#End Region

End Interface
