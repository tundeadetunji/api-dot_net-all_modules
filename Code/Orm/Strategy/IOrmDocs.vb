
''' <summary>
''' Object Relational Mapping. Joins and transactions supported.
''' Uses Lazy Intialization for Singletons.
''' SqlServer is currently the only supported database provider.
''' Support for PostgreSql and MySql, as well as Logging and Transaction Isolation Levels, is actively <b>WIP</b>>.
''' Parent-to-child foreign key convention is {ParentTypeName}_{idColumn} (e.g. Person {Id, Name, List(Of Phone)} => Phone {Id, Person_Id, PhoneNumber}).
''' Supports 1 : many (e.g. 1 Person can have many Phones).
''' <br />
''' Picks up only non read-only properties only, and ignores RowVersion property if it doesnt exist.
''' <br />
''' You may need to install appropriate Sql Clients or Providers from nuget.org, e.g. System.Data.SqlClient.
''' </summary>
''' <remarks>
''' Author: Tunde Adetunji (tundeadetunji2017@gmail.com)
''' Date: May 3, 2025
''' </remarks>
Public Interface IOrmDocs

End Interface
