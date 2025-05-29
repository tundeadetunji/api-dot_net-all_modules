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
Public Interface IOrmDocs

End Interface
