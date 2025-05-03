
''' <summary>
''' Object Relational Mapping. Joins and concurrency supported.
''' SqlServer is the only supported database provider, support for PostgreSql and MySql is actively WIP.
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
