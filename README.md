# Code Repertoire

<br>
A programmer's toolkit - Code Repertoire is multi-purpose for everyday use in .Net, bring it in via <a href="https://www.nuget.org/packages/inovationware.code/#versions-body-tab" target="_blank">Nuget</a>.

<br>
<br>
The Java version is <a href="https://github.com/tundeadetunji/api-java-code">here</a>.

<br>
<br>
The Android version is <a href="https://github.com/tundeadetunji/api-android-general_module">here</a>.

<br/>
<br>
<br/>
Following are usage examples.

<br/>
<br/>



---
## SequelOrm
Lightweight ORM.
- ﻿SqlServer is currently the only supported database provider; this implementation was tested on version 2016. Support for PostgreSql and MySql are actively WIP.
- Parent-to-child foreign key convention is {ParentTypeName}_{idColumn} (e.g. Person {Id, Name, List(Of Phone)} => Phone {Id, Person_Id, PhoneNumber}).
- Supports 1 : many (e.g. 1 Person can have many Phones); the rest, as well as Logging, Caching and Transaction Isolation Levels, are actively WIP and will be available in the next few weeks.
- Direct support for Enum types is now available.
- Picks up non read-only properties only, and ignores RowVersion property if it doesnt exist.
- Finally, you may need to install appropriate Sql Clients or Providers from nuget.org, e.g. System.Data.SqlClient.

C#:
```csharp
using iNovation.Code;

// prerequisites
SupportedDbProvider provider = SupportedDbProvider.SqlServer; // others are WIP
string connectionString = "my connection string";

// the sync version
IOrm orm = SequelOrm.GetInstance(provider, connectionString);

// create a Person object
List<Phone> amyPhoneNumbers = new List<Phone> { new Phone { PhoneNumber = "01 001 AMY" } };
Person amy = new Person { Name = "Amy", Phones = amyPhoneNumbers, Account = AccountType.Checking };

// create a Person record
Person created = orm.Create<Person>(amy);

// find records of Person by sepcified conditions - the paged version
Condition condition = new Condition { Column = "PhoneNumber", SqlComparison = SqlComparisonOperator.SqlEquals, Value = "01 001 AMY" };
Page<Person> sought = orm.FindByPaged<Person>(new List<Condition> { condition }, 1, 10);
// retrieve the results by
List<Person> found = sought.Records;


// for the async version
IOrmAsync ormAsync = SequelOrmAsync.GetInstance(provider, connectionString);

// just call the async version of the methods, e.g.
Person createdByAsync = await ormAsync.CreateAsync<Person>(amy);


// please read the individual classes' documentations for important information.

// the above is based on these entities:

using iNovation.Code;

public class Person {
public int Id { get; set; }

[SqlType("NVARCHAR(100)")]
public string Name { get; set; }
public List<Phone> Phones { get; set; } = new List<Phone>();

[EnumStorage(EnumStorageType.String)]
public AccountType Account {get; set;} = AccountType.Checking;

public override string ToString()
{
    return Name + " (Id: " + Id + ")";
}

public enum AccountType
{
	Savings = 0,
	Current = 1,
	Checking = 2
}

public class Phone {
    public int Id { get; set; }
    public int Person_Id { get; set; }
    public string PhoneNumber { get; set; }
    public override string ToString()
    {
        return PhoneNumber;
    }
}
```
VB:
```vbnet
Imports iNovation.Code

' prerequisites
Dim provider As SupportedDbProvider = SupportedDbProvider.SqlServer ' other providers are WIP
Dim connectionString As String = "my connection string"


' the sync version
Dim orm As IOrm = SequelOrm.GetInstance(provider, connectionString)

' create a Person object
Dim amyPhoneNumbers As New List(Of Phone) From {New Phone With {.PhoneNumber = "01 001 AMY"}}
Dim amy As New Person With {.Name = "Amy", .Phones = amyPhoneNumbers, .Account=AccountType.Checking}


' create a Person record
Dim created As Person = orm.Create(Of Person)(amy)

' find records of Person by specified conditions - the paged version
Dim conditions As New List(Of Condition) From {New Condition With {.Column = "PhoneNumber", .SqlComparison = SqlComparisonOperator.SqlEquals, .Value = "01 001 AMY"}}
Dim sought As Page(Of Person) = orm.FindByPaged(Of Person)(conditions, 1, 10)
' retrieve the results by
Dim found As List(Of Person) = sought.Records


' for the async version
Dim ormAsync As IOrmAsync = SequelOrmAsync.GetInstance(provider, connectionString)

' just call the async version of the methods, e.g.
Dim personCreatedByAsync As Person = Await ormAsync.CreateAsync(Of Person)(amy)


' please read the individual classes' documentations for important information.

' the above is based on these entities:
Imports iNovation.Code

Public Class Person
    Public Property Id As Integer

    <SqlType("NVARCHAR(100)")>
    Public Property Name As String
    Public Property Phones As List(Of Phone) = New List(Of Phone)
    <EnumStorage(EnumStorageType.String)>
    Public Property Account As AccountType = AccountType.Checking

    Public Overrides Function ToString() As String
        Return Name & " (Id: " & Id & ")"
    End Function

End Class

Public Enum AccountType
    Savings = 0
    Current = 1
    Checking = 2
End Enum

Public Class Phone
    Public Property Id As Integer
    Public Property Person_Id As Integer
    Public Property PhoneNumber As String
    Public Overrides Function ToString() As String
        Return PhoneNumber
    End Function

End Class
```

---
## Bootstrap
Contains methods based on Bootstrap, currently 4 but upgrades are WIP.

C#:
```csharp
using static iNovation.Code.Bootstrap

Alert("This is an alert");
```
VB:
```vbnet
Imports iNovation.Code.Bootstrap

Alert("This is an alert")
```

---
## Charts
Contains methods for dynamically creating charts (Bar, Pie, Doughnut, Line) using given values or directly from database columns.

C#:
```csharp
using static iNovation.Code.Charts

List<string> labels = new List<string> { "Q1", "Q2", "Q3" };
List<string> values  = new List<string> { "50", "25", "75" };
string html = BarChart(labels, values);

```
VB:
```vbnet
Imports iNovation.Code.Charts

Dim labels As New List(Of String) From {"Q1", "Q2", "Q3"}
Dim values As New List(Of String) From {"50", "25", "75"}
Dim html As String = BarChart(labels, values)
```

---
## Desktop
Contains methods mainly towards desktop development.

C#:
```csharp
using static iNovation.Code.Desktoop

enum TitleOfCourtesy{
    Mr, Mrs, Ms
}

//binds values of TitleOfCourtesy to comboBox1
EnumDrop(comboBox1, new TitleOfCourtesy());
```
VB:
```vbnet
Imports iNovation.Code.Desktop

Enum TitleOfCourtesy
    Mr
    Mrs
    Ms
End Enum

' binds values of TitleOfCourtesy to ComboBox1
EnumDrop(ComboBox1, New TitleOfCourtesy())
```

---
## DesktopExtensions
Contains extension methods based on methods from Desktop.

C#:
```csharp
using static iNovation.Code.DesktopExtensions

//The items of comboBox1 are turned into a List
comboBox1.ToList();
```
VB:
```vbnet
Imports iNovation.Code.DesktopExtensions

' The items of ComboBox1 are turned into a List
ComboBox1.ToList()
```

---
## Encryption
Ligthweight Encryption/Decryption.

C#:
```csharp
using static iNovation.Code.Encryption

string key = "MyKey";
string value = "What to encrypt";
string encrypted = Encrypt(key, value);

string original = Decrypt(key, encrypted);
```
VB:
```vbnet
Imports iNovation.Code.Encryption

Dim key As String = "MyKey"
Dim value = "What to encrypt"
Dim encrypted As String = Encrypt(key, value)

Dim original As String = Decrypt(key, encrypted)
```

---
## EncryptionExtensions
Contains extension methods based on methods from Encryption.

C#:
```csharp
using static iNovation.Code.EncryptionExtensions

string key = "MyKey";
string value = "What to encrypt";
string encrypted = value.Encrypt(key);

string original = encrypted.Decrypt(key);
```
VB:
```vbnet
Imports iNovation.Code.EncryptionExtensions

Dim key As String = "MyKey"
Dim value = "What to encrypt"
Dim encrypted As String = value.Encrypt(key)

Dim original As String = value.Decrypt(key)
```

---
## Feedback
Contains routines for giving feedback with Text-To-Speech and MessageBox.

C#:
```csharp
using iNovation.Code.Feedback

//Text to speech
string s = "Hi";
Feedback feedback = new Feedback();
feedback.Inform(s);
```
VB:
```vbnet
Imports iNovation.Code.Feedback

' Text to speech
Dim s As String = "Hi"
Dim f As New Feedback()
f.Inform(s)
```

---
## FeedbackExtensions
Contains extension methods based on methods from Feedback.

C#:
```csharp
using iNovation.Code.FeedbackExtensions

//Text to speech
string s = "Hi";
s.Inform();
```
VB:
```vbnet
Imports iNovation.Code.FeedbackExtensions

' Text to speech
Dim s As String = "Hi"
s.Inform()
```

---
## General
Contains methods for general purposes.

C#:
```csharp
using static iNovation.Code.General

string email = "@provider.com";
bool valid = IsEmail(email); //false
```
VB:
```vbnet
Imports iNovation.Code.General

Dim email As String = "@provider.com"
Dim valid = IsEmail(email) ' False
```

---
## GeneralExtensions
Contains extension methods based on methods from General.

C#:
```csharp
using static iNovation.Code.GeneralExtensions

string input = "how Are you doing";
string output = input.ToTitleCase(); //How Are You Doing
```
VB:
```vbnet
Imports iNovation.Code.GeneralExtensions

Dim input As String = "how Are you doing"
Dim output = input.ToTitleCase() ' How Are You Doing
```

---
## LoggingExtensions
Contains extension methods useful for logging common data types.

C#:
```csharp
using static iNovation.Code.LoggingExtensions

string input = "just a log";
input.Log();
```
VB:
```vbnet
Imports iNovation.Code.LoggingExtensions

Dim input As String = "just a log"
input.Log()
```

---
## Machine
Contains methods for dealing with the machine directly.

C#:
```csharp
using static iNovation.Code.Machine

//Mutes the PC
Mute(this);
```
VB:
```vbnet
Imports iNovation.Code.Machine

' Mutes the PC
Mute(Me)
```

---
## Sequel
Contains methods for database access.

C#:
```csharp
using static iNovation.Code.Sequel

string table = "TableName";
string[] fields = { "ColumnName" };
string[] where_params = { "Id" };
object[] where_params_values = { "Id", 10 };
string connection_string = "Connection_String";

//creates select query
string query = iNovation.Code.General.BuildSelectString(table, fields, where_params);

//creates a DataTable
DataTable dataTable = QDataTable(query, connection_string, where_params_values);
```
VB:
```vbnet
Imports iNovation.Code.Sequel

Dim table As String = "TableName"
Dim fields As String() = {"ColumnName"}
Dim where_params As String() = {"Id"}
Dim where_params_values As Object() = {"Id", 10}
Dim connection_string As String = "Connection_String"

' creates select query
Dim query As String = iNovation.Code.General.BuildSelectString(table, fields, where_params)

' creates a DataTable
Dim d As DataTable = QDataTable(query, connection_string, where_params_values)
```

---
## Styler
Contains methods for desktop development, particularly, styling the Form.

C#:
```csharp
using static iNovation.Code.Styler

Style(this, true);
```
VB:
```vbnet
Imports iNovation.Code.Styler

Style(Me, True)
```


