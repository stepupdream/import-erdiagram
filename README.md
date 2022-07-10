# Import ERDiagram
[![Latest Stable Version](https://img.shields.io/badge/stable-v1.0.0-blue)](https://img.shields.io/badge/stable-v1.0.0-blue)

## Introduction

Read ER diagram and output table definition to Google Spreadsheet

## Features

- Read the ER diagram written in UML and output it to Google Spreadsheet
- The output is fast because it is made lightweight

## Usage
1. Open the Script Editor of Google Spreadsheet and paste the source code.  
   The only difference between en.js and jp.js is the comment part. Please paste either one.
2. Follow the sample description method below and create an ER diagram with UML
3. Enter the path of Google Drive where UML is saved in the **** part of the source code
4. Reads the one that matches the ○ part of <<★★★★, ○○○○>> and the title in Spreadsheet.  
   In the case of the example, please use the title name "User".
5. You can output the contents of UML by executing Script

## Example UML

```
@startuml

'Hide division ruled lines
hide empty members

'Definition of mark and background color
!define master_data_db #E2EFDA-C6E0B4
!define user_db #FCE4D6-F8CBAD
!define MASTER_DATA AAFFAA
!define User FFAA00

'Define setting color
skinparam class {
    BorderColor Black
    ArrowColor Black
}

package "user" as user {
    entity "accounts [AccountsTable]" as accounts <<U, User>> user_db {
        + id : bigInteger [ID]
        # user_id : bigInteger  [UserID]
    }

    entity "users [UsersTable]" as users <<U, User>> user_db {
        + id : bigInteger [ID]
        name : string [UserName]
    }

    entity "country_names [Country name]" as country_names <<M, MASTER_DATA>> master_data_db {
        + id : integer [ID]
        name : string [Name]
    }
}

'Define relationship diagram
users        ||-up-||     accounts

'memo
note right of country_names : comment

@enduml
```

## Writing Rules

```
entity "'TableName' ['TableDescription']" as 'TableName' <<'CategoryMark', 'Category'>> 'ConnectionName' {
  'ColumnName' : 'DataType' ['ColumnDescription']
}
```

## Contributing
Please see [CONTRIBUTING](https://github.com/stepupdream/import-erdiagram/blob/main/.github/CONTRIBUTING.md) for details.
 
## License

The Spreadsheet converter is open-sourced software licensed under the [MIT license](https://choosealicense.com/licenses/mit/)
