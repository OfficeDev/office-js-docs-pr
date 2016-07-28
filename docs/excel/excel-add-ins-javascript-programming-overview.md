# Excel REST API

## Objects 

* Worksheet: The Worksheet object is a member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.
  * Worksheet Collection: A collection of all the Workbook objects that are part of the workbook. 
* Range: Range represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells.  
* Table: Represents collection of organized cells designed to make management of the data easy. 
  * Table Collection: A collection of Tables in a workbook or worksheet. 
  * TableColumn Collection: A collection of all the columns in a Table. 
  * TableRow Collection: A collection of all the rows in a Table. 
* Chart: Represents a chart object in a workbook, which is a visual representation of underlying data.  
  * Chart Collection: A collection of charts in a workbook or a worksheet 
* NamedItem: Represents a defined name for a range of cells or a value. Names can be primitive named objects (as seen in the type below), range object, etc.
  * NamedItem Collection: a collection of named items of a workbook.

Following sections provide important programming details related to Excel REST APIs.

* [Authorization and scopes](#authorization-and-scopes)
* [The basics](#the-basics)
* [Worksheet operations](#worksheet-operations)
* [Chart operations](#chart-operations)
* [Table operations](#table-operations)
* [Range operations](#range-operations)
* [Null-Input](#null-input)
* [Null-Response](#null-response)
* [Blank Input and Output](#blank-input-and-output)
* [Unbounded-Range](#unbounded-range)
* [Large-Range](#large-range)
* [Single Input Copy](#single-input-copy)


### Authorization and scopes

The standard OAuth2 based authorization used across mechanism applies to Excel APIs. All APIs require the `Authorization: Bearer {access-tken}` HTTP header.   
Please refer to the autorization section of the docs to learn more.  
  

##### Scopes
One of the following scopes is required to execute Excel API:

* Files.Read 
* Files.ReadWrite


#### The Basics

Excel REST APIs allow web and mobile applications to read and modify workbook stored on the supported storage platforms (OneDrive, SharePoint, etc.). `Workbook` (or Excel file) is the top level object, which consists of all other Excel objects through relationships. A workbook is addressed through drive API by identifying the location of the file in the URL. Example:

`https://graph.microsoft.com/{ver}/me/drive/items/{id}/workbook/`  
`https://graph.microsoft.com/{ver}/me/drive/root:/{item-path}:/workbook/`  

A set of Excel objects (such as Table, Range, Chart, etc.) could be accessed using standard REST interfaces to perform CRUD (create, read, update, delete) operation on the workbook. For example, 
`https://graph.microsoft.com/{ver}/me/drive/items/{id}/workbook/`  
returns a collection of worksheet objects part of the workbook.    

### Excel Session and persistance

Excel APIs can be called in one of two modes: 

1. Persistent session: In this mode, all changes made to the workbook are persisted (saved). This is the usual mode of operation. 
2. Non-persistent session: In this mode, changes made by the API are not saved to the source location. Instad, Excel backend server keeps a temporary copy of the file that reflects the changes made during that particualr API session. Once the excel session expires, the changes are lost. This mode is useful to apps that may need to do analysis or obtain result of calculatio or a chart image, etc.; at the same time not impact the document state itself.   

Session is represented in the API using `workbook-session-id: {session-id}` header. 

_Is the session header required?_ No. Session header is not required for an Excel API to work. However, using the session is a good practice to get better performance. If no session header is used, change made during the API call _is_ persisted to the file.  

#### API call to get a session. 

##### Request 

Pass a JSON object by setting the `persistchanges` value to `true` or `false`. 

```http
POST /{ver}/me/drive/items/01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN/workbook/CreateSession
content-type: Application/Json 
authorization: Bearer {access-token}
 
{ "persistChanges": true }
```

When the value of `persistChanges` is set to `false`, a non-persistant session id is returned.  


##### Response

```http
HTTP code: 201, Created
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#microsoft.graph.sessionInfo",
  "id": "{session-id}",
  "persistChanges": true
}
```

##### Usage 

Session Id returned from the previous call is passed as a header on subsequent API requests in  
`workbook-session-id` HTTP header. 

```http
GET /{ver}/me/drive/items/01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN/workbook/Worksheets
authorization: Bearer {access-token} 
workbook-session-id: {session-id}
```
[top](#excel-rest-api)

### Worksheet operations

#### List worksheets part of the workbook 
Request 

```http
GET /{ver}/me/drive/items/01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN/workbook/Worksheets
accept: Application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id}
```

Response
 
```http
HTTP code: 200, OK
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN')/workbook/worksheets",
  "value": [
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN')/workbook/worksheets(%27%7B00000000-0001-0000-0000-000000000000%7D%27)",
      "id": "{00000000-0001-0000-0000-000000000000}",
      "name": "Sheet1",
      "position": 0,
      "visibility": "Visible"
    },
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN')/workbook/worksheets(%27%7B00000000-0001-0000-0100-000000000000%7D%27)",
      "id": "{00000000-0001-0000-0100-000000000000}",
      "name": "Sheet57664",
      "position": 1,
      "visibility": "Visible"
    }
  ]
}
```
#### Add a new worksheet 
 
```http
POST /{ver}/me/drive/items/01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN/workbook/Worksheets
content-type: Application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id}

{ "name": "Sheet32243" }
```

Response 
```http
HTTP code: 201, Created
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN')/workbook/worksheets/$entity",
  "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN')/workbook/worksheets(%27%7B75A18F35-34AA-4F44-97CC-FDC3C05D9F40%7D%27)",
  "id": "{75A18F35-34AA-4F44-97CC-FDC3C05D9F40}",
  "name": "Sheet32243",
  "position": 5,
  "visibility": "Visible"
}
```

#### Delete a worksheet

Request
```
DELETE /{ver}/me/drive/items/01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN/workbook/Worksheets('%7B75A18F35-34AA-4F44-97CC-FDC3C05D9F40%7D')
content-type: Application/Json 
```

Response
```http
HTTP code: 204, No Content
```


#### Update worksheet properties

Request 

```
PATCH /{ver}/me/drive/items/01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN/workbook/Worksheets('%7B00000000-0001-0000-0100-000000000000%7D')
content-type: Application/Json 
accept: application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id}

{ "name": "SheetA", "position": 3 }
```

Response
 
```http
HTTP code: 200, OK
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN')/workbook/worksheets/$entity",
  "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJGUJ7JHBSZDFZFL25KSZGQTVAUN')/workbook/worksheets(%27%7B00000000-0001-0000-0100-000000000000%7D%27)",
  "id": "{00000000-0001-0000-0100-000000000000}",
  "name": "SheetA",
  "position": 3,
  "visibility": "Visible"
}
```
[top](#excel-rest-api)

### Chart operations

#### List charts that are part of the worksheet 

Request
```http 
GET /{ver}/me/drive/items/01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL/workbook/Worksheets('%7B00000000-0001-0000-0000-000000000000%7D')/Charts
accept: Application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id} 
```

Response
```http
HTTP code: 200, OK
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL')/workbook/worksheets('%7B00000000-0001-0000-0000-000000000000%7D')/charts",
  "value": [
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL')/workbook/worksheets(%27%7B00000000-0001-0000-0000-000000000000%7D%27)/charts(%27%7B00000000-0008-0000-0100-000003000000%7D%27)",
      "height": 235.5,
      "id": "{00000000-0008-0000-0100-000003000000}",
      "left": 276.0,
      "name": "Chart 2",
      "top": 0.0,
      "width": 401.25
    }
  ]
}
```

#### Get chart image

Request 
```http
GET /{ver}/me/drive/items/01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL/workbook/Worksheets('%7B00000000-0001-0000-0000-000000000000%7D')/Charts('%7B00000000-0008-0000-0100-000003000000%7D')/Image(width=0,height=0,fittingMode='fit')
authorization: Bearer {access-token} 
workbook-session-id: {session-id} 
accept-encoding: gzip;q=1.0,deflate;q=0.6,ident 
```

Response
```http
HTTP code: 200, OK
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#Edm.String",
  "value": "{base-64-string}"
}
```

#### Add a chart  

Request

```http
POST /{ver}/me/drive/items/01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL/workbook/Worksheets('%7B00000000-0001-0000-0000-000000000000%7D')/Charts/Add
content-type: Application/Json 
accept: application/Json 
authorization: Bearer {access-token} 

{ "type": "ColumnClustered", "sourcedata": "A1:C4", "seriesby": "Auto" }
```

Response 
```http
HTTP code: 201, Created
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#chart",
  "@odata.type": "#microsoft.graph.chart",
  "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL')/workbook/worksheets(%27%7B00000000-0001-0000-0000-000000000000%7D%27)/charts(%27%7B2D421098-FA19-41F7-8528-EE7B00E4BB42%7D%27)",
  "height": 216.0,
  "id": "{2D421098-FA19-41F7-8528-EE7B00E4BB42}",
  "left": 0.0,
  "name": "Chart 2",
  "top": 0.0,
  "width": 360.0
}
```

#### Update a chart

````http 
PATCH /{ver}/me/drive/items/01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL/workbook/Worksheets('%7B00000000-0001-0000-0000-000000000000%7D')/Charts('%7B2D421098-FA19-41F7-8528-EE7B00E4BB42%7D')
content-type: Application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id}

{ "height": 216.0, "left": 0, "name": "NewName", "top": 0, "width": 360.0 }

```
Response 

```http
HTTP code: 200, OK
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL')/workbook/worksheets('%7B00000000-0001-0000-0000-000000000000%7D')/charts/$entity",
  "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL')/workbook/worksheets(%27%7B00000000-0001-0000-0000-000000000000%7D%27)/charts(%27%7B2D421098-FA19-41F7-8528-EE7B00E4BB42%7D%27)",
  "height": 216.0,
  "id": "{2D421098-FA19-41F7-8528-EE7B00E4BB42}",
  "left": 0.0,
  "name": "NewName",
  "top": 0.0,
  "width": 360.0
}
```

#### Update chart source data 

Request
```http
POST /{ver}/me/drive/items/01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL/workbook/Worksheets('%7B00000000-0001-0000-0000-000000000000%7D')/Charts('%7B2D421098-FA19-41F7-8528-EE7B00E4BB42%7D')/setData
content-type: Application/Json 
accept: application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id}

{ "sourceData": "A1:C4", "seriesBy": "Auto" }
```

Response
```http
HTTP code: 204, No Content
```

### Table operations 

#### Get list of tables 

Request 
```http
GET /{ver}/me/drive/items/01CYZLFJB6K563VVUU2ZC2FJBAHLSZZQXL/workbook/Worksheets('%7B00000000-0001-0000-0000-000000000000%7D')/Tables
accept: Application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id}
```

Response
```http
HTTP code: 200, OK
content-type: application/json;odata.metadat 
```

#### Update table

Request 
```http 
PATCH /beta/me/drive/items/01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4/workbook/Tables('2')
content-type: Application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id}

{ "name": "NewTableName", "showHeaders": true, "showTotals": false, "style": "TableStyleMedium4" }
```

Response 
```http
HTTP code: 200, OK
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/beta/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables/$entity",
  "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%272%27)",
  "id": "2",
  "name": "NewTableName",
  "showHeaders": true,
  "showTotals": false,
  "style": "TableStyleMedium4"
}
```

#### Get list of table rows
Request 

```http
GET /{ver}/me/drive/items/01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4/workbook/Tables('4')/Rows
authorization: Bearer {access-token} 
workbook-session-id: {session-id}
```

Response

```http
HTTP code: 200, OK
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables('4')/rows",
  "value": [
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/rows/itemAt(0)",
      "index": 0,
      "values": [
        [
          42019,
          53,
          34
        ]
      ]
    },
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/rows/itemAt(1)",
      "index": 1,
      "values": [
        [
          42020,
          45,
          39
        ]
      ]
    },
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/rows/itemAt(2)",
      "index": 2,
      "values": [
        [
          42021,
          50,
          31
        ]
      ]
    },
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/rows/itemAt(3)",
      "index": 3,
      "values": [
        [
          42022,
          43,
          39
        ]
      ]
    },
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/rows/itemAt(4)",
      "index": 4,
      "values": [
        [
          42023,
          45,
          41
        ]
      ]
    },
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/rows/itemAt(5)",
      "index": 5,
      "values": [
        [
          42024,
          52,
          40
        ]
      ]
    }
  ]
}
```

### Get list of table columns

Request
```http
GET /{ver}/me/drive/items/01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4/workbook/Tables('4')/Columns
authorization: Bearer {access-token} 
workbook-session-id: {session-id}
```

Response 

```http
HTTP code: 200, OK 
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/{ver}/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables('4')/columns",
  "value": [
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/columns(%271%27)",
      "id": "1",
      "index": 0,
      "name": "Date",
      "values": [
        [
          "Date"
        ],
        [
          42019
        ],
        [
          42020
        ],
        [
          42021
        ],
        [
          42022
        ],
        [
          42023
        ],
        [
          42024
        ]
      ]
    },
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/columns(%272%27)",
      "id": "2",
      "index": 1,
      "name": "High (F)",
      "values": [
        [
          "High (F)"
        ],
        [
          53
        ],
        [
          45
        ],
        [
          50
        ],
        [
          43
        ],
        [
          45
        ],
        [
          52
        ]
      ]
    },
    {
      "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/columns(%273%27)",
      "id": "3",
      "index": 2,
      "name": "Low (F)",
      "values": [
        [
          "Low (F)"
        ],
        [
          34
        ],
        [
          39
        ],
        [
          31
        ],
        [
          39
        ],
        [
          41
        ],
        [
          40
        ]
      ]
    }
  ]
}
```


#### Add a table row

Request
```http
POST /beta/me/drive/items/01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4/workbook/Tables('4')/Rows
content-type: Application/Json 
authorization: Bearer {access-token} 
workbook-session-id: {session-id}

{ "values": [ [ "Jan-15-2016", "49", "37" ] ], "index": null }
```

Response 
```http
HTTP code: 201, Created
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/beta/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables('4')/rows/$entity",
  "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%274%27)/rows(null)",
  "index": 6,
  "values": [
    [
      "Jan-15-2016",
      49,
      37
    ]
  ]
}

#### Add a table column 

Request 
```http 
POST /beta/me/drive/items/01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4/workbook/Tables('2')/Columns
content-type: Application/Json 
accept: application/Json 


{ "values": [ [ "Status" ], [ "Open" ], [ "Closed" ] ], "index": 2 }
```

Response 

```http 
HTTP code: 201, Created
content-type: application/json;odata.metadat 

{
  "@odata.context": "https://graph.microsoft.com/beta/$metadata#users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables('2')/columns/$entity",
  "@odata.id": "/users('f6d92604-4b76-4b70-9a4c-93dfbcc054d5')/drive/items('01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4')/workbook/tables(%272%27)/columns(%274%27)",
  "id": "4",
  "index": 2,
  "name": "Status",
  "values": [
    [
      "Status"
    ],
    [
      "Open"
    ],
    [
      "Closed"
    ]
  ]
}
```

#### Delete table row

Request 
```http  
DELETE /beta/me/drive/items/01CYZLFJDYBLIGAE7G5FE3I4VO2XP7BLU4/workbook/Tables('4')/Rows/$/ItemAt(index=6)
authorization: Bearer {access-token} 
workbook-session-id: {session-id}
```

Response 
```http
HTTP code: 204, No Content
```

#### delete column 



#### convert table to range 


[top](#excel-rest-api)


### Null-Input

#### null input in 2-D Array

`null` input inside two-dimensional array (for values, number-format, formula) is ignored in the range/table-row/table-column update APIs. No update will take place to the intended target (cell) when `null` input is sent in values or number-format or formula grid of values.

Example: In order to only update specific parts of the Range, such as some cell's Number Format, and to retain the existing number-format on other parts of the Range, set desired Number Format where needed and send `null` for the other cells.

In the set request below, only some parts of the Range Number Format are set while retaining the existing Number Format on the remaining part (by passing nulls).

```json
{
  "values" : [["Eurasia", "29.96", "0.25", "15-Feb" ]],
  "numberFormat" : [[null, null, null, "m/d/yyyy;@"]]
}
```

#### null input for a property

`null` is not a valid single input for the entire property. For example, the following is not valid as the entire values cannot be set to null or ignored.

```json
{
 "values":  null
}

```

The following is not valid either as null is not a valid color value.

```json
{
 "color" =  null
}
```

### Null-Response

Representation of formatting properties that consists of non-uniform values would result in the return of a null value in the response.

Example: A Range can consist of one of more cells. In cases where the individual cells contained in the Range specified don't have uniform formatting values, the range level representation will be undefined.

```json
{
  "size: : null,
  "color" : null
}
```

### Blank Input and Output

Blank values in update requests are treated as instruction to clear or reset the respective property. Blank value is represented by two double quotation marks with no space in-between. `""`

Example:

* For `values`, the range value is cleared out. This is the same as clearing the contents in the application.

* For `numberFormat`, the number format is set to `General`.

* For `formula` and `formulaLocale`, the formula values are cleared.


For read operations, expect to receive blank values if the contents of the cells are blanks. If the cell contains no data or value, then the API returns a blank value. Blank value is represented by two double quotation marks with no space in-between. `""`.

```json
{
  "values" : [["", "some", "data", "in", "other", "cells", ""]]
}
```

```json
{
  "formula" = [["", "", "=Rand()"]]
}
```

### Unbounded Range

#### Read

Unbounded range address contains only column or row identifiers and unspecified row identifier or column identifiers (respectively), such as:

* `C:C`, `A:F`, `A:XFD` (contains unspecified rows)
* `2:2`, `1:4`, `1:1048546` (contains unspecified columns)

When the API makes a request to retrieve an unbounded Range (e.g., `getRange('C:C')`, the response returned contains `null` for cell level properties such as `values`, `text`, `numberFormat`, `formula`, etc.. Other Range properties such as `address`, `cellCount`, etc. will reflect the unbounded range.

#### Write

Setting cell level properties (such as values, numberFormat, etc.) on unbounded Range is **not allowed** as the input request might be too large to handle.

Example: The following is not a valid update request because the requested range is unbounded.

```http
PATCH /workbook/worksheets('Sheet1')/Range(address="A:B")

{
  "values" = 'Due Date'
}
```

When an update operation is attempted on such a Range, the API will return an error.


### Large Range

Large Range implies a Range whose size is too large for a single API call. Many factors such as number of cells, values, numberFormat, formulas, etc. contained in the range can make the response so large that it becomes unsuitable for API interaction. The API makes a best attempt to return or write to the requested data. However, the large size involved might result in an API error condition because of the large resource utilization.

To avoid such a condition, using read or write for large Range in multiple smaller range sizes is recommended.


### Single Input Copy

To support updating a range with the same values or number-format or applying same formula across a range, the following convention is used in the set API. In Excel, this behavior is similar to inputting values or formulas to a range in the CTRL+Enter mode.

The API will look for a *single cell value* and, if the target range dimension doesn't match the input range dimension, it will apply the update to the entire range in the CTRL+Enter model with the value or formula provided in the request.

#### Examples

The following request updates the selected range with the text of "Due Date". Note that Range has 200 cells, whereas the provided input only has 1 cell value.

```js
```http
PATCH /workbook/worksheets('Sheet1')/Range(address="A1:B100")

{
  "values" = 'Due Date'
}
```

The following request updates the selected range with the date of '3/11/2015'.

```js
Excel.run(function (ctx) {
  var sheetName = 'Sheet1';
  var rangeAddress = 'A1:A20';
  var worksheet = ctx.workbook.worksheets.getItem(sheetName);
  var range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
  return ctx.sync().then(function() {
    console.log(range.text);
  });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
The following request updates the selected range with a formula that will be applied across the range in the CTRL+Enter mode.

```js
Excel.run(function (ctx) {
  var sheetName = 'Sheet1';
  var rangeAddress = 'A1:A20';
  var worksheet = ctx.workbook.worksheets.getItem(sheetName);
  var range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
  return ctx.sync().then(function() {
    console.log(range.text);
  });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Error information 


