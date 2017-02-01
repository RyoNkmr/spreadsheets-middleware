# spreadsheets-middleware
express middleware

## setting

```javascript
// app.js

var express = require('express');
var app = express();

var spreadsheetsMiddleware = require('spreadsheets-middleware');

var spreadsheetSettings = {
 // ex.) https://docs.google.com/spreadsheets/d/HOGEHOGEHOGEHOGEHOGE/edit
  sheetId: 'sheetid is HOGEHOGEHOGEHOGEHOGE',
  //and you must get creds for spreadsheets API at https://console.developers.google.com/
  privateKey: '-----BEGIN PRIVATE KEY-----\nYOU_MUST_GET_THIS_KEY_FROM_CONSOLE_DEVELOPERS_GOOGLE_COM_FOR_EXAMPLE_YOU_ADD_A_PROJECT_AND_ENABLE_GOOGLE_DRIVE_API_AND_THEN_GENERATE_SERVICE_ACCOUNT_WITH_KEY_FILE_AS_JSON_THEN_OPEN_THAT_JSON_FILE_CONTAINS_PRIVATE_KEY_LIKE_THIS!!_\n-----END PRIVATE KEY-----\n',
  clientEmail: 'this.is.also.included@in.secrets.json.file'
};

app.use('/spreadsheets', spreadsheetsMiddlware(spreadsheetSettings));
```

## usage
* [GET] ```/spreadsheets/test``` - test connecting and work with sheetAPI.

    1. get information
    1. add a worksheet
    1. change title
    1. set header row
    1. get rows 
    1. resize worksheet
    1. add rows
    1. delete worksheet

```javascript
// no need for Request body nor params

// but response like
{
  setTitle: true,
  setHeaderRow: true,
  getRows: true,
  resize: true,
  addRow: true,
  del: true
}
```

* [POST] ```/spreadsheets``` - add data to worksheet(*children of sheet*), ```_id``` stands worksheet title
```javascript
// requestBody
{
  _id: 'workSheetTitle', // --- required
  col1: 'data1',
  col2: 'data2',
  col3: 'data3',
  col4: 'data4',
  ...
}
```


* [POST] ```/spreadsheets/:worksheetTitle``` - add data to worksheet
```javascript
// requestBody
{
  //_id: 'workSheetTitle', --- ignored
  col1: 'data1',
  col2: 'data2',
  col3: 'data3',
  col4: 'data4',
  ...
}
```

In both post method endpoints, add a row on worksheet which has the same name. if no sheet exists has the same name, adding new worksheet to sheet named :worksheetTitle.


