This is a [Koa.js](http://koajs.com/) web server which communicates with a [Mono-based](http://www.mono-project.com/) CLI program to wrap [SpreadsheetGear](http://www.spreadsheetgear.com/) to process Excel spreadsheets in Linux. IPC is done using [ZeroMQ](http://zeromq.org).

### Pre-requisites
Pkg-config, libsodium and native ZeroMQ libraries might need to be installed before running this. On dev machines (Macs) you can try using `brew`:

```
brew install pkg-config
brew install libsodium
brew install zeromq
```

### How to use
Run this via docker. e.g. 

```
docker build -t hsce:0.1 .
docker run -p 3000:3000 -d hsce:0.1
```

Nothing is persisted.

### Git Hooks
Be sure to symlink JS lint

```
ln -s -f src/node/hooks/pre-commit .git/hooks/pre-commit
```

### Known Issues
* **At the moment this repo only works with the embedded spreadsheet "Testfilev2.xlsx"**

### HTTP/REST JSON Requests API
Two paths are available, details and evaluate

#### Details
Sample output for `GET http://192.168.33.100:3000/api/details`

```json
{
  "id": "TestFilev2.xlsx",
  "purpose": "test spreadsheet for POC",
  "inputs": {
    "input": [
      {
        "Name": "3",
        "FgColor": "#604A7B",
        "BgColor": "#FFFFFF",
        "ValueType": "",
        "Value": "3",
        "DataType": "Number",
        "HTMLDataType": "number"
      },
      {
        "Name": "Boolean",
        "FgColor": "#000000",
        "BgColor": "#FFFFFF",
        "ValueType": "",
        "Value": "TRUE",
        "DataType": "Logical",
        "HTMLDataType": "checkbox"
      },
      {
        "Name": "CheckNumberRange",
        "FgColor": "#000000",
        "BgColor": "#FFFFFF",
        "ValueType": "WholeNumber",
        "Value": "5",
        "DataType": "Number",
        "HTMLDataType": "number"
      },
      {
        "Name": "ComplexValidation",
        "FgColor": "#006100",
        "BgColor": "#C6EFCE",
        "ValueType": "Custom",
        "Value": "2342",
        "DataType": "Number",
        "HTMLDataType": "number"
      },
      {
        "Name": "DateInput",
        "FgColor": "#000000",
        "BgColor": "#FFFFFF",
        "ValueType": "Date",
        "Value": "2016-06-03",
        "DataType": "Date",
        "HTMLDataType": "date"
      },
      {
        "Name": "FirstInput",
        "FgColor": "#000000",
        "BgColor": "#FFFFFF",
        "ValueType": "",
        "Value": "1",
        "DataType": "Number",
        "HTMLDataType": "number"
      },
      {
        "Name": "WithListofValues",
        "FgColor": "#000000",
        "BgColor": "#FFFFFF",
        "ValueType": "List",
        "Value": "Yellow",
        "DataType": "Text",
        "HTMLDataType": "text"
      },
      {
        "Name": "WithSomeColour",
        "FgColor": "#000000",
        "BgColor": "#FFFF00",
        "ValueType": "",
        "Value": "2",
        "DataType": "Number",
        "HTMLDataType": "number"
      }
    ]
  }
}
```

#### Evaluate
Sample output for `POST http://192.168.33.100:3000/api/evaluate`
##### Body
```json
{
  "data": [
    {
      "Name": "3",
      "FgColor": "#604A7B",
      "BgColor": "#FFFFFF",
      "ValueType": "",
      "Value": "18",
      "DataType": "Number",
      "HTMLDataType": "number"
    },
    {
      "Name": "Boolean",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "",
      "Value": "TRUE",
      "DataType": "Logical",
      "HTMLDataType": "checkbox"
    },
    {
      "Name": "CheckNumberRange",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "WholeNumber",
      "Value": "5",
      "DataType": "Number",
      "HTMLDataType": "number"
    },
    {
      "Name": "ComplexValidation",
      "FgColor": "#006100",
      "BgColor": "#C6EFCE",
      "ValueType": "Custom",
      "Value": "2342",
      "DataType": "Number",
      "HTMLDataType": "number"
    },
    {
      "Name": "DateInput",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "Date",
      "Value": "2016-06-03",
      "DataType": "Date",
      "HTMLDataType": "date"
    },
    {
      "Name": "FirstInput",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "",
      "Value": "1",
      "DataType": "Number",
      "HTMLDataType": "number"
    },
    {
      "Name": "WithListofValues",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "List",
      "Value": "Yellow",
      "DataType": "Text",
      "HTMLDataType": "text"
    },
    {
      "Name": "WithSomeColour",
      "FgColor": "#000000",
      "BgColor": "#FFFF00",
      "ValueType": "",
      "Value": "2",
      "DataType": "Number",
      "HTMLDataType": "number"
    }
  ]
}
```

##### Result
```json
{
  "log": {
    "EngineID": "6DF4ACEF3B91003E2794ABF0176985AC",
    "ServiceName": "CalcService-TestFileV2",
    "ServiceRevision": "1.17b",
    "TimeStamp": "11/25/2016 03:49:16",
    "Xinputs": {
      "3": "18",
      "Boolean": "FALSE",
      "CheckNumberRange": "5",
      "ComplexValidation": "2342",
      "DateInput": "2016-06-03",
      "FirstInput": "1",
      "WithListofValues": "Yellow",
      "WithSomeColour": "2"
    },
    "Xoutputs": {
      "SomeOther": "84312 concatenated with Yellow",
      "Sum": "2368",
      "TestingARange": "[{\"Column1\":\"Xname\",\"Column2\":\"CalcService-TestFileV2\"},{\"Column1\":\"XdoNotLog\",\"Column2\":\"False\"},{\"Column1\":\"Xnotes\",\"Column2\":\"This is the notes field ? in testfilev2 ?. \"},{\"Column1\":\"XlogMessage\",\"Column2\":\"calling a fuction: 42699.1592211688\"},{\"Column1\":\"XerrorMessage\",\"Column2\":\"\"},{\"Column1\":\"Xauthor\",\"Column2\":\"Peter Roschke\"},{\"Column1\":\"XreleaseDate\",\"Column2\":\"42622\"}]",
      "TableTest": "[{\"Name\":\"Anya\",\"Age\":12.0,\"AOBPIAH\":2.333},{\"Name\":\"Brent\",\"Age\":10.0,\"AOBPIAH\":7.814},{\"Name\":\"Bob\",\"Age\":24.0,\"AOBPIAH\":43.21},{\"Name\":\"Bobby\",\"Age\":36.0,\"AOBPIAH\":2.01114},{\"Name\":\"Bobbers\",\"Age\":44.0,\"AOBPIAH\":2.3312}]"
    },
    "SourceSystem": "/tmp/spreadsheets/Testfilev2.xlsx",
    "Tags": "",
    "LogMessage": "calling a fuction: 42699.1592211688",
    "ErrorMessage": "",
    "CalcTime": 0.048100000000000004,
    "CallPurpose": "Test Purpose",
    "DictXoutputs": {
      "SomeOther": "84312 concatenated with Yellow",
      "Sum": "2368",
      "TestingARange": "[{\"Column1\":\"Xname\",\"Column2\":\"CalcService-TestFileV2\"},{\"Column1\":\"XdoNotLog\",\"Column2\":\"False\"},{\"Column1\":\"Xnotes\",\"Column2\":\"This is the notes field ? in testfilev2 ?. \"},{\"Column1\":\"XlogMessage\",\"Column2\":\"calling a fuction: 42699.1592211688\"},{\"Column1\":\"XerrorMessage\",\"Column2\":\"\"},{\"Column1\":\"Xauthor\",\"Column2\":\"Peter Roschke\"},{\"Column1\":\"XreleaseDate\",\"Column2\":\"42622\"}]",
      "TableTest": "[{\"Name\":\"Anya\",\"Age\":12.0,\"AOBPIAH\":2.333},{\"Name\":\"Brent\",\"Age\":10.0,\"AOBPIAH\":7.814},{\"Name\":\"Bob\",\"Age\":24.0,\"AOBPIAH\":43.21},{\"Name\":\"Bobby\",\"Age\":36.0,\"AOBPIAH\":2.01114},{\"Name\":\"Bobbers\",\"Age\":44.0,\"AOBPIAH\":2.3312}]"
    }
  },
  "input": [
    {
      "Name": "3",
      "FgColor": "#604A7B",
      "BgColor": "#FFFFFF",
      "ValueType": "",
      "Value": "18",
      "DataType": "Number",
      "HTMLDataType": "number"
    },
    {
      "Name": "Boolean",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "",
      "Value": "FALSE",
      "DataType": "Logical",
      "HTMLDataType": "checkbox"
    },
    {
      "Name": "CheckNumberRange",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "WholeNumber",
      "Value": "5",
      "DataType": "Number",
      "HTMLDataType": "number"
    },
    {
      "Name": "ComplexValidation",
      "FgColor": "#006100",
      "BgColor": "#C6EFCE",
      "ValueType": "Custom",
      "Value": "2342",
      "DataType": "Number",
      "HTMLDataType": "number"
    },
    {
      "Name": "DateInput",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "Date",
      "Value": "2016-06-03",
      "DataType": "Date",
      "HTMLDataType": "date"
    },
    {
      "Name": "FirstInput",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "",
      "Value": "1",
      "DataType": "Number",
      "HTMLDataType": "number"
    },
    {
      "Name": "WithListofValues",
      "FgColor": "#000000",
      "BgColor": "#FFFFFF",
      "ValueType": "List",
      "Value": "Yellow",
      "DataType": "Text",
      "HTMLDataType": "text"
    },
    {
      "Name": "WithSomeColour",
      "FgColor": "#000000",
      "BgColor": "#FFFF00",
      "ValueType": "",
      "Value": "2",
      "DataType": "Number",
      "HTMLDataType": "number"
    }
  ],
  "output": [
    {
      "Name": "SomeOther",
      "Value": "84312 concatenated with Yellow"
    },
    {
      "Name": "Sum",
      "Value": "2368"
    },
    {
      "Name": "TestingARange",
      "Value": "[{\"Column1\":\"Xname\",\"Column2\":\"CalcService-TestFileV2\"},{\"Column1\":\"XdoNotLog\",\"Column2\":\"False\"},{\"Column1\":\"Xnotes\",\"Column2\":\"This is the notes field ? in testfilev2 ?. \"},{\"Column1\":\"XlogMessage\",\"Column2\":\"calling a fuction: 42699.1592211688\"},{\"Column1\":\"XerrorMessage\",\"Column2\":\"\"},{\"Column1\":\"Xauthor\",\"Column2\":\"Peter Roschke\"},{\"Column1\":\"XreleaseDate\",\"Column2\":\"42622\"}]"
    },
    {
      "Name": "TableTest",
      "Value": "[{\"Name\":\"Anya\",\"Age\":12.0,\"AOBPIAH\":2.333},{\"Name\":\"Brent\",\"Age\":10.0,\"AOBPIAH\":7.814},{\"Name\":\"Bob\",\"Age\":24.0,\"AOBPIAH\":43.21},{\"Name\":\"Bobby\",\"Age\":36.0,\"AOBPIAH\":2.01114},{\"Name\":\"Bobbers\",\"Age\":44.0,\"AOBPIAH\":2.3312}]"
    }
  ],
  "charts": []
}
```
