# omnis-xlsx

An Omnis Studio JavaScript Worker module built on top of [SheetJS](https://sheetjs.com) for reading and writing spreadsheet files (Microsoft Excel, LibreOffice/OpenOffice).

## API

The module exposes four methods:

| Method                   | Description                                                                              |
|--------------------------|------------------------------------------------------------------------------------------|
| `write(writeParam)`      | Synchronously writes a spreadsheet file.                                                 |
| `read(readParam)`        | Synchronously reads a spreadsheet file and returns its contents.                         |
| `writeAsync(writeParam)` | Asynchronously writes a spreadsheet file using Node.js promises.                         |
| `readAsync(readParam)`   | Asynchronously reads a spreadsheet file using Node.js promises and returns its contents. |

Non-blocking methods (async) are recommended for large files.

### `writeParam`

| Property      | Type              | Required | Description                                                                                                           |
|---------------|-------------------|:--------:|-----------------------------------------------------------------------------------------------------------------------|
| `filename`    | `string`          |    ✓     | Full path of the spreadsheet file to create.                                                                          |
| `data`        | `Array<Array<*>>` |    ✓     | Data to write. Each nested array represents a row.                                                                    |
| `sheetName`   | `string`          |          | Worksheet name. Default: `Sheet1`.                                                                                    |
| `dateIndexes` | `Array<Object>`   |          | Columns to convert to Excel dates (e.g. `[{ col:1, format:"dd/mm/yyyy" }, { col:3, format:"dd/mm/yyyy hh:mm:ss" }]`). |
| `rowHeader`   | `Array<string>`   |          | Header row written before the data.                                                                                   |

### `readParam`

| Property    | Type     | Required | Description                                                              |
|-------------|----------|:--------:|--------------------------------------------------------------------------|
| `filename`  | `string` |    ✓     | Full path of the spreadsheet file to read.                               |
| `sheetName` | `string` |          | Worksheet to read. If omitted or not found, the first worksheet is used. |

#### Dates

Date values must be provided as ISO 8601 strings (for example, using Omnis `omnisToISO8601`).
The date and time are interpreted exactly as provided, with no UTC conversion or timezone adjustment, and are converted directly to the corresponding Excel serial date value.


## Usage

Create an object (e.g. oXlsx) whose superclass is `.OW3.OW3 Worker Objects\JAVASCRIPTWorker`.

### Write

```
# local variables
#   - loXlsx        Object     Subtype oXlsx
#   - lstData       List
#   - lstHeaders    List

Do loXlsx.$init() Returns #F
Do loXlsx.$start() Returns #F

Do loXlsx.$write(lstData,lstHeader,'/path/to/file.xlsx')
```

```
## oXlslx.$write

# parameters
#   - pData         List       Data to write
#   - pHeader       List       Column headers (optional)
#   - pPath         String     Full path to the output file
#   - pSheetName    String     Worksheet name (optional - Default: "Sheet1")
# variables
#   - lParam        Row        Parameters passed to the JavaScript Worker

Do lParam.$cols.$add('filename',kCharacter,kSimplechar)
Do lParam.$cols.$add('sheetName',kCharacter,kSimplechar)
Do lParam.$cols.$add('data',kList)
Calculate lParam.filename as pPath
Calculate lParam.sheetName as pSheetName
Calculate lParam.data as pData
If pHeader.$linecount
  Do lParam.$cols.$add('rowHeader',kList)
  Calculate lParam.rowHeader as pHeader
End If

Do $cinst.$callmethod('xlsx','write',lParam,kTrue)
```

### Read

```
Do loXlsx.$init() Returns #F
Do loXlsx.$start() Returns #F

Do loXlsx.$read('/path/to/file.xlsx',lstData)
```

```
## oXlslx.$read

# parameters
#   - pFilename     String            Full path to the input file
#   - pList         Field reference

Do lParam.$cols.$add('filename',kCharacter,kSimplechar)
Calculate lParam.filename as pFilename

Do $cinst.$callmethod('xlsx','read',lParam,kTrue,lErreurText)

If pList.$colcount
  # pList is defined
  Do pList.$merge(iReadList)
Else
  # pList is undefined
  Calculate pList as iReadList
End If


## oXlslx.$methodReturn

# parameters
#   - wReturn       Row        Parameters received from the JavaScript Worker
# local variables
#   - lJson         Binary
# instance variable
#   - iReadList     List

If 'read'=wReturn.__method
  Do OJSON.$listorrowtojson(wReturn.data) Returns lJson
  Do OJSON.$arrayarraytolist(json) Returns iReadList
End If
```

## nmp registry
https://www.npmjs.com/package/@equapro/omnis-xlsx

A GitHub Action automatically publishes the package to npmjs.com.

The workflow is triggered whenever a tag matching the format vX.Y.Z is pushed.
