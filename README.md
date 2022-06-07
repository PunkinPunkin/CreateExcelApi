# CreateExcelApi
Generating excel by request json.

|Method|URL                          |Response Format|
|------|-----------------------------|---------------|
|GET   |swagger/index.html           |html           |
|POST  |api/Excel?fileName={fileName}|attachment     |

## Schemas
### ExcelInfo
|Name         |Type             |Null?|Default     |Comment|
|-------------|-----------------|-----|------------|-------|
|printTime    |Date             |V    |Current Time|       |
|printEmployee|string           |     |            |       |
|sheets       |Array\<[SheetInfo](#sheetinfo)>||||

### SheetInfo
|Name           |Type                 |Null?|Default |Comment     |
|---------------|---------------------|-----|--------|------------|
|name           |string               |     |"Sheet1"|sheet's name|
|title          |string               |V    |        |            |
|subTitle       |string               |V    |        |            |
|searchCondition|string               |V    |        |            |
|tableHeaders   |Array\<[CellInfo](#cellinfo)>||||
|data           |Array<Array\<string>>||||

### CellInfo
|Name           |Type  |Null?|Default|Comment            |
|---------------|------|-----|-------|-------------------|
|name           |string|     |       |table header's name|
|comment        |string|V    |null   |                   |
|fontSize       |number|V    |12     |                   |
|fontColor      |number|V    |8      |`Black`: 8<br>`White`: 9<br>`Red`: 10<br>`Blue`: 12<br>`Green`: 17|
|horizontalAlign|number|V    |0      |`General`: 0<br>`Left`: 1<br>`Center`: 2<br>`Right`: 3<br>`Justify`: 5<br>`Fill`: 4<br>`CenterSelection`: 6<br>`Distributed`: 7|
|verticalAlign  |number|V    |1      |`Top`: 0<br>`Center`: 1<br>`Bottom`: 2<br>`Justify`: 3<br>`Distributed`: 4|

## Demo
* [sample.xlsx](/CreateExcelApi/Sample/sample.xlsx)
* [sample.json](/CreateExcelApi/Sample/sample.json)
```json
{
  "printTime": "2022-06-05T15:51:37.762Z",
  "printEmployee": "admin",
  "sheets": [
    {
      "name": "SalesOrders",
      "title": "Title",
      "subTitle": "Sub Title",
      "searchCondition": "Search by: OrderDate in 2022",
      "tableHeaders": [
        {
          "name": "OrderDate",
          "horizontalAlign": 3
        },
        {
          "name": "Region"
        },
        {
          "name": "Rep"
        },
        {
          "name": "Item"
        },
        {
          "name": "Units",
          "horizontalAlign": 3
        },
        {
          "name": "Units Cost",
          "horizontalAlign": 3
        },
        {
          "name": "Total",
          "comment": "Units * Units Cost",
          "fontColor": 10,
          "horizontalAlign": 3
        }
      ],
      "data": [
        ["1/15/22","Central","Gill","Binder","46","8.99","413.54"],
        ["2/1/22","Central","Smith","Binder","87","15.00","1,305.00"],
        ["2/18/22","East","Jones","Binder","4","4.99","19.96"],
        ["3/7/22","West","Sorvino","Binder","7","19.99","139.93"],
        ["3/24/22","Central","Jardine","Pen Set","50","4.99","249.50"],
        ["4/10/22","Central","Andrews","Pencil","66","1.99","131.34"],
        ["4/27/22","East","Howard","Pen","96","4.99","479.04"],
        ["5/14/22","Central","Gill","Pencil","53","1.29","68.37"],
        ["5/31/22","Central","Gill","Binder","80","8.99","719.20"],
        ["6/17/22","Central","Kivell","Desk","5","125.00","625.00"],
        ["7/4/22","East","Jones","Pen Set","62","4.99","309.38"],
        ["7/21/22","Central","Morgan","Pen Set","55","12.49","686.95"],
        ["8/7/22","Central","Kivell","Pen Set","42","23.95","1,005.90"],
        ["8/24/22","West","Sorvino","Desk","3","275.00","825.00"],
        ["9/10/22","Central","Gill","Pencil","7","1.29","9.03"],
        ["9/27/22","West","Sorvino","Pen","76","1.99","151.24"],
        ["10/14/22","West","Thompson","Binder","57","19.99","1,139.43"],
        ["10/31/22","Central","Andrews","Pencil","14","1.29","18.06"],
        ["11/17/22","Central","Jardine","Binder","11","4.99","54.89"],
        ["12/4/22","Central","Jardine","Binder","94","19.99","1,879.06"],
        ["12/21/22","Central","Andrews","Binder","28","4.99","139.72"]
      ]
    }
  ]
}
```
![sample.jpg](/CreateExcelApi/Sample/sample.jpg "excel result")
