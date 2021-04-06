# poi-mapper

**Model to Excel, Excel to Model mapper based on apache poi.**

Add annotations to your model you already have. And convert to Excel, or import values from Excel

![](https://user-images.githubusercontent.com/61041926/95656676-dee58000-0b4a-11eb-936e-3cc22d5a4432.png)

- [Basic Usage](./README.md)
- [Dynamic Array Cell](./array-cell.md)

# Include

## Gradle setting

```
repositories {
    maven { url 'https://github.com/MegazoneDSG/maven-repo/raw/master/snapshots' }
}
dependencies {
    compile "com.mz:poi-mapper:1.1.1-SNAPSHOT"
}
```

# Basic Usage

The sample model `PurchaseOrderTemplate` used in the photo at the top of the document is included in the package, and its usage is as follows.

## Insert values to model

```java
public class ExcelMapperSpec {

  private PurchaseOrderTemplate createModel() {
    return PurchaseOrderTemplate
        .builder()
        .sheet(
            OrderSheet
                .builder()
                .titleRow(new TitleRow())
                .infoTable(
                    Stream.of(
                        InfoRow.builder()
                            .vendorTitle("Name").vendorValue("S.Park")
                            .toTitle("Name").toValue("John").build(),
                        InfoRow.builder()
                            .vendorTitle("Company Name").vendorValue("Megazone")
                            .toTitle("Company Name").toValue("DSG").build(),
                        InfoRow.builder()
                            .vendorTitle("Address")
                            .vendorValue("MEGAZONE B/D Yeoksam-dong Gangnam-gu")
                            .toTitle("Address").toValue("DSG B/D Yeoksam-dong Gangnam-gu").build(),
                        InfoRow.builder()
                            .vendorTitle("CT, ST ZIP").vendorValue("SEOUL 06235 KOREA")
                            .toTitle("CT, ST ZIP").toValue("SEOUL 12345 KOREA").build(),
                        InfoRow.builder()
                            .vendorTitle("Phone").vendorValue("T.82(0)2 2108 9105")
                            .toTitle("Phone").toValue("T. 82 (0)2 2109 2500").build()
                    ).collect(Collectors.toList())
                )
                .shipTable(
                    Stream.of(
                        ShipRow.builder()
                            .requester("L.J")
                            .via("Purchase Part")
                            .fob("")
                            .terms("")
                            .deliveryDate(LocalDate.now())
                            .build()
                    ).collect(Collectors.toList())
                )
                .itemTable(
                    Stream.of(
                        ItemRow.builder()
                            .name("#228839221").description("Product ABC").qty(1)
                            .unitPrice(BigDecimal.valueOf(150L)).build(),
                        ItemRow.builder()
                            .name("#428832121").description("Product EFG").qty(15)
                            .unitPrice(BigDecimal.valueOf(12L)).build(),
                        ItemRow.builder()
                            .name("#339884344").description("Product XYZ").qty(78)
                            .unitPrice(BigDecimal.valueOf(1.75)).build()
                    ).collect(Collectors.toList())
                )
                .summaryRow(new SummaryRow())
                .build()
        ).build();
  }
}
```

## Model to Excel

```java
public class ExcelMapperSpec {

  @Test
  public void model_to_excel() throws IOException {
    PurchaseOrderTemplate model = this.createModel();
    Workbook excel = ExcelMapper.toExcel(model);
  }
}
```

## Excel to Model

```java
public class ExcelMapperSpec {

  @Test
  public void excel_to_model() {
    PurchaseOrderTemplate model = this.createModel();
    Workbook excel = ExcelMapper.toExcel(model);
    PurchaseOrderTemplate fromModel = ExcelMapper.fromExcel(excel, PurchaseOrderTemplate.class);
  }
}
```

## Model to Streaming Excel (SXSSF)

```java
public class ExcelMapperSpec {

  @Test
    public void model_to_stream_excel() throws IOException {
      PurchaseOrderTemplate model = this.createModel();
      Workbook excel = ExcelMapper.toExcel(model, new SXSSFWorkbook(50));
    }
}
```

# Create your own model

Let's see how to customize your own model based on the PurchaseOrderTemplate example.

## @Excel

Add `@Excel` annotaion to your model.

```java
@Excel(
    defaultStyle = @CellStyle(
        font = @Font(fontName = "Arial")
    ),
    dateFormatZoneId = "Asia/Seoul"
)
public class PurchaseOrderTemplate {

}
```

### Annotation Description

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| defaultStyle | [@CellStyle](#CellStyle) | [@CellStyle](#CellStyle) |Default cell style of excel |
| dateFormatZoneId | String | Empty | Zone ID to be used when converting dates. If empty, is uses system default zone |

## @Sheet

```java
@Excel(
    defaultStyle = @CellStyle(
        font = @Font(fontName = "Arial")
    ),
    dateFormatZoneId = "Asia/Seoul"
)
public class PurchaseOrderTemplate {

  @Sheet(
      name = "Order",
      index = 0,
      columnWidths = {
          @ColumnWidth(column = 0, width = 25)
      },
      defaultColumnWidth = 20,
      defaultRowHeightInPoints = 20
  )
  private OrderSheet sheet;
}
```

### Annotation Description

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| name | String | None(Required) | sheet name |
| index | int | None((Required)) | sheet index |
| protect | boolean | false | sheet protect or not |
| protectKey | String | Empty String | sheet protect key |
| columnWidths | Array of [@ColumnWidth](#ColumnWidth) | Empty | Specific column width |
| defaultRowHeightInPoints | int | 20 | Default row height of sheet |
| defaultColumnWidth | int | 20 | Default column width of sheet |
| defaultStyle | [@CellStyle](#CellStyle) | [@Excel](#Excel).defaultStyle | Default cell style of sheet |

## Rows

The sample of OrderSheet.

```java
public class OrderSheet {

  @Row(
      row = 0,
      defaultStyle = @CellStyle(
          font = @Font(fontHeightInPoints = 20)
      ),
      heightInPoints = 40
  )
  TitleRow titleRow;

  @DataRows(
      row = 2,
      match = Match.REQUIRED,
      headers = {
          @Header(name = "VENDOR", mappings = {"vendorTitle", "vendorValue"}),
          @Header(name = "SHIP TO", mappings = {"toTitle", "toValue"})
      },
      headerStyle = @CellStyle(
          font = @Font(color = IndexedColors.WHITE),
          fillForegroundColor = IndexedColors.DARK_BLUE,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      )
  )
  List<InfoRow> infoTable;

  @DataRows(
      rowAfter = "infoTable",
      rowAfterOffset = 1,
      match = Match.REQUIRED,
      headers = {
          @Header(name = "REQUESTER", mappings = {"requester"}),
          @Header(name = "SHIP VIA", mappings = {"via"}),
          @Header(name = "F.O.B", mappings = {"fob"}),
          @Header(name = "SHIPPING TERMS", mappings = {"terms"}),
          @Header(name = "DELIVERY DATE", mappings = {"deliveryDate"})
      },
      headerStyle = @CellStyle(
          font = @Font(color = IndexedColors.WHITE),
          fillForegroundColor = IndexedColors.DARK_BLUE,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      ),
      dataStyle = @CellStyle(
          borderTop = BorderStyle.THIN,
          borderBottom = BorderStyle.THIN,
          borderLeft = BorderStyle.THIN,
          borderRight = BorderStyle.THIN
      )
  )
  List<ShipRow> shipTable;

  @DataRows(
      rowAfter = "shipTable",
      rowAfterOffset = 1,
      match = Match.REQUIRED,
      headers = {
          @Header(name = "ITEM", mappings = {"name"}),
          @Header(name = "DESCRIPTION", mappings = {"description"}),
          @Header(name = "QTY", mappings = {"qty"}),
          @Header(name = "UNIT PRICE", mappings = {"unitPrice"}),
          @Header(name = "TOTAL", mappings = {"total"})
      },
      headerStyle = @CellStyle(
          font = @Font(color = IndexedColors.WHITE),
          fillForegroundColor = IndexedColors.DARK_BLUE,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      ),
      dataStyle = @CellStyle(
          borderTop = BorderStyle.THIN,
          borderBottom = BorderStyle.THIN,
          borderLeft = BorderStyle.THIN,
          borderRight = BorderStyle.THIN
      )
  )
  List<ItemRow> itemTable;

  @Row(rowAfter = "itemTable")
  SummaryRow summaryRow;
}
```

## @Row

### Annotation Description

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| row | int | 0 | Index of row |
| rowAfter | String | Empty String | If it is not empty, it is placed after a specific row. You can specify the row field name of the sheet model. |
| rowAfterOffset | int | 0 | If rowAfter is specified, it is offset from the specified row. |
| heightInPoints | int | [@Sheet](#@Sheet).defaultRowHeightInPoints | Specified row height. |
| defaultStyle | [@CellStyle](#CellStyle) | [@Sheet](#Sheet).defaultStyle | Default cell style of row |

## @DataRows

@Datarows should be annotated in List.class. (Other collection class not supported yet.)

### Annotation Description

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| row | int | 0 | Index of row |
| rowAfter | String | Empty String | If it is not empty, it is placed after a end of specific row. You can specify the row field name of the sheet model. |
| rowAfterOffset | int | 0 | If rowAfter is specified, it is offset from the specified row. |
| headerHeightInPoints | int | [@Sheet](#@Sheet).defaultRowHeightInPoints | Specified header row height. |
| headerStyle | [@CellStyle](#CellStyle) | [@Sheet](#Sheet).defaultStyle | Default cell style of header row |
| headers | Array of [@Header](#Header) | Empty | Array of Headers |
| arrayHeaders | Array of [@ArrayHeader](#ArrayHeader) | Empty | Array of ArrayHeaders |
| hideHeader | boolean | false | Whether hide headers or not |
| dataHeightInPoints | int | [@Sheet](#@Sheet).defaultRowHeightInPoints | Specified data row height. |
| dataStyle | [@CellStyle](#CellStyle) | [@Sheet](#Sheet).defaultStyle | Default cell style of data row |
| match | [Match](#Match) | Match.ALL | DataRow recognition condition when converting Excel to model |

### Match

Row recognition condition when converting Excel to model.

| option  | description |
|------------|--------|
| ALL | All column values must exist to be recognized as DataRow.  |
| REQUIRED | If only the value of the column annotated with [@Cell](#Cell).required exists, it is recognized as DataRow.  |
| STOP_ON_BLANK | It is recognized as a data row until it encounters a blank row.  |

## @Cell

### Annotation Description

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| column | int | 0 | Index of column |
| cols | int | 1 | The number of columns to be merged |
| columnAfter | String | Empty String | If it is not empty, it is placed after a end of specific cell. You can specify the cell field name of the row model. |
| columnAfterOffset | int | 0 | If columnAfter is specified, it is offset from the specified cell. |
| cellType | [CellType](#CellType) | None(Required) | CellType of cell |
| ignoreParse | boolean | false | When converting Excel to Model, do not bind values. |
| required | boolean | false | Required value when converting Excel to Model. See [Match](#Match) |
| headers | Array of [@Header](#Header) | Empty | Array of Headers |
| style | [@CellStyle](#CellStyle) | [@Row](#Row).defaultStyle, [@DataRows](#DataRows).defaultStyle | cell style |

## @ArrayCell

### Annotation Description

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| column | int | 0 | Index of column |
| cols | int | 1 | The number of columns to be merged |
| columnAfter | String | Empty String | If it is not empty, it is placed after a end of specific cell. You can specify the cell field name of the row model. |
| columnAfterOffset | int | 0 | If columnAfter is specified, it is offset from the specified cell. |
| cellType | [CellType](#CellType) | None(Required) | CellType of cell |
| ignoreParse | boolean | false | When converting Excel to Model, do not bind values. |
| required | boolean | false | Required value when converting Excel to Model. See [Match](#Match) |
| headers | Array of [@Header](#Header) | Empty | Array of Headers |
| style | [@CellStyle](#CellStyle) | [@Row](#Row).defaultStyle, [@DataRows](#DataRows).defaultStyle | cell style |
| size | int | 0 | The size of the array cells. When reading or writing cells, only the size is applied. You can resize dynamically at runtime by referring to the following documentation:The size of the array cells. When reading or writing cells, only the size is applied. [You can resize dynamically at runtime by referring to the following documentation](./array-cell.md) |

### CellType

This is the CellType enums. If the model class does not match CellType, it will not be converted.

| option  | Matching Java Class | description |
|------------|--------|--------|
| NONE | Any | When creating Excel, do not assign a specific cell type to the cell. |
| STRING | String | -  |
| NUMERIC | Double,Float,Long,Short,BigDecimal,BigInteger,Integer | - |
| BLANK | None | When creating an Excel, it becomes an empty cell, and the value is not converted when reading the Excel. |
| BOOLEAN | Boolean | - |
| DATE | LocalDate,LocalDateTime | - |
| FORMULA | String | You can use the [FormulaAddressExpression](#FormulaAddressExpression) to write native formulas in Excel with specific cell locations in the model. |

### FormulaAddressExpression

It is an expression that converts a specific cell location in the model to an Excel address such as (A1,B2....)

*Only addresses within the same sheet can be converted.*

| expression  | example | description |
|------------|--------|--------|
| {{rowFiledName.cellFieldName}} | titleRow.title  | If titleRow's row is 0 and title's column is 0, return **A1** |
| {{rowFiledName[last].cellFieldName}} | itemTable[last].total  | Works only in DataRows. If itemTable's end row is 15 and total's column is 5, return **F16**. See [SummaryRow](#SummaryRow) sample |
| {{rowFiledName[Number].cellFieldName}} | itemTable[0].total  | Works only in DataRows. If itemTable's start row is 13 and total's column is 5, return **F14**. See [SummaryRow](#SummaryRow) sample |
| {{rowFiledName.cellFieldName[last]}} | itemTable.qty[last]  | Works only in ArrayCell. See [Dynamic Array Cell](./array-cell.md) sample |
| {{rowFiledName.cellFieldName[Number]}} | itemTable.qty[0]  | Works only in ArrayCell. See [Dynamic Array Cell](./array-cell.md) sample |
| {{this.cellFieldName}} | this.qty  | Works only in DataRows. expression `this` means cell's current row. If the qty's current row is 13 and qty's coumn is 3, return **D14**. See [ItemRow](#ItemRow) sample |


### Row model samples includes cells.

The samples of row model includes cells.

![](https://user-images.githubusercontent.com/61041926/95660437-ddc04d00-0b62-11eb-9e02-be6b6a839b88.png)

### TitleRow

```java
public class TitleRow {

  @Cell(
      column = 0,
      cols = 6,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String title = "PURCHASE ORDER";
}
```

### InfoRow

```java
public class InfoRow {

  @Cell(
      column = 0,
      cellType = CellType.STRING
  )
  private String vendorTitle;

  @Cell(
      column = 1,
      cols = 2,
      cellType = CellType.STRING,
      required = true
  )
  private String vendorValue;

  @Cell(
      column = 3,
      cellType = CellType.STRING
  )
  private String toTitle;

  @Cell(
      column = 4,
      cols = 2,
      cellType = CellType.STRING,
      required = true
  )
  private String toValue;
}
```

### ShipRow

```java
public class ShipRow {

  @Cell(
      column = 0,
      cellType = CellType.STRING
  )
  private String requester;

  @Cell(
      column = 1,
      cellType = CellType.STRING
  )
  private String via;

  @Cell(
      column = 2,
      cellType = CellType.STRING
  )
  private String fob;

  @Cell(
      column = 3,
      cols = 2,
      cellType = CellType.STRING
  )
  private String terms;

  @Cell(
      column = 5,
      cellType = CellType.DATE,
      style = @CellStyle(dataFormat = "yyyy-MM-dd"),
      required = true
  )
  private LocalDate deliveryDate;
}
```

### ItemRow

```java
public class ItemRow {

  @Cell(
      column = 0,
      cellType = CellType.STRING,
      required = true
  )
  private String name;

  @Cell(
      column = 1,
      cols = 2,
      cellType = CellType.STRING,
      required = true
  )
  private String description;

  @Cell(
      column = 3,
      cellType = CellType.NUMERIC,
      required = true
  )
  private long qty;

  @Cell(
      column = 4,
      cellType = CellType.NUMERIC,
      style = @CellStyle(dataFormat = "#,##0.00"),
      required = true
  )
  private BigDecimal unitPrice;

  @Cell(
      column = 5,
      cellType = CellType.FORMULA,
      style = @CellStyle(
          dataFormat = "#,##0.00",
          fillForegroundColor = IndexedColors.GREY_25_PERCENT,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      ),
      ignoreParse = true
  )
  private String total = "product({{this.qty}},{{this.unitPrice}})";
}
```

### SummaryRow

```java
public class SummaryRow {

  @Cell(
      column = 4,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String title = "SUBTOTAL";

  @Cell(
      column = 5,
      cellType = CellType.FORMULA,
      style = @CellStyle(
          fillForegroundColor = IndexedColors.AQUA,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      ),
      ignoreParse = true
  )
  private String formula = "SUM({{itemTable[0].total}}:{{itemTable[last].total}})";
}
```

## @CellStyle

### Annotation Description

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| font | [@Font](#Font) | [@Font](#Font) | font of cell |
| dataFormat | String | General | DataFormat of cell. See [BuiltinFormats](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BuiltinFormats.html). 
In addition to BuiltinFormats, you can also use generic dateFormats (such as yyyy.MM.dd). |
| hidden | boolean | false | whether the cell's using this style are to be hidden |
| locked | boolean | false | whether the cell's using this style are to be locked |
| quotePrefixed | boolean | false | Is "Quote Prefix" or "123 Prefix" enabled for the cell |
| alignment | [HorizontalAlignment](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/HorizontalAlignment.html) | GENERAL | the type of horizontal alignment for the cell |
| wrapText | boolean | false | whether the text should be wrapped |
| verticalAlignment | [VerticalAlignment](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/VerticalAlignment.html) | BOTTOM | the type of vertical alignment for the cell |
| rotation | short | 0 | the degree of rotation for the text in the cell. |
| indention | short | 0 | the number of spaces to indent the text in the cell. |
| borderLeft | [BorderStyle](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BorderStyle.html) | NONE | the type of border to use for the left border of the cell |
| borderRight | [BorderStyle](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BorderStyle.html) | NONE | the type of border to use for the right border of the cell |
| borderTop | [BorderStyle](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BorderStyle.html) | NONE | the type of border to use for the top border of the cell |
| borderBottom | [BorderStyle](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BorderStyle.html) | NONE | the type of border to use for the bottom border of the cell |
| leftBorderColor | [IndexedColors](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html) | AUTOMATIC | the color to use for the left border |
| rightBorderColor | [IndexedColors](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html) | AUTOMATIC | the color to use for the right border |
| topBorderColor | [IndexedColors](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html) | AUTOMATIC | the color to use for the top border |
| bottomBorderColor | [IndexedColors](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html) | AUTOMATIC | the color to use for the bottom border |
| fillPattern | [FillPatternType](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/FillPatternType.html) | NO_FILL | the fill pattern |
| fillBackgroundColor | [IndexedColors](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html) | AUTOMATIC | the background fill color |
| fillForegroundColor | [IndexedColors](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html) | AUTOMATIC | the foreground fill color |
| shrinkToFit | boolean | false | Should the Cell be auto-sized by Excel to shrink it to fit if this text is too long |

## @Font

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| fontName | String | Arial | the name for the font |
| fontHeightInPoints | short | 10 | the font height |
| italic | boolean | false | whether to use italics or not |
| strikeout | boolean | false | whether to use a strikeout horizontal line through the text or not |
| color | [IndexedColors](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html) | AUTOMATIC | the color for the font |
| typeOffset | short | 0 | 0 = NONE, 1 = SUPER, 2 = SUB |
| underline | short | 0 | type of text underlining to use. 0 = NONE, 1 = SINGLE, 2 = DOUBLE, SINGLE_ACCOUNTING = 0x21, DOUBLE_ACCOUNTING = 0x22 |
| charSet | int | 0 | 0 = ANSI_CHARSET, 1 = DEFAULT_CHARSET, 2 = SYMBOL_CHARSET |
| bold | boolean | false | whether to use bold or not |

## @Header

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| mappings | Array of String | Empty | These are the cell field names of the datarow to be mapped. |
| style | [@CellStyle](#CellStyle) | [@DataRows](#DataRows).defaultStyle | cell style |

## @ArrayHeader

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| mapping | Array of String | Empty | The cell field name of the datarow to be mapped. |
| style | [@CellStyle](#CellStyle) | [@DataRows](#DataRows).defaultStyle | cell style |
| simpleNameExpression | String | {{index}} | Simply indicate the header name of the array cell using the provided index string. [YAt runtime, you can use a separate expression class to express more specific values](./array-cell.md) |

## @ColumnWidth

| attribute  | type | default | description |
|------------|--------|----------|-----------------------|
| column | int | None(Required) | column index of sheet |
| width | int | None(Required) | column width |

## License

Apache License Version 2.0, January 2004
[http://www.apache.org/licenses/](http://www.apache.org/licenses/)

## Contact && Issue

If you find a bug or want to improve the function, please create a github issue.

Any other questions: seungpilpark@mz.co.kr

Author **Seungpil Park, Megazone Inc.**
