# Dynamic Array Cell

## Output

![](https://user-images.githubusercontent.com/13447690/110363606-1dde1a00-8086-11eb-9003-f6f32bccbc55.png)

## Insert values to model

```java
public class ArrayCellTemplateSpec {

  private ArrayCellTemplate createModel() {
    ArrayList<ItemRow> itemRows = new ArrayList<>();
    for (int i = 0; i < 5; i++) {
      itemRows.add(
          ItemRow.builder()
              .name("#228839221")
              .description("Product ABC")
              .qty(
                  Stream.of(
                      getRandomNumber(1, 10),
                      getRandomNumber(5, 15),
                      getRandomNumber(0, 30),
                      getRandomNumber(15, 20)
                  ).collect(Collectors.toList())
              )
              .unitPrice(BigDecimal.valueOf(getRandomNumber(1, 200))).build());
    }

    return ArrayCellTemplate
        .builder()
        .sheet(
            ArrayCellSheet
                .builder()
                .itemTable(itemRows)
                .summaryRow(new SummaryRow())
                .build()
        ).build();
  }

  public int getRandomNumber(int min, int max) {
    return (int) ((Math.random() * (max - min)) + min);
  }
}
```

## Model to Excel

```java
public class ExcelMapperSpec {

  @Test
  public void model_to_excel() throws IOException {
    ArrayCellTemplate model = this.createModel();
    Workbook excel = ExcelMapper.toExcel(model);
    File file = new File("array_cell_test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }
}
```

## Excel to Model

```java
public class ExcelMapperSpec {

  @Test
  public void excel_to_model() {
    FileInputStream fis = new FileInputStream("array_cell_test.xlsx");
    XSSFWorkbook excel = new XSSFWorkbook(fis);
    ArrayCellTemplate fromModel = ExcelMapper.fromExcel(excel, ArrayCellTemplate.class);
  }
}
```

## Excel

```java
@Excel(
    defaultStyle = @CellStyle(
        font = @Font(fontName = "Arial")
    ),
    dateFormatZoneId = "Asia/Seoul"
)
public class ArrayCellTemplate {

  @Sheet(
      name = "Order",
      index = 0,
      columnWidths = {
          @ColumnWidth(column = 0, width = 25)
      },
      defaultColumnWidth = 20,
      defaultRowHeightInPoints = 20
  )
  private ArrayCellSheet sheet = new ArrayCellSheet();
}
```

## Sheet

```java
public class ArrayCellSheet {

  @DataRows(
      row = 0,
      match = Match.REQUIRED,
      headers = {
          @Header(name = "ITEM & DESCRIPTION", mappings = {"name", "description"}),
          @Header(name = "UNIT PRICE", mappings = {"unitPrice"}),
          @Header(name = "TOTAL", mappings = {"total"})
      },
      arrayHeaders = {
          @ArrayHeader(simpleNameExpression = "QTY {{index}}", mapping = "qty")
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

## Row

```java
public class ItemRow {

  @Cell(
      column = 0,
      cellType = CellType.STRING,
      required = true
  )
  private String name;

  @Cell(
      columnAfter = "name",
      cols = 2,
      cellType = CellType.STRING,
      required = true
  )
  private String description;

  @ArrayCell(
      size = 4,
      columnAfter = "description",
      cellType = CellType.NUMERIC,
      required = true
  )
  private List<Integer> qty;

  @Cell(
      columnAfter = "qty",
      cellType = CellType.NUMERIC,
      style = @CellStyle(dataFormat = "#,##0.00"),
      required = true
  )
  private BigDecimal unitPrice;

  @Cell(
      columnAfter = "unitPrice",
      cellType = CellType.FORMULA,
      style = @CellStyle(
          dataFormat = "#,##0.00",
          fillForegroundColor = IndexedColors.GREY_25_PERCENT,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      ),
      ignoreParse = true
  )
  private String total = "product(sum({{this.qty[0]}}:{{this.qty[last]}}),{{this.unitPrice}})";
}
```

```java
public class SummaryRow {

  @Cell(
      column = 6,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String title = "SUBTOTAL";

  @Cell(
      columnAfter = "title",
      columnAfterOffset = 1,
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

## Custom Array Header Name

### Output

![](https://user-images.githubusercontent.com/13447690/110363630-22a2ce00-8086-11eb-8928-bfcdc6f872ed.png)

```java
public class ArrayCellTemplate {

  @Test
  public void model_to_excel_with_dynamic_array_header() throws IOException {
    ExcelStructure structure = new ExcelStructure().build(ArrayCellTemplate.class);
    DataRowsAnnotation dataRowsAnnotation =
        (DataRowsAnnotation) structure.getSheet("sheet").getRow("itemTable").getAnnotation();
    dataRowsAnnotation
        .getArrayHeaders()
        .get(0)
        .setArrayHeaderNameExpression(
            new ArrayHeaderNameExpression() {
              private LocalDate start = LocalDate.now();

              @Override
              public String get(int index) {
                return start.plusDays(index).format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
              }
            }
        );
    ArrayCellTemplate model = this.createModel();
    Workbook excel = ExcelMapper.toExcel(model, structure);
    File file = new File("dynamic_array_header_test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }
}
```  
