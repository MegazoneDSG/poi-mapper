package com.mz.poi.mapper;


import com.mz.poi.mapper.example.arraycell.ArrayCellSheet;
import com.mz.poi.mapper.example.arraycell.ArrayCellTemplate;
import com.mz.poi.mapper.example.arraycell.ItemRow;
import com.mz.poi.mapper.example.arraycell.SummaryRow;
import com.mz.poi.mapper.expression.ArrayHeaderNameExpression;
import com.mz.poi.mapper.structure.ArrayCellAnnotation;
import com.mz.poi.mapper.structure.DataRowsAnnotation;
import com.mz.poi.mapper.structure.ExcelStructure;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ArrayCellTemplateSpec extends MapperTestSupport {

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

  @Test
  public void model_to_excel() throws IOException {
    ArrayCellTemplate model = this.createModel();
    Workbook excel = ExcelMapper.toExcel(model);
    File file = new File(testDir + "/array_cell_test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }

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
    File file = new File(testDir + "/dynamic_array_header_test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }

  @Test
  public void model_to_excel_with_dynamic_array_size() throws IOException {
    ExcelStructure structure = new ExcelStructure().build(ArrayCellTemplate.class);
    ArrayCellAnnotation arrayCellAnnotation =
        (ArrayCellAnnotation) structure
            .getSheet("sheet")
            .getRow("itemTable")
            .getCell("qty")
            .getAnnotation();
    arrayCellAnnotation.setSize(8);

    ArrayCellTemplate model = this.createModel();
    Workbook excel = ExcelMapper.toExcel(model, structure);
    File file = new File(testDir + "/dynamic_array_size_test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }

  @Test
  public void excel_to_model() throws IOException {
    FileInputStream fis = new FileInputStream(testDir + "/array_cell_test.xlsx");
    XSSFWorkbook excel = new XSSFWorkbook(fis);
    ArrayCellTemplate fromModel = ExcelMapper.fromExcel(excel, ArrayCellTemplate.class);
    assert fromModel.getSheet().getItemTable().size() == 5;
    fromModel.getSheet().getItemTable().forEach(itemRow -> {
      System.out.println(Arrays.toString(itemRow.getQty().toArray(new Integer[0])));
      assert itemRow.getQty().size() == 4;
    });
  }

}
