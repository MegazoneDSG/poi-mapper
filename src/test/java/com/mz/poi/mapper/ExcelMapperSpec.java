package com.mz.poi.mapper;


import com.mz.poi.mapper.sample.OrderDataRow;
import com.mz.poi.mapper.sample.OrderExcelDto;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


public class ExcelMapperSpec {

  private OrderExcelDto createExcelDto() {
    OrderExcelDto testDto = new OrderExcelDto();
    for (int i = 0; i < 100; i++) {
      testDto.getSheet().getItems().add(
          OrderDataRow
              .builder()
              .productName("productName")
              .productNumber((long) (i + 100))
              .skuId((long) (i + 200))
              .build()
      );
    }
    testDto.getSheet().getInfo2().setOrdererValue("A");
    testDto.getSheet().getInfo2().setSupplierNameValue("B");
    return testDto;
  }

  @Test
  public void model_to_excel() throws IOException {
    OrderExcelDto excelDto = this.createExcelDto();
    XSSFWorkbook excel = ExcelMapper.toExcel(excelDto);
    File file = new File("test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }

  @Test
  public void excel_to_model() {
    OrderExcelDto excelDto = this.createExcelDto();
    XSSFWorkbook excel = ExcelMapper.toExcel(excelDto);
    OrderExcelDto orderExcelDto = ExcelMapper.fromExcel(excel, OrderExcelDto.class);

    assert orderExcelDto.getSheet().getItems().size() == 100;
  }

}
