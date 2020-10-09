package com.mz.poi.mapper;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMapper {

  private ExcelMapper() {
    throw new UnsupportedOperationException();
  }

  public static <T> T fromExcel(
      XSSFWorkbook workbook, final Class<T> excelDtoType) {
    return new ExcelReader(workbook).read(excelDtoType);
  }

  public static XSSFWorkbook toExcel(Object excelDto) {
    return new ExcelGenerator(excelDto).generate();
  }
}
