package com.mz.poi.mapper;

import com.mz.poi.mapper.structure.ExcelStructure;

import org.apache.poi.ss.usermodel.Workbook;

public class ExcelMapper {

  private ExcelMapper() {
    throw new UnsupportedOperationException();
  }

  public static <T> T fromExcel(
      Workbook workbook, final Class<T> excelDtoType) {
    return new ExcelReader(workbook).read(excelDtoType);
  }

  public static <T> T fromExcel(
      Workbook workbook, final Class<T> excelDtoType, final ExcelStructure excelStructure) {
    return new ExcelReader(workbook).read(excelDtoType, excelStructure);
  }

  public static Workbook toExcel(Object excelDto) {
    return new ExcelGenerator(excelDto).generate();
  }

  public static Workbook toExcel(Object excelDto, final ExcelStructure excelStructure) {
    return new ExcelGenerator(excelDto).generate(excelStructure);
  }
}
