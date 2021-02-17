package com.mz.poi.mapper;

import com.mz.poi.mapper.exception.ExcelConvertException;
import com.mz.poi.mapper.structure.ExcelStructure;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMapper {

  private ExcelMapper() {
    throw new UnsupportedOperationException();
  }

  public static <T> T fromExcel(
      XSSFWorkbook workbook, final Class<T> excelDtoType) {
    return new ExcelReader(workbook).read(excelDtoType);
  }

  public static <T> T fromExcel(
      XSSFWorkbook workbook, final Class<T> excelDtoType, final ExcelStructure excelStructure) {
    return new ExcelReader(workbook).read(excelDtoType, excelStructure);
  }

  public static <T> T fromExcel(
      SXSSFWorkbook workbook, final Class<T> excelDtoType) {
    return new ExcelReader(convert(workbook)).read(excelDtoType);
  }

  public static <T> T fromExcel(
      SXSSFWorkbook workbook, final Class<T> excelDtoType, final ExcelStructure excelStructure) {
    return new ExcelReader(convert(workbook)).read(excelDtoType, excelStructure);
  }

  public static SXSSFWorkbook toExcel(Object excelDto) {
    return new ExcelGenerator(excelDto).generate();
  }

  public static SXSSFWorkbook toExcel(Object excelDto, final ExcelStructure excelStructure) {
    return new ExcelGenerator(excelDto).generate(excelStructure);
  }

  private static XSSFWorkbook convert(SXSSFWorkbook workbook) {
    try {
      ByteArrayOutputStream out = new ByteArrayOutputStream();
      workbook.write(out);
      workbook.close();
      ByteArrayInputStream in = new ByteArrayInputStream(out.toByteArray());
      return new XSSFWorkbook(in);
    } catch (IOException ex) {
      throw new ExcelConvertException("Failed to convert SXSSFWorkbook to XSSFWorkbook ", ex);
    }
  }
}
