package com.mz.poi.mapper;

import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.helper.FormulaHelper;
import com.mz.poi.mapper.structure.CellAnnotation;
import com.mz.poi.mapper.structure.DataRowsAnnotation;
import com.mz.poi.mapper.structure.ExcelStructure;
import com.mz.poi.mapper.structure.ExcelStructure.CellStructure;
import com.mz.poi.mapper.structure.ExcelStructure.RowStructure;
import com.mz.poi.mapper.structure.ExcelStructure.SheetStructure;
import com.mz.poi.mapper.structure.SheetAnnotation;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;

@Getter
public class ExcelReader {

  private ExcelStructure structure;
  private FormulaHelper formulaHelper;
  private XSSFWorkbook workbook;

  public ExcelReader(XSSFWorkbook workbook) {
    this.workbook = workbook;
    this.formulaHelper = new FormulaHelper();
  }

  public <T> T read(final Class<T> excelDtoType) {
    this.structure = new ExcelStructure().build(excelDtoType);
    T excelDto = BeanUtils.instantiateClass(excelDtoType);

    List<SheetStructure> sheets = this.structure.getSheets();
    sheets.stream().sorted(
        Comparator.comparing(sheetStructure -> sheetStructure.getAnnotation().getIndex())
    ).forEach(sheetStructure -> {
      SheetAnnotation annotation = sheetStructure.getAnnotation();
      XSSFSheet sheet = this.workbook.getSheetAt(annotation.getIndex());

      //init new sheet class
      Field sheetField = sheetStructure.getField();
      sheetField.setAccessible(true);
      Object sheetObj = BeanUtils.instantiateClass(sheetField.getType());
      try {
        sheetField.set(excelDto, sheetObj);
      } catch (IllegalAccessException e) {
        throw new IllegalArgumentException(e);
      }

      while (!sheetStructure.isAllRowsGenerated()) {
        RowStructure rowStructure = sheetStructure.nextRowStructure();

        if (!rowStructure.isDataRow()) {
          this.readRow(rowStructure, sheet, sheetObj);
        } else {
          this.readDataRows(rowStructure, sheet, sheetObj);
        }
      }
    });
    return excelDto;
  }

  private void readRow(RowStructure rowStructure, XSSFSheet sheet, Object sheetObj) {
    //init new row class
    Field rowField = rowStructure.getField();
    rowField.setAccessible(true);
    Object rowObj = BeanUtils.instantiateClass(rowField.getType());
    try {
      rowField.set(sheetObj, rowObj);
    } catch (IllegalAccessException e) {
      throw new IllegalArgumentException(e);
    }

    XSSFRow row = sheet.getRow(rowStructure.getStartRowNum());
    if (row == null) {
      rowStructure.setGenerated(true);
      return;
    }
    List<CellStructure> cells = rowStructure.getCells();
    cells.forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();
      XSSFCell cell = row.getCell(cellAnnotation.getColumn());
      this.bindCellValue(cell, cellStructure, rowObj);
    });

    rowStructure.setGenerated(true);
  }

  private void readDataRows(RowStructure rowStructure, XSSFSheet sheet, Object sheetObj) {
    //init new data row class
    List collection = new ArrayList<>();
    try {
      Field sheetField = rowStructure.getSheetField();
      Field collectionField = sheetField.getType().getDeclaredField(rowStructure.getFieldName());
      collectionField.setAccessible(true);
      collectionField.set(sheetObj, collection);
    } catch (IllegalAccessException | NoSuchFieldException e) {
      throw new IllegalArgumentException(e);
    }

    AtomicInteger currentRowNum = new AtomicInteger(rowStructure.getStartRowNum());
    AtomicBoolean readFinished = new AtomicBoolean(false);
    while (!readFinished.get()) {
      boolean isMatch =
          this.readDataRow(rowStructure, currentRowNum.incrementAndGet(), sheet, collection);
      if (!isMatch) {
        readFinished.set(true);
        currentRowNum.decrementAndGet(); // rollback current rowNumber
      }
    }
    rowStructure.setGenerated(true);
    rowStructure.setEndRowNum(currentRowNum.get());
  }

  private boolean readDataRow(
      RowStructure rowStructure, int rowNum, XSSFSheet sheet, List collection) {
    DataRowsAnnotation annotation = (DataRowsAnnotation) rowStructure.getAnnotation();

    XSSFRow row = sheet.getRow(rowNum);
    if (row == null) {
      return false;
    }
    ParameterizedType genericType =
        (ParameterizedType) rowStructure.getField().getGenericType();
    Class<?> dataRowClass = (Class<?>) genericType.getActualTypeArguments()[0];
    Object rowObj = BeanUtils.instantiateClass(dataRowClass);

    //all match 일경우는 하나만 없어도 실패, required 일 경우는 필수값이 없다면 실패
    Match match = annotation.getMatch();
    AtomicBoolean isMatch = new AtomicBoolean(true);

    List<CellStructure> cells = rowStructure.getCells();
    cells.forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();

      XSSFCell cell = row.getCell(cellAnnotation.getColumn());
      boolean isBind = this.bindCellValue(cell, cellStructure, rowObj);
      if (Match.ALL.equals(match)) {
        if (!isBind) {
          isMatch.set(false);
        }
      } else {
        if (!isBind && cellAnnotation.isRequired()) {
          isMatch.set(false);
        }
      }
    });
    if (isMatch.get()) {
      //noinspection unchecked
      collection.add(rowObj);
      return true;
    }
    return false;
  }

  private boolean bindCellValue(Cell cell, CellStructure cellStructure, Object rowObj) {
    try {
      if (!Optional.ofNullable(cell).isPresent()) {
        return false;
      }
      AtomicBoolean isBind = new AtomicBoolean(false);
      cellStructure.getField().setAccessible(true);
      CellType expectedCellType = cellStructure.getAnnotation().getCellType();
      switch (expectedCellType) {
        case STRING:
          if (cell.getCellType().equals(expectedCellType)) {
            cellStructure.getField().set(rowObj, cell.getStringCellValue());
            isBind.set(true);
          }
          break;
        case NUMERIC:
          if (cell.getCellType().equals(expectedCellType)) {
            double cellValue = cell.getNumericCellValue();

            Field cellField = cellStructure.getField();
            cellField.setAccessible(true);
            Class<?> numberType = cellField.getType();
            if (Double.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, cellValue);
            } else if (Float.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, (float) cellValue);
            } else if (Long.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, Double.valueOf(cellValue).longValue());
            } else if (Short.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, Double.valueOf(cellValue).shortValue());
            } else {
              cellField.set(rowObj, Double.valueOf(cellValue).intValue());
            }
            isBind.set(true);
          }
          break;
        case BOOLEAN:
          if (cell.getCellType().equals(expectedCellType)) {
            cellStructure.getField().setBoolean(rowObj, cell.getBooleanCellValue());
            isBind.set(true);
          }
          break;
        default:
          break;
      }
      return isBind.get();
    } catch (IllegalAccessException e) {
      throw new IllegalArgumentException(
          String.format("Invalid type match cell rowNum: %s columnIndex: %s",
              cell.getAddress().getRow(), cell.getAddress().getColumn()));
    }
  }
}
