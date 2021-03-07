package com.mz.poi.mapper;

import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.exception.ExcelReadException;
import com.mz.poi.mapper.exception.ReadExceptionAddress;
import com.mz.poi.mapper.helper.DateFormatHelper;
import com.mz.poi.mapper.helper.FormulaHelper;
import com.mz.poi.mapper.helper.InheritedFieldHelper;
import com.mz.poi.mapper.structure.CellAnnotation;
import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.CellType;
import com.mz.poi.mapper.structure.DataRowsAnnotation;
import com.mz.poi.mapper.structure.ExcelStructure;
import com.mz.poi.mapper.structure.RowStructure;
import com.mz.poi.mapper.structure.SheetAnnotation;
import com.mz.poi.mapper.structure.SheetStructure;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.BeanUtils;

@Getter
public class ExcelReader {

  private ExcelStructure structure;
  private FormulaHelper formulaHelper;
  private Workbook workbook;

  public ExcelReader(Workbook workbook) {
    this.workbook = workbook;
    this.formulaHelper = new FormulaHelper();
  }

  public <T> T read(final Class<T> excelDtoType, ExcelStructure excelStructure) {
    this.structure = excelStructure;
    return this.read(excelDtoType);
  }

  public <T> T read(final Class<T> excelDtoType) {
    if (this.structure == null) {
      this.structure = new ExcelStructure().build(excelDtoType);
    }
    this.structure.prepareReadStructure();
    T excelDto = BeanUtils.instantiateClass(excelDtoType);

    List<SheetStructure> sheets = this.structure.getSheets();
    sheets.stream().sorted(
        Comparator.comparing(sheetStructure -> sheetStructure.getAnnotation().getIndex())
    ).forEach(sheetStructure -> {
      SheetAnnotation annotation = sheetStructure.getAnnotation();
      Sheet sheet = this.workbook.getSheetAt(annotation.getIndex());

      //init new sheet class
      Field sheetField = sheetStructure.getField();
      sheetField.setAccessible(true);
      Object sheetObj = BeanUtils.instantiateClass(sheetField.getType());
      try {
        sheetField.set(excelDto, sheetObj);
      } catch (IllegalAccessException e) {
        throw new ExcelReadException("Invalid sheet class", e,
            new ReadExceptionAddress(annotation.getIndex()));
      }

      while (!sheetStructure.isAllRowsRead()) {
        RowStructure rowStructure = sheetStructure.nextReadRowStructure();
        if (!rowStructure.isDataRow()) {
          this.readRow(rowStructure, sheet, sheetObj);
        } else {
          this.readDataRows(rowStructure, sheet, sheetObj);
        }
      }
    });
    return excelDto;
  }

  private void readRow(RowStructure rowStructure, Sheet sheet, Object sheetObj) {
    //init new row class
    Field rowField = rowStructure.getField();
    rowField.setAccessible(true);
    Object rowObj = BeanUtils.instantiateClass(rowField.getType());
    try {
      rowField.set(sheetObj, rowObj);
    } catch (IllegalAccessException e) {
      throw new ExcelReadException("Invalid row class", e,
          new ReadExceptionAddress(
              this.workbook.getSheetIndex(sheet.getSheetName()), rowStructure.getStartRowNum())
      );
    }

    Row row = sheet.getRow(rowStructure.getStartRowNum());
    if (row == null) {
      rowStructure.setRead(true);
      return;
    }
    List<CellStructure> cells = rowStructure.getCells();
    cells.forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();
      Cell cell = row.getCell(cellAnnotation.getColumn());
      this.bindCellValue(cell, cellStructure, rowObj);
    });

    rowStructure.setRead(true);
  }

  private void readDataRows(RowStructure rowStructure, Sheet sheet, Object sheetObj) {
    //init new data row class
    List collection = new ArrayList<>();
    try {
      Field sheetField = rowStructure.getSheetField();
      Field collectionField = InheritedFieldHelper
          .getDeclaredField(sheetField.getType(), rowStructure.getFieldName());
      collectionField.setAccessible(true);
      collectionField.set(sheetObj, collection);
    } catch (IllegalAccessException | NoSuchFieldException e) {
      throw new ExcelReadException("Invalid data row class", e,
          new ReadExceptionAddress(
              this.workbook.getSheetIndex(sheet.getSheetName()), rowStructure.getStartRowNum())
      );
    }

    AtomicInteger currentRowNum = new AtomicInteger(rowStructure.getStartRowNum());
    if (rowStructure.isDataRowAndHideHeader()) {
      currentRowNum.decrementAndGet();
    }
    AtomicBoolean readFinished = new AtomicBoolean(false);
    while (!readFinished.get()) {
      boolean isMatch =
          this.readDataRow(rowStructure, currentRowNum.incrementAndGet(), sheet, collection);
      if (!isMatch) {
        readFinished.set(true);
        currentRowNum.decrementAndGet(); // rollback current rowNumber
      }
    }
    rowStructure.setRead(true);
    rowStructure.setEndRowNum(currentRowNum.get());
  }

  private boolean readDataRow(
      RowStructure rowStructure, int rowNum, Sheet sheet, List collection) {

    Row row = sheet.getRow(rowNum);
    if (row == null) {
      return false;
    }
    ParameterizedType genericType =
        (ParameterizedType) rowStructure.getField().getGenericType();
    Class<?> dataRowClass = (Class<?>) genericType.getActualTypeArguments()[0];
    Object rowObj = BeanUtils.instantiateClass(dataRowClass);

    DataRowsAnnotation annotation = (DataRowsAnnotation) rowStructure.getAnnotation();
    Match match = annotation.getMatch();
    boolean isMatch = false;
    if (Match.REQUIRED.equals(match)) {
      isMatch = this.isRequiredMatch(rowStructure, row, rowObj);
    } else if (Match.ALL.equals(match)) {
      isMatch = this.isAllMatch(rowStructure, row, rowObj);
    } else if (Match.STOP_ON_BLANK.equals(match)) {
      isMatch = this.isNotBlankMatch(rowStructure, row, rowObj);
    }
    if (isMatch) {
      collection.add(rowObj);
      return true;
    }
    return false;
  }

  private boolean isNotBlankMatch(RowStructure rowStructure, Row row, Object rowObj) {
    boolean isMatch = rowStructure.getCells().stream()
        .anyMatch(cellStructure -> {
          int column = cellStructure.getAnnotation().getColumn();
          Cell cell = row.getCell(column);
          return !(cell == null
              || cell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK);
        });
    List<CellStructure> cells = rowStructure.getCells();
    cells.forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();
      Cell cell = row.getCell(cellAnnotation.getColumn());
      this.bindCellValue(cell, cellStructure, rowObj);
    });
    return isMatch;
  }

  private boolean isRequiredMatch(RowStructure rowStructure, Row row, Object rowObj) {
    AtomicBoolean isMatch = new AtomicBoolean(true);
    rowStructure.getCells().forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();
      Cell cell = row.getCell(cellAnnotation.getColumn());
      boolean isBind = this.bindCellValue(cell, cellStructure, rowObj);
      if (cellAnnotation.isRequired() && !isBind) {
        isMatch.set(false);
      }
    });
    return isMatch.get();
  }

  private boolean isAllMatch(RowStructure rowStructure, Row row, Object rowObj) {
    AtomicBoolean isMatch = new AtomicBoolean(true);
    rowStructure.getCells().forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();
      Cell cell = row.getCell(cellAnnotation.getColumn());
      boolean isBind = this.bindCellValue(cell, cellStructure, rowObj);
      if (!isBind) {
        isMatch.set(false);
      }
    });
    return isMatch.get();
  }

  private boolean bindCellValue(Cell cell, CellStructure cellStructure, Object rowObj) {
    try {
      if (!Optional.ofNullable(cell).isPresent() || cellStructure.getAnnotation().isIgnoreParse()) {
        return false;
      }
      AtomicBoolean isBind = new AtomicBoolean(false);
      cellStructure.getField().setAccessible(true);
      CellType expectedCellType = cellStructure.getAnnotation().getCellType();
      org.apache.poi.ss.usermodel.CellType expectedExcelCellType = expectedCellType
          .toExcelCellType();
      switch (expectedCellType) {
        case DATE:
          boolean isCellDateFormatted;
          try {
            isCellDateFormatted = DateUtil.isCellDateFormatted(cell);
          } catch (IllegalStateException ex) {
            break;
          }
          if (isCellDateFormatted) {
            Date cellValue = cell.getDateCellValue();

            Field cellField = cellStructure.getField();
            cellField.setAccessible(true);
            Class<?> dateType = cellField.getType();
            if (LocalDateTime.class.isAssignableFrom(dateType)) {
              cellField.set(rowObj, DateFormatHelper.getLocalDateTime(cellValue,
                  this.structure.getAnnotation().getDateFormatZoneId()));
            } else if (LocalDate.class.isAssignableFrom(dateType)) {
              cellField.set(rowObj, DateFormatHelper.getLocalDate(cellValue,
                  this.structure.getAnnotation().getDateFormatZoneId()));
            } else {
              throw new ExcelReadException(
                  String.format("not supported date type %s", dateType.getName()),
                  new ReadExceptionAddress(
                      this.workbook.getSheetIndex(cell.getSheet().getSheetName()),
                      cell.getRowIndex(),
                      cell.getColumnIndex()
                  )
              );
            }
            isBind.set(true);
          }
          break;
        case STRING:
          if (cell.getCellType().equals(expectedExcelCellType)) {
            cellStructure.getField().set(rowObj, cell.getStringCellValue());
            isBind.set(true);
          }
          break;
        case NUMERIC:
          if (cell.getCellType().equals(expectedExcelCellType)) {
            double cellValue = cell.getNumericCellValue();

            Field cellField = cellStructure.getField();
            cellField.setAccessible(true);
            Class<?> numberType = cellField.getType();
            if (Double.class.isAssignableFrom(numberType) ||
                double.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, cellValue);
            } else if (Float.class.isAssignableFrom(numberType) ||
                float.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, (float) cellValue);
            } else if (Long.class.isAssignableFrom(numberType) ||
                long.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, Double.valueOf(cellValue).longValue());
            } else if (Short.class.isAssignableFrom(numberType) ||
                short.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, Double.valueOf(cellValue).shortValue());
            } else if (BigDecimal.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, BigDecimal.valueOf(cellValue));
            } else if (BigInteger.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, BigInteger.valueOf(Double.valueOf(cellValue).longValue()));
            } else if (Integer.class.isAssignableFrom(numberType) ||
                int.class.isAssignableFrom(numberType)) {
              cellField.set(rowObj, Double.valueOf(cellValue).intValue());
            } else {
              throw new ExcelReadException(
                  String.format("not supported number type %s", numberType.getName()),
                  new ReadExceptionAddress(
                      this.workbook.getSheetIndex(cell.getSheet().getSheetName()),
                      cell.getRowIndex(),
                      cell.getColumnIndex()
                  )
              );
            }
            isBind.set(true);
          }
          break;
        case BOOLEAN:
          if (cell.getCellType().equals(expectedExcelCellType)) {
            cellStructure.getField().setBoolean(rowObj, cell.getBooleanCellValue());
            isBind.set(true);
          }
          break;
        default:
          break;
      }
      return isBind.get();
    } catch (IllegalAccessException e) {
      throw new ExcelReadException("can not set cell value", e,
          new ReadExceptionAddress(
              this.workbook.getSheetIndex(cell.getSheet().getSheetName()),
              cell.getRowIndex(),
              cell.getColumnIndex()
          )
      );
    }
  }
}
