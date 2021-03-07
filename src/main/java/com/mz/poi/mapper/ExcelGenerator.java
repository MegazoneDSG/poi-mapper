package com.mz.poi.mapper;

import com.mz.poi.mapper.exception.ExcelGenerateException;
import com.mz.poi.mapper.helper.DateFormatHelper;
import com.mz.poi.mapper.helper.FormulaHelper;
import com.mz.poi.mapper.structure.CellAnnotation;
import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.CellStyleAnnotation;
import com.mz.poi.mapper.structure.CellType;
import com.mz.poi.mapper.structure.DataRowsAnnotation;
import com.mz.poi.mapper.structure.ExcelStructure;
import com.mz.poi.mapper.structure.FontAnnotation;
import com.mz.poi.mapper.structure.RowAnnotation;
import com.mz.poi.mapper.structure.RowStructure;
import com.mz.poi.mapper.structure.SheetAnnotation;
import com.mz.poi.mapper.structure.SheetStructure;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Collection;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

@Getter
public class ExcelGenerator {

  private Workbook workbook;
  private ExcelStructure structure;
  private Object excelDto;
  private FormulaHelper formulaHelper;

  public ExcelGenerator(Object excelDto, Workbook workbook) {
    this.workbook = workbook;
    this.excelDto = excelDto;
    this.formulaHelper = new FormulaHelper();
  }

  public Workbook generate(final ExcelStructure excelStructure) {
    this.structure = excelStructure;
    return this.generate();
  }

  public Workbook generate() {
    if (this.structure == null) {
      this.structure = new ExcelStructure().build(excelDto.getClass());
    }
    this.structure.prepareGenerateStructure(excelDto);

    List<SheetStructure> sheets = this.structure.getSheets();
    sheets.stream().sorted(
        Comparator.comparing(sheetStructure -> sheetStructure.getAnnotation().getIndex())
    ).forEach(sheetStructure -> {
      SheetAnnotation annotation = sheetStructure.getAnnotation();
      Sheet sheet = this.workbook.createSheet(annotation.getName());
      if (annotation.isProtect()) {
        sheet.protectSheet(annotation.getProtectKey());
      }
      sheet.setDefaultRowHeightInPoints(annotation.getDefaultRowHeightInPoints());
      sheet.setDefaultColumnWidth(annotation.getDefaultColumnWidth());
      annotation.getColumnWidths()
          .forEach(columnWidth -> {
            sheet.setColumnWidth(
                columnWidth.getColumn(),
                columnWidth.getWidth() * 256
            );
          });

      sheetStructure.getRows()
          .stream()
          .sorted(Comparator.comparing(RowStructure::getStartRowNum))
          .forEach(rowStructure -> {
            if (!rowStructure.isDataRow()) {
              this.drawRow(rowStructure, sheet);
            } else {
              this.drawDataRows(rowStructure, sheet);
            }
          });
      // this.formulaHelper.applySheetFormulas(sheetStructure);
    });
    return this.workbook;
  }

  private void drawRow(RowStructure rowStructure, Sheet sheet) {
    Row row = sheet.createRow(rowStructure.getStartRowNum());

    RowAnnotation rowAnnotation = (RowAnnotation) rowStructure.getAnnotation();
    if (rowAnnotation.isUseRowHeightInPoints()) {
      row.setHeightInPoints(rowAnnotation.getHeightInPoints());
    }
    List<CellStructure> cells = rowStructure.getCells();
    cells.forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();

      //스타일 적용
      CellStyle cellStyle = this.createCellStyle(cellAnnotation.getStyle());
      Cell cell = row.createCell(
          cellAnnotation.getColumn(), cellAnnotation.getCellType().toExcelCellType()
      );
      cell.setCellStyle(cellStyle);
      //cols 적용
      this.mergeCell(cell, cellAnnotation.getColumn(), cellAnnotation.getCols());
      //값 바인딩
      Object cellValue = this.findCellValue(cellStructure);
      this.bindCellValue(cell, cellAnnotation.getCellType(), cellValue, cellStructure,
          rowStructure.getStartRowNum());
    });
  }

  private void drawDataRows(RowStructure rowStructure, Sheet sheet) {
    AtomicInteger currentRowNum = new AtomicInteger(rowStructure.getStartRowNum());
    DataRowsAnnotation annotation = (DataRowsAnnotation) rowStructure.getAnnotation();

    // draw header
    if (rowStructure.isDataRowAndHideHeader()) {
      currentRowNum.decrementAndGet();
    } else {
      this.drawDataHeaderRow(annotation, currentRowNum.get(), sheet);
    }

    // draw cachedDataRowStyle
    Map<String, CellStyle> cachedDataRowStyle = this.createCachedDataRowStyle(rowStructure);
    Collection<?> items = rowStructure.findRowDataCollection(this.excelDto);
    if (items == null) {
      return;
    }
    items.forEach(item -> {
      this.drawDataRow(rowStructure, currentRowNum.incrementAndGet(), item, cachedDataRowStyle,
          sheet);
    });
  }

  private Map<String, CellStyle> createCachedDataRowStyle(RowStructure rowStructure) {
    List<CellStructure> cells = rowStructure.getCells();
    Map<String, CellStyle> cachedDataRowStyle = new HashMap<>();
    cells.forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();
      cachedDataRowStyle.put(
          cellStructure.getFieldName(),
          this.createCellStyle(cellAnnotation.getStyle()));
    });
    return cachedDataRowStyle;
  }

  private Optional<CellStyle> getCachedDataRowStyle(
      Map<String, CellStyle> cachedDataRowStyle, String fieldName) {
    if (!cachedDataRowStyle.containsKey(fieldName)) {
      return Optional.empty();
    }
    return Optional.of(cachedDataRowStyle.get(fieldName));
  }

  private void drawDataRow(
      RowStructure rowStructure, int rowNum, Object item, Map<String, CellStyle> cachedDataRowStyle,
      Sheet sheet) {

    DataRowsAnnotation annotation = (DataRowsAnnotation) rowStructure.getAnnotation();

    Row row = sheet.createRow(rowNum);
    if (annotation.isUseDataHeightInPoints()) {
      row.setHeightInPoints(annotation.getDataHeightInPoints());
    }

    List<CellStructure> cells = rowStructure.getCells();
    cells.forEach(cellStructure -> {
      CellAnnotation cellAnnotation = cellStructure.getAnnotation();
      Cell cell = row.createCell(
          cellAnnotation.getColumn(), cellAnnotation.getCellType().toExcelCellType()
      );
      this.getCachedDataRowStyle(cachedDataRowStyle, cellStructure.getFieldName())
          .ifPresent(cell::setCellStyle);
      //cols 적용
      this.mergeCell(cell, cellAnnotation.getColumn(), cellAnnotation.getCols());
      //값 바인딩
      Object cellValue;
      try {
        cellValue = cellStructure.getField().get(item);
      } catch (IllegalAccessException e) {
        throw new ExcelGenerateException(
            String.format("can not find field from data item, %s", cellStructure.getFieldName()),
            e);
      }
      this.bindCellValue(cell, cellAnnotation.getCellType(), cellValue, cellStructure, rowNum);
    });
  }

  private void drawDataHeaderRow(DataRowsAnnotation annotation, int rowNum, Sheet sheet) {
    Row row = sheet.createRow(rowNum);
    if (annotation.isUseHeaderHeightInPoints()) {
      row.setHeightInPoints(annotation.getHeaderHeightInPoints());
    }
    annotation.getHeaders().forEach(headerAnnotation -> {
      CellStyle cellStyle = this.createCellStyle(headerAnnotation.getStyle());
      Cell cell = row.createCell(
          headerAnnotation.getColumn(), org.apache.poi.ss.usermodel.CellType.STRING
      );
      cell.setCellStyle(cellStyle);
      this.mergeCell(cell, headerAnnotation.getColumn(), headerAnnotation.getCols());
      this.bindCellValue(cell, CellType.STRING, headerAnnotation.getName());
    });
  }

  private void mergeCell(Cell cell, int columnIndex, int cols) {
    if (cols < 2) {
      return;
    }
    CellRangeAddress cellRangeAddress = new CellRangeAddress(
        cell.getRow().getRowNum(), cell.getRow().getRowNum(),
        columnIndex, (columnIndex + cols - 1));
    //merge cell
    cell.getSheet().addMergedRegion(cellRangeAddress);
    //apply merged border
    RegionUtil.setBorderTop(
        cell.getCellStyle().getBorderTop(), cellRangeAddress, cell.getSheet());
    RegionUtil.setBorderLeft(
        cell.getCellStyle().getBorderLeft(), cellRangeAddress, cell.getSheet());
    RegionUtil.setBorderRight(
        cell.getCellStyle().getBorderRight(), cellRangeAddress, cell.getSheet());
    RegionUtil.setBorderBottom(
        cell.getCellStyle().getBorderBottom(), cellRangeAddress, cell.getSheet());
  }

  private CellStyle createCellStyle(CellStyleAnnotation style) {
    FontAnnotation fontAnnotation = style.getFont();
    Font font = this.workbook.createFont();
    fontAnnotation.applyFont(font);

    CellStyle cellStyle = this.workbook.createCellStyle();
    style.applyStyle(cellStyle, font, this.workbook);
    return cellStyle;
  }

  private Object findCellValue(CellStructure cellStructure) {
    Field sheetField = cellStructure.getSheetField();
    try {
      Object sheetObj = sheetField.get(this.excelDto);
      if (sheetObj == null) {
        return null;
      }
      Field rowField = cellStructure.getRowField();
      Object rowObj = rowField.get(sheetObj);
      if (rowObj == null) {
        return null;
      }
      Field cellField = cellStructure.getField();
      return cellField.get(rowObj);
    } catch (IllegalAccessException e) {
      throw new ExcelGenerateException(
          String.format("can not find cell class, %s", cellStructure.getFieldName()), e);
    }
  }

  private void bindCellValue(Cell cell, CellType cellType, Object value) {
    this.bindCellValue(cell, cellType, value, null, 0);
  }

  private void bindCellValue(Cell cell, CellType cellType, Object value,
      CellStructure cellStructure, int rowIndex) {
    if (value == null) {
      return;
    }
    switch (cellType) {
      case BLANK:
        cell.setBlank();
        break;
      case STRING:
        if (value instanceof String) {
          cell.setCellValue((String) value);
        }
        break;
      case NUMERIC:
        if (value instanceof Number) {
          cell.setCellValue(((Number) value).doubleValue());
        }
        break;
      case BOOLEAN:
        if (value instanceof Boolean) {
          cell.setCellValue((Boolean) value);
        }
        break;
      case FORMULA:
        if (value instanceof String && cellStructure != null) {
          this.formulaHelper.applyFormula(cell, (String) value, cellStructure, rowIndex);
        }
        break;
      case DATE:
        if (value instanceof LocalDate) {
          cell.setCellValue(DateFormatHelper
              .getDate((LocalDate) value, this.structure.getAnnotation().getDateFormatZoneId()));
        } else if (value instanceof LocalDateTime) {
          cell.setCellValue(DateFormatHelper
              .getDate((LocalDateTime) value,
                  this.structure.getAnnotation().getDateFormatZoneId()));
        }
        break;
      default:
        break;
    }
  }
}
