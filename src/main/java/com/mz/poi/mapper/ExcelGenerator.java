package com.mz.poi.mapper;

import com.mz.poi.mapper.exception.ExcelGenerateException;
import com.mz.poi.mapper.expression.ArrayHeaderNameExpression;
import com.mz.poi.mapper.helper.DateFormatHelper;
import com.mz.poi.mapper.helper.FormulaHelper;
import com.mz.poi.mapper.structure.AbstractCellAnnotation;
import com.mz.poi.mapper.structure.AbstractHeaderAnnotation;
import com.mz.poi.mapper.structure.ArrayCellAnnotation;
import com.mz.poi.mapper.structure.ArrayHeaderAnnotation;
import com.mz.poi.mapper.structure.CellAnnotation;
import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.CellStyleAnnotation;
import com.mz.poi.mapper.structure.CellType;
import com.mz.poi.mapper.structure.DataRowsAnnotation;
import com.mz.poi.mapper.structure.ExcelStructure;
import com.mz.poi.mapper.structure.FontAnnotation;
import com.mz.poi.mapper.structure.HeaderAnnotation;
import com.mz.poi.mapper.structure.RowAnnotation;
import com.mz.poi.mapper.structure.RowStructure;
import com.mz.poi.mapper.structure.SheetAnnotation;
import com.mz.poi.mapper.structure.SheetStructure;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
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
    });
    return this.workbook;
  }

  private void drawRow(RowStructure rowStructure, Sheet sheet) {
    int rowIndex = rowStructure.getStartRowNum();
    Row row = sheet.createRow(rowIndex);

    RowAnnotation rowAnnotation = (RowAnnotation) rowStructure.getAnnotation();
    if (rowAnnotation.isUseRowHeightInPoints()) {
      row.setHeightInPoints(rowAnnotation.getHeightInPoints());
    }
    Object rowData = rowStructure.findRowData(excelDto);
    List<CellStructure> cells = rowStructure.getCells();
    cells.forEach(cellStructure -> {
      if (cellStructure.getAnnotation() instanceof ArrayCellAnnotation) {
        this.drawArrayCell(cellStructure, row, rowData, rowIndex);
      } else {
        this.drawCell(cellStructure, row, rowData, rowIndex);
      }
    });
  }

  private void drawDataRows(RowStructure rowStructure, Sheet sheet) {
    AtomicInteger currentRowNum = new AtomicInteger(rowStructure.getStartRowNum());
    DataRowsAnnotation annotation = (DataRowsAnnotation) rowStructure.getAnnotation();

    // draw header
    if (rowStructure.isDataRowAndHideHeader()) {
      currentRowNum.decrementAndGet();
    } else {
      this.drawDataHeaderRow(rowStructure, currentRowNum.get(), sheet);
    }

    // draw cachedDataRowStyle
    Map<String, CellStyle> cachedDataRowStyle = this.createCachedDataRowStyle(rowStructure);
    Collection<Object> rowDataList = rowStructure.findRowDataCollection(this.excelDto);
    if (rowDataList.isEmpty()) {
      return;
    }
    rowDataList.forEach(rowData ->
        this.drawDataRow(rowStructure, currentRowNum.incrementAndGet(), rowData, cachedDataRowStyle,
            sheet));
  }

  private void drawDataRow(
      RowStructure rowStructure, int rowIndex, Object rowData,
      Map<String, CellStyle> cachedDataRowStyle, Sheet sheet) {

    DataRowsAnnotation annotation = (DataRowsAnnotation) rowStructure.getAnnotation();

    Row row = sheet.createRow(rowIndex);
    if (annotation.isUseDataHeightInPoints()) {
      row.setHeightInPoints(annotation.getDataHeightInPoints());
    }

    List<CellStructure> cells = rowStructure.getCells();
    cells.forEach(cellStructure -> {
      if (cellStructure.getAnnotation() instanceof ArrayCellAnnotation) {
        this.drawArrayCell(cellStructure, row, rowData, rowIndex, cachedDataRowStyle);
      } else {
        this.drawCell(cellStructure, row, rowData, rowIndex, cachedDataRowStyle);
      }
    });

  }

  private void drawDataHeaderRow(RowStructure rowStructure, int rowNum, Sheet sheet) {
    Row row = sheet.createRow(rowNum);
    DataRowsAnnotation annotation = (DataRowsAnnotation) rowStructure.getAnnotation();
    if (annotation.isUseHeaderHeightInPoints()) {
      row.setHeightInPoints(annotation.getHeaderHeightInPoints());
    }
    ArrayList<AbstractHeaderAnnotation> list = new ArrayList<>();
    list.addAll(annotation.getArrayHeaders());
    list.addAll(annotation.getHeaders());
    list.forEach(abstractHeaderAnnotation -> {
      CellStyle cellStyle = this.createCellStyle(abstractHeaderAnnotation.getStyle());

      // 어레이셀 헤더 - 어레이셀과 매핑되며, 어레이셀의 크기만큼 헤더를 생성한다.
      if (abstractHeaderAnnotation instanceof ArrayHeaderAnnotation) {
        ArrayHeaderAnnotation arrayHeaderAnnotation = (ArrayHeaderAnnotation) abstractHeaderAnnotation;
        CellStructure cellStructure = rowStructure
            .findCellByFieldName(arrayHeaderAnnotation.getMapping());
        if (!cellStructure.isArrayCell()) {
          throw new ExcelGenerateException(
              String.format("array header should mapping array cell, %s",
                  arrayHeaderAnnotation.getMapping()));
        }
        int column = cellStructure.getAnnotation().getColumn();
        int size = cellStructure.getAnnotation().getColumnSize();
        int cols = cellStructure.getAnnotation().getCols();
        IntStream.range(0, size).forEach(index -> {
          int currentColumn = column + (index * cols);
          Cell cell = row.createCell(currentColumn, org.apache.poi.ss.usermodel.CellType.STRING);
          cell.setCellStyle(cellStyle);
          this.mergeCell(cell, currentColumn, cols);

          String name;
          if (arrayHeaderAnnotation.getArrayHeaderNameExpression() != null) {
            name = arrayHeaderAnnotation.getArrayHeaderNameExpression().get(index);
          } else {
            name = arrayHeaderAnnotation.getSimpleNameExpression()
                .replaceAll("\\{\\{index}}", Integer.toString(index));
          }
          this.bindCellValue(cell, CellType.STRING, name);
        });
      }
      // 일반 헤더 - 복수의 셀과 매핑될 수 있으며, 복수의 셀이 차지하는 영역만큼 병합되어 표현된다.
      else {
        HeaderAnnotation headerAnnotation = (HeaderAnnotation) abstractHeaderAnnotation;
        List<CellStructure> mappingCellStructures = headerAnnotation.getMappings().stream()
            .map(rowStructure::findCellByFieldName)
            .sorted(
                Comparator.comparing(cellStructure -> cellStructure.getAnnotation().getColumn()))
            .collect(Collectors.toList());
        if (mappingCellStructures.isEmpty()) {
          return;
        }
        int column = mappingCellStructures.get(0).getAnnotation().getColumn();
        AbstractCellAnnotation lastMappingAnnotation = mappingCellStructures
            .get(mappingCellStructures.size() - 1).getAnnotation();
        int cols = lastMappingAnnotation.getColumn() +
            (lastMappingAnnotation.getColumnSize() * lastMappingAnnotation.getCols()) - column;
        Cell cell = row.createCell(column, org.apache.poi.ss.usermodel.CellType.STRING);
        cell.setCellStyle(cellStyle);
        this.mergeCell(cell, column, cols);
        this.bindCellValue(cell, CellType.STRING, headerAnnotation.getName());
      }
    });
  }

  private Map<String, CellStyle> createCachedDataRowStyle(RowStructure rowStructure) {
    List<CellStructure> cells = rowStructure.getCells();
    Map<String, CellStyle> cachedDataRowStyle = new HashMap<>();
    cells.forEach(cellStructure -> {
      AbstractCellAnnotation annotation = cellStructure.getAnnotation();
      cachedDataRowStyle.put(
          cellStructure.getFieldName(),
          this.createCellStyle(annotation.getStyle()));
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

  private void drawCell(CellStructure cellStructure, Row row, Object rowData, int rowIndex) {
    this.drawCell(cellStructure, row, rowData, rowIndex, null);
  }

  private void drawCell(CellStructure cellStructure, Row row, Object rowData, int rowIndex,
      Map<String, CellStyle> cachedDataRowStyle) {
    CellAnnotation annotation = (CellAnnotation) cellStructure.getAnnotation();
    Cell cell = row.createCell(
        annotation.getColumn(), annotation.getCellType().toExcelCellType()
    );

    //스타일 적용
    if (cellStructure.getRowStructure().isDataRow()) {
      this.getCachedDataRowStyle(cachedDataRowStyle, cellStructure.getFieldName())
          .ifPresent(cell::setCellStyle);
    } else {
      cell.setCellStyle(this.createCellStyle(annotation.getStyle()));
    }

    //cols 적용
    this.mergeCell(cell, annotation.getColumn(), annotation.getCols());
    //값 바인딩
    Object cellValue = cellStructure.findCellValue(rowData);
    this.bindCellValue(cell, annotation.getCellType(), cellValue, cellStructure, rowIndex);
  }

  private void drawArrayCell(CellStructure cellStructure, Row row, Object rowData, int rowIndex) {
    this.drawArrayCell(cellStructure, row, rowData, rowIndex);
  }

  private void drawArrayCell(CellStructure cellStructure, Row row, Object rowData, int rowIndex,
      Map<String, CellStyle> cachedDataRowStyle) {
    ArrayCellAnnotation annotation = (ArrayCellAnnotation) cellStructure.getAnnotation();
    AtomicInteger currentColumn = new AtomicInteger(annotation.getColumn());

    //스타일 적용
    CellStyle cellStyle;
    //스타일 적용
    if (cellStructure.getRowStructure().isDataRow()) {
      cellStyle = this.getCachedDataRowStyle(cachedDataRowStyle, cellStructure.getFieldName())
          .orElse(this.createCellStyle(annotation.getStyle()));
    } else {
      cellStyle = this.createCellStyle(annotation.getStyle());
    }

    Collection<?> cellDataList = cellStructure.findCellDataCollection(rowData)
        .stream()
        .limit(annotation.getSize())
        .collect(Collectors.toList());
    if (cellDataList.isEmpty()) {
      return;
    }
    cellDataList.forEach(cellData -> {
      Cell cell = row.createCell(
          currentColumn.get(), annotation.getCellType().toExcelCellType());
      cell.setCellStyle(cellStyle);
      //cols 적용
      this.mergeCell(cell, currentColumn.get(), annotation.getCols());
      //값 바인딩
      this.bindCellValue(cell, annotation.getCellType(), cellData, cellStructure, rowIndex);
      //컬럼 증가
      currentColumn.set(currentColumn.get() + annotation.getCols());
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
