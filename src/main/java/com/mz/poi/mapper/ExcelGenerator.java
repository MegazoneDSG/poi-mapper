package com.mz.poi.mapper;

import com.mz.poi.mapper.exception.ExcelGenerateException;
import com.mz.poi.mapper.expression.CellGenerator;
import com.mz.poi.mapper.expression.CustomCellExpression;
import com.mz.poi.mapper.helper.CachedStyleRepository;
import com.mz.poi.mapper.helper.DateFormatHelper;
import com.mz.poi.mapper.helper.FormulaHelper;
import com.mz.poi.mapper.structure.AbstractCellAnnotation;
import com.mz.poi.mapper.structure.AbstractHeaderAnnotation;
import com.mz.poi.mapper.structure.ArrayCellAnnotation;
import com.mz.poi.mapper.structure.ArrayHeaderAnnotation;
import com.mz.poi.mapper.structure.CellAnnotation;
import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.CellType;
import com.mz.poi.mapper.structure.DataRowsAnnotation;
import com.mz.poi.mapper.structure.ExcelStructure;
import com.mz.poi.mapper.structure.HeaderAnnotation;
import com.mz.poi.mapper.structure.RowAnnotation;
import com.mz.poi.mapper.structure.RowStructure;
import com.mz.poi.mapper.structure.SheetAnnotation;
import com.mz.poi.mapper.structure.SheetStructure;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Comparator;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Getter
@Slf4j
public class ExcelGenerator {

    private Workbook workbook;
    private ExcelStructure structure;
    private Object excelDto;
    private CachedStyleRepository styleRepository;

    public ExcelGenerator(Object excelDto, Workbook workbook) {
        this.workbook = workbook;
        this.excelDto = excelDto;
        this.styleRepository = new CachedStyleRepository(this.workbook);
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
        log.info("styleRepository.getSize() = {}", this.styleRepository.getSize());
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

        Collection<Object> rowDataList = rowStructure.findRowDataCollection(this.excelDto);
        if (rowDataList.isEmpty()) {
            return;
        }
        rowDataList.forEach(rowData ->
            this.drawDataRow(rowStructure, currentRowNum.incrementAndGet(), rowData, sheet));
    }

    private void drawDataRow(
        RowStructure rowStructure, int rowIndex, Object rowData, Sheet sheet) {

        DataRowsAnnotation annotation = (DataRowsAnnotation) rowStructure.getAnnotation();

        Row row = sheet.createRow(rowIndex);
        if (annotation.isUseDataHeightInPoints()) {
            row.setHeightInPoints(annotation.getDataHeightInPoints());
        }

        List<CellStructure> cells = rowStructure.getCells();
        cells.forEach(cellStructure -> {
            if (cellStructure.getAnnotation() instanceof ArrayCellAnnotation) {
                this.drawArrayCell(cellStructure, row, rowData, rowIndex);
            } else {
                this.drawCell(cellStructure, row, rowData, rowIndex);
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
            CellStyle cellStyle = this.styleRepository.createStyle(abstractHeaderAnnotation.getStyle());

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
                    this.bindHeaderCellValue(cell, name);
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
                this.bindHeaderCellValue(cell, headerAnnotation.getName());
            }
        });
    }

    private void drawCell(CellStructure cellStructure, Row row, Object rowData, int rowIndex) {
        CellAnnotation annotation = (CellAnnotation) cellStructure.getAnnotation();
        Object cellValue = cellStructure.findCellValue(rowData);

        CellGenerator cellGenerator;
        if (cellValue instanceof CustomCellExpression) {
            cellGenerator = ((CustomCellExpression) cellValue).toCell(annotation.getStyle(), this.styleRepository);
        } else {
            cellGenerator = CellGenerator.builder()
                .cellType(annotation.getCellType())
                .style(this.styleRepository.createStyle(annotation.getStyle()))
                .value(cellValue)
                .build();
        }
        Cell cell = cellGenerator.getCellType().createCell(row, annotation.getColumn());
        //스타일 적용
        cell.setCellStyle(cellGenerator.getStyle());
        //cols 적용
        this.mergeCell(cell, annotation.getColumn(), annotation.getCols());
        //값 바인딩
        this.bindCellValue(cell, cellGenerator.getCellType(), cellGenerator.getValue(), cellStructure, rowIndex);
    }

    private void drawArrayCell(CellStructure cellStructure, Row row, Object rowData, int rowIndex) {
        ArrayCellAnnotation annotation = (ArrayCellAnnotation) cellStructure.getAnnotation();
        AtomicInteger currentColumn = new AtomicInteger(annotation.getColumn());

        Collection<?> cellDataList = cellStructure.findCellDataCollection(rowData)
            .stream()
            .limit(annotation.getSize())
            .collect(Collectors.toList());
        if (cellDataList.isEmpty()) {
            return;
        }
        if (cellDataList.size() < annotation.getSize()) {
            IntStream.range(0, annotation.getSize() - cellDataList.size())
                .forEach(i -> cellDataList.add(null));
        }

        cellDataList.forEach(cellData -> {

            CellGenerator cellGenerator;
            if (cellData instanceof CustomCellExpression) {
                cellGenerator = ((CustomCellExpression) cellData).toCell(annotation.getStyle(), this.styleRepository);
            } else {
                cellGenerator = CellGenerator.builder()
                    .cellType(annotation.getCellType())
                    .style(this.styleRepository.createStyle(annotation.getStyle()))
                    .value(cellData)
                    .build();
            }
            Cell cell = cellGenerator.getCellType().createCell(row, currentColumn.get());
            //스타일
            cell.setCellStyle(cellGenerator.getStyle());
            //cols 적용
            this.mergeCell(cell, currentColumn.get(), annotation.getCols());
            //값 바인딩
            this.bindCellValue(cell, cellGenerator.getCellType(), cellGenerator.getValue(), cellStructure, rowIndex);
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

    private void bindHeaderCellValue(Cell cell, Object value) {
        this.bindCellValue(cell, CellType.STRING, value, null, 0);
    }

    private void bindCellValue(Cell cell, CellType cellType, Object value,
                               CellStructure cellStructure, int rowIndex) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String && CellType.FORMULA.equals(cellType)) {
            FormulaHelper.applyFormula(cell, (String) value, cellStructure, rowIndex);
        } else if (value instanceof String && (CellType.NONE.equals(cellType) || CellType.STRING.equals(cellType))) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number && (CellType.NONE.equals(cellType) || CellType.NUMERIC.equals(cellType))) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean && (CellType.NONE.equals(cellType) || CellType.BOOLEAN.equals(cellType))) {
            cell.setCellValue((Boolean) value);
        } else if (CellType.NONE.equals(cellType) || CellType.DATE.equals(cellType)) {
            if (value instanceof LocalDate) {
                cell.setCellValue(DateFormatHelper
                    .getDate((LocalDate) value,
                        this.structure.getAnnotation().getDateFormatZoneId()));
            } else if (value instanceof LocalDateTime) {
                cell.setCellValue(DateFormatHelper
                    .getDate((LocalDateTime) value,
                        this.structure.getAnnotation().getDateFormatZoneId()));
            }
        }
    }
}
