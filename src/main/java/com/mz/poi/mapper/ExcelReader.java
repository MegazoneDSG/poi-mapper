package com.mz.poi.mapper;

import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.exception.ExcelReadException;
import com.mz.poi.mapper.exception.ReadExceptionAddress;
import com.mz.poi.mapper.expression.CustomCellExpression;
import com.mz.poi.mapper.helper.DateFormatHelper;
import com.mz.poi.mapper.helper.FormulaHelper;
import com.mz.poi.mapper.helper.InheritedFieldHelper;
import com.mz.poi.mapper.structure.ArrayCellAnnotation;
import com.mz.poi.mapper.structure.CellAnnotation;
import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.CellType;
import com.mz.poi.mapper.structure.DataRowsAnnotation;
import com.mz.poi.mapper.structure.ExcelStructure;
import com.mz.poi.mapper.structure.RowStructure;
import com.mz.poi.mapper.structure.SheetAnnotation;
import com.mz.poi.mapper.structure.SheetStructure;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.BeanUtils;

import java.lang.reflect.Constructor;
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
import java.util.stream.IntStream;

@Getter
@Slf4j
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
            if (cellStructure.getAnnotation() instanceof ArrayCellAnnotation) {
                this.readArrayCells(cellStructure, row, rowObj);
            } else {
                this.readCell(cellStructure, row, rowObj);
            }
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
            .flatMapToInt(cellStructure -> {
                int column = cellStructure.getAnnotation().getColumn();
                int columnSize = cellStructure.getAnnotation().getColumnSize();
                int cols = cellStructure.getAnnotation().getCols();
                if (columnSize == 0) {
                    return IntStream.empty();
                } else {
                    return IntStream.range(column, column + (columnSize * cols));
                }
            })
            .anyMatch(column -> {
                Cell cell = row.getCell(column);
                return !(cell == null
                    || cell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK);
            });
        List<CellStructure> cells = rowStructure.getCells();
        cells.forEach(cellStructure -> {
            if (cellStructure.getAnnotation() instanceof ArrayCellAnnotation) {
                this.readArrayCells(cellStructure, row, rowObj);
            } else {
                this.readCell(cellStructure, row, rowObj);
            }
        });
        return isMatch;
    }

    private boolean isRequiredMatch(RowStructure rowStructure, Row row, Object rowObj) {
        AtomicBoolean isMatch = new AtomicBoolean(true);
        rowStructure.getCells().forEach(cellStructure -> {
            boolean isBind;
            if (cellStructure.getAnnotation() instanceof ArrayCellAnnotation) {
                isBind = this.readArrayCells(cellStructure, row, rowObj);
            } else {
                isBind = this.readCell(cellStructure, row, rowObj);
            }
            if (cellStructure.getAnnotation().isRequired() && !isBind) {
                isMatch.set(false);
            }
        });
        return isMatch.get();
    }

    private boolean isAllMatch(RowStructure rowStructure, Row row, Object rowObj) {
        AtomicBoolean isMatch = new AtomicBoolean(true);
        rowStructure.getCells().forEach(cellStructure -> {
            boolean isBind;
            if (cellStructure.getAnnotation() instanceof ArrayCellAnnotation) {
                isBind = this.readArrayCells(cellStructure, row, rowObj);
            } else {
                isBind = this.readCell(cellStructure, row, rowObj);
            }
            if (!isBind) {
                isMatch.set(false);
            }
        });
        return isMatch.get();
    }

    private boolean readCell(CellStructure cellStructure, Row row, Object rowObj) {
        CellAnnotation cellAnnotation = (CellAnnotation) cellStructure.getAnnotation();
        Cell cell = row.getCell(cellAnnotation.getColumn());
        return this.bindCellValue(cell, cellStructure, rowObj);
    }

    private boolean readArrayCells(CellStructure cellStructure, Row row, Object rowObj) {
        List collection = new ArrayList<>();
        Class<?> cellClass = null;
        Field collectionField = null;
        try {
            collectionField = InheritedFieldHelper
                .getDeclaredField(rowObj.getClass(), cellStructure.getFieldName());

            collectionField.setAccessible(true);
            collectionField.set(rowObj, collection);

            ParameterizedType genericType =
                (ParameterizedType) cellStructure.getField().getGenericType();
            cellClass = (Class<?>) genericType.getActualTypeArguments()[0];

        } catch (IllegalAccessException | NoSuchFieldException e) {
            throw new ExcelReadException("Invalid array cell class", e,
                new ReadExceptionAddress(
                    cellStructure.getRowStructure().getSheetStructure().getAnnotation().getIndex(),
                    cellStructure.getRowStructure().getStartRowNum())
            );
        }

        AtomicBoolean isMatch = new AtomicBoolean(true);
        ArrayCellAnnotation arrayCellAnnotation = (ArrayCellAnnotation) cellStructure.getAnnotation();
        Class<?> finalCellClass = cellClass;
        IntStream.range(0, arrayCellAnnotation.getSize()).forEach(index -> {
            Cell cell = row.getCell(
                arrayCellAnnotation.getColumn() + (arrayCellAnnotation.getCols() * index));
            boolean isBind = this.bindArrayCellValue(cell, cellStructure, collection, finalCellClass);
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
            Field cellField = cellStructure.getField();
            cellField.setAccessible(true);
            Class<?> cellClass = cellField.getType();
            CellType cellType = cellStructure.getAnnotation().getCellType();

            if (CustomCellExpression.class.isAssignableFrom(cellClass)) {
                Field field = InheritedFieldHelper
                    .getDeclaredField(cellClass, "value");
                Class<?> type = field.getType();
                Constructor<?> constructor = cellClass.getConstructor(type);
                CustomCellExpression customCellExpression = (CustomCellExpression) BeanUtils.instantiateClass(constructor, (Object) null);
                customCellExpression.fromCell(cell);
                isBind.set(true);
            } else {
                Object cellValue = this.getCellValueByCellType(cell, cellType, cellClass);
                if (cellValue != null) {
                    cellField.set(rowObj, cellValue);
                    isBind.set(true);
                }
            }
            return isBind.get();
        } catch (IllegalAccessException | NoSuchFieldException | NoSuchMethodException e) {
            throw new ExcelReadException("can not set cell value", e,
                new ReadExceptionAddress(
                    this.workbook.getSheetIndex(cell.getSheet().getSheetName()),
                    cell.getRowIndex(),
                    cell.getColumnIndex()
                )
            );
        }
    }

    @SuppressWarnings("unchecked")
    private boolean bindArrayCellValue(Cell cell, CellStructure cellStructure,
                                       List collection, Class<?> cellClass) {
        try {
            if (!Optional.ofNullable(cell).isPresent() || cellStructure.getAnnotation().isIgnoreParse()) {
                return false;
            }
            AtomicBoolean isBind = new AtomicBoolean(false);
            CellType cellType = cellStructure.getAnnotation().getCellType();

            if (CustomCellExpression.class.isAssignableFrom(cellClass)) {
                Field field = InheritedFieldHelper
                    .getDeclaredField(cellClass, "value");
                Class<?> type = field.getType();
                Constructor<?> constructor = cellClass.getConstructor(type);
                CustomCellExpression customCellExpression = (CustomCellExpression) BeanUtils.instantiateClass(constructor, (Object) null);
                customCellExpression.fromCell(cell);
                collection.add(customCellExpression);
                isBind.set(true);
            } else {
                Object cellValue = this.getCellValueByCellType(cell, cellType, cellClass);
                if (cellValue != null) {
                    collection.add(cellValue);
                    isBind.set(true);
                } else {
                    collection.add(null);
                }
            }
            return isBind.get();
        } catch (NoSuchFieldException | NoSuchMethodException e) {
            throw new ExcelReadException("can not set cell value", e,
                new ReadExceptionAddress(
                    this.workbook.getSheetIndex(cell.getSheet().getSheetName()),
                    cell.getRowIndex(),
                    cell.getColumnIndex()
                )
            );
        }
    }

    private Object getCellValueByCellType(Cell cell, CellType cellType, Class<?> cellClass) {
        boolean isDate;
        try {
            isDate = DateUtil.isCellDateFormatted(cell);
        } catch (IllegalStateException ex) {
            isDate = false;
        }
        boolean isString = cell.getCellType().equals(org.apache.poi.ss.usermodel.CellType.STRING);
        boolean isNumber = cell.getCellType().equals(org.apache.poi.ss.usermodel.CellType.NUMERIC);
        boolean isBoolean = cell.getCellType().equals(org.apache.poi.ss.usermodel.CellType.BOOLEAN);
        if (isDate && (CellType.NONE.equals(cellType) || CellType.DATE.equals(cellType))) {

            Date cellValue = cell.getDateCellValue();
            if (cellValue == null) {
                return null;
            } else if (LocalDateTime.class.isAssignableFrom(cellClass)) {
                return DateFormatHelper.getLocalDateTime(cellValue,
                    this.structure.getAnnotation().getDateFormatZoneId());
            } else if (LocalDate.class.isAssignableFrom(cellClass)) {
                return DateFormatHelper.getLocalDate(cellValue,
                    this.structure.getAnnotation().getDateFormatZoneId());
            } else if (Object.class.isAssignableFrom(cellClass)) {
                return cellValue;
            } else {
                throw new ExcelReadException(
                    String.format("not supported date type %s", cellClass.getName()),
                    new ReadExceptionAddress(
                        this.workbook.getSheetIndex(cell.getSheet().getSheetName()),
                        cell.getRowIndex(),
                        cell.getColumnIndex()
                    )
                );
            }
        } else if (isString && (CellType.NONE.equals(cellType) || CellType.STRING.equals(cellType))) {
            if (String.class.isAssignableFrom(cellClass)) {
                return cell.getStringCellValue();
            } else if (Object.class.isAssignableFrom(cellClass)) {
                return cell.getStringCellValue();
            }
        } else if (isNumber && (CellType.NONE.equals(cellType) || CellType.NUMERIC.equals(cellType))) {
            double cellValue = cell.getNumericCellValue();
            if (Double.class.isAssignableFrom(cellClass) || double.class.isAssignableFrom(cellClass)) {
                return cellValue;
            } else if (Float.class.isAssignableFrom(cellClass) || float.class.isAssignableFrom(cellClass)) {
                return (float) cellValue;
            } else if (Long.class.isAssignableFrom(cellClass) || long.class.isAssignableFrom(cellClass)) {
                return Double.valueOf(cellValue).longValue();
            } else if (Short.class.isAssignableFrom(cellClass) || short.class.isAssignableFrom(cellClass)) {
                return Double.valueOf(cellValue).shortValue();
            } else if (BigDecimal.class.isAssignableFrom(cellClass)) {
                return BigDecimal.valueOf(cellValue);
            } else if (BigInteger.class.isAssignableFrom(cellClass)) {
                return BigInteger.valueOf(Double.valueOf(cellValue).longValue());
            } else if (Integer.class.isAssignableFrom(cellClass) || int.class.isAssignableFrom(cellClass)) {
                return Double.valueOf(cellValue).intValue();
            } else if (Object.class.isAssignableFrom(cellClass)) {
                return cellValue;
            } else {
                throw new ExcelReadException(
                    String.format("not supported number type %s", cellClass.getName()),
                    new ReadExceptionAddress(
                        this.workbook.getSheetIndex(cell.getSheet().getSheetName()),
                        cell.getRowIndex(),
                        cell.getColumnIndex()
                    )
                );
            }
        } else if (isBoolean && (CellType.NONE.equals(cellType) || CellType.BOOLEAN.equals(cellType))) {
            if (Boolean.class.isAssignableFrom(cellClass) || boolean.class.isAssignableFrom(cellClass)) {
                return cell.getBooleanCellValue();
            } else if (Object.class.isAssignableFrom(cellClass)) {
                return cell.getBooleanCellValue();
            }
        }
        return null;
    }
}
