package com.mz.poi.mapper.helper;

import com.mz.poi.mapper.exception.ExcelGenerateException;
import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.RowStructure;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;

import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FormulaHelper {

    public static void applyFormula(Cell cell, String formulaExpresion, CellStructure cellStructure,
                                    int rowIndex) {
        String addressRx = "\\{\\{(.+?)}}";
        String formula = formulaExpresion;
        Matcher m = Pattern.compile(addressRx).matcher(formulaExpresion);
        while (m.find()) {
            String addressToReplace = m.group(0);
            String addressStr = m.group(1);
            String[] split = addressStr.split("\\.");
            Optional<String> cellAddress = Optional.empty();
            if (split.length == 2 && "this".equals(split[0])) {
                cellAddress = getSelfDataRowCellAddress(cellStructure, addressStr, rowIndex);
            } else if (split.length == 2) {
                cellAddress = getCellAddress(cellStructure, addressStr);
            }
            String finalAddress = cellAddress.orElseThrow(() ->
                new ExcelGenerateException(String.format("Not found cellAddress %s", addressStr)));
            formula = formula.replace(addressToReplace, finalAddress);
        }
        cell.setCellFormula(formula);
    }

    private static Optional<String> getCellAddress(CellStructure cellStructure, String addressStr) {
        String[] split = addressStr.split("\\.");
        String rowNumRex = getIndexRex(split[0]);
        String rowFieldName =
            rowNumRex == null ? split[0] : split[0].replace("[" + rowNumRex + "]", "");
        String cellNumRex = getIndexRex(split[1]);
        String cellFieldName =
            cellNumRex == null ? split[1] : split[1].replace("[" + cellNumRex + "]", "");

        RowStructure rowStructure = cellStructure.getRowStructure().getSheetStructure()
            .findRowByFieldName(rowFieldName);
        if (!rowStructure.isDataRow() && rowNumRex != null) {
            throw new ExcelGenerateException(String.format("Invalid formula pattern %s", addressStr));
        }
        int rowNum = 0;
        if (rowNumRex == null) {
            rowNum = rowStructure.getStartRowNum();
        } else if ("last".equals(rowNumRex)) {
            rowNum = rowStructure.getEndRowNum();
        } else {
            rowNum = rowStructure.getStartRowNum() + Integer.parseInt(rowNumRex) +
                (rowStructure.isDataRowAndHideHeader() ? 0 : 1);
        }

        CellStructure cellByFieldName = rowStructure.findCellByFieldName(cellFieldName);
        if (!cellByFieldName.isArrayCell() && cellNumRex != null) {
            throw new ExcelGenerateException(String.format("Invalid formula pattern %s", addressStr));
        }
        int cellNum = 0;
        int column = cellByFieldName.getAnnotation().getColumn();
        int columnSize = cellByFieldName.getAnnotation().getColumnSize();
        int cols = cellByFieldName.getAnnotation().getCols();
        if (cellNumRex == null) {
            cellNum = column;
        } else if ("last".equals(cellNumRex)) {
            cellNum = column + ((columnSize - 1) * cols);
        } else {
            cellNum = column + ((Integer.parseInt(cellNumRex)) * cols);
        }
        return Optional.of(new CellAddress(rowNum, cellNum).formatAsString());
    }

    private static Optional<String> getSelfDataRowCellAddress(CellStructure cellStructure,
                                                              String addressStr, int rowIndex) {
        String[] split = addressStr.split("\\.");
        String cellNumRex = getIndexRex(split[1]);
        String cellFieldName =
            cellNumRex == null ? split[1] : split[1].replace("[" + cellNumRex + "]", "");

        RowStructure rowStructure = cellStructure.getRowStructure();
        CellStructure cellByFieldName = rowStructure.findCellByFieldName(cellFieldName);
        if (!cellByFieldName.isArrayCell() && cellNumRex != null) {
            throw new ExcelGenerateException(String.format("Invalid formula pattern %s", addressStr));
        }
        int cellNum = 0;
        int column = cellByFieldName.getAnnotation().getColumn();
        int columnSize = cellByFieldName.getAnnotation().getColumnSize();
        int cols = cellByFieldName.getAnnotation().getCols();
        if (cellNumRex == null) {
            cellNum = column;
        } else if ("last".equals(cellNumRex)) {
            cellNum = column + ((columnSize - 1) * cols);
        } else {
            cellNum = column + ((Integer.parseInt(cellNumRex)) * cols);
        }
        return Optional.of(new CellAddress(rowIndex, cellNum).formatAsString());
    }

    public static String getIndexRex(String fieldRex) {
        String indexRex = "\\[(.+?)]";
        Matcher m = Pattern.compile(indexRex).matcher(fieldRex);
        String group = null;
        while (m.find()) {
            group = m.group(1);
        }
        if (group != null && !"last".equals(group)) {
            try {
                Integer.parseInt(group);
            } catch (NumberFormatException e) {
                throw new ExcelGenerateException(String.format("Invalid formula pattern %s", fieldRex));
            }
        }
        return group;
    }
}
