package com.mz.poi.mapper.structure;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public enum CellType {
    NONE,
    NUMERIC,
    STRING,
    FORMULA,
    BLANK,
    BOOLEAN,
    DATE;

    public org.apache.poi.ss.usermodel.CellType toExcelCellType() {
        switch (this) {
            case NUMERIC:
                return org.apache.poi.ss.usermodel.CellType.NUMERIC;
            case STRING:
            case DATE:
                return org.apache.poi.ss.usermodel.CellType.STRING;
            case FORMULA:
                return org.apache.poi.ss.usermodel.CellType.FORMULA;
            case BLANK:
                return org.apache.poi.ss.usermodel.CellType.BLANK;
            case BOOLEAN:
                return org.apache.poi.ss.usermodel.CellType.BOOLEAN;
            default:
                return org.apache.poi.ss.usermodel.CellType._NONE;
        }
    }

    public Cell createCell(Row row, int column) {
        org.apache.poi.ss.usermodel.CellType cellType = this.toExcelCellType();
        if (org.apache.poi.ss.usermodel.CellType._NONE.equals(cellType)) {
            return row.createCell(column);
        } else {
            return row.createCell(column, this.toExcelCellType());
        }
    }
}
