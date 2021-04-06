package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.exception.ExcelGenerateException;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;

@Getter
@NoArgsConstructor
public class CellStructure {

    private RowStructure rowStructure;
    private AbstractCellAnnotation annotation;
    private Field sheetField;
    private Field rowField;
    private Field field;
    private String fieldName;

    public void setAnnotation(AbstractCellAnnotation annotation) {
        this.annotation = annotation;
    }

    public boolean isAfterColumn() {
        String columnAfter = this.annotation.getColumnAfter();
        return columnAfter != null && columnAfter.length() > 0;
    }

    public boolean isArrayCell() {
        return this.annotation instanceof ArrayCellAnnotation;
    }

    private boolean calculated;

    public void setCalculated(boolean calculated) {
        this.calculated = calculated;
    }

    public void calculateColumn() {
        if (this.isAfterColumn()) {
            CellStructure beforeCellStructure = this.rowStructure.findCellByFieldName(
                this.annotation.getColumnAfter()
            );
            if (!beforeCellStructure.calculated) {
                beforeCellStructure.calculateColumn();
            }
            this.annotation.setColumn(
                beforeCellStructure.getAnnotation().getColumn() +
                    (beforeCellStructure.getAnnotation().getColumnSize() *
                        beforeCellStructure.getAnnotation().getCols()) +
                    this.annotation.getColumnAfterOffset());
        }
        this.calculated = true;
    }

    public Object findCellValue(Object item) {
        try {
            if (item == null) {
                return null;
            }
            Field cellField = this.field;
            return cellField.get(item);
        } catch (IllegalAccessException e) {
            throw new ExcelGenerateException(
                String.format("can not find cell class, %s", this.fieldName), e);
        }
    }

    @SuppressWarnings("unchecked")
    public <T> Collection<T> findCellDataCollection(Object item) {
        try {
            if (item == null) {
                return new ArrayList<>();
            }
            Object cellObj = this.field.get(item);
            if (cellObj == null) {
                return new ArrayList<>();
            }
            return (Collection<T>) cellObj;
        } catch (IllegalAccessException e) {
            throw new ExcelGenerateException(
                String.format("can not find data row collection, %s", this.fieldName), e);
        }
    }

    @Builder
    public CellStructure(
        RowStructure rowStructure,
        CellAnnotation annotation, Field sheetField, Field rowField, Field field,
        String fieldName) {
        this.rowStructure = rowStructure;
        this.sheetField = sheetField;
        this.rowField = rowField;
        this.annotation = annotation;
        this.field = field;
        this.fieldName = fieldName;
    }
}
