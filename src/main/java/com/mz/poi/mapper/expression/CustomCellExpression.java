package com.mz.poi.mapper.expression;

import com.mz.poi.mapper.helper.CachedStyleRepository;
import com.mz.poi.mapper.structure.CellStyleAnnotation;
import com.mz.poi.mapper.structure.CellType;
import org.apache.poi.ss.usermodel.Cell;

public class CustomCellExpression<T> {

    private T value;

    public T getValue() {
        return value;
    }

    public void setValue(T value) {
        this.value = value;
    }

    public CustomCellExpression(T value) {
        this.value = value;
    }

    public CellGenerator toCell(CellStyleAnnotation styleAnnotation, CachedStyleRepository styleRepository) {
        return CellGenerator.builder()
            .cellType(CellType.NONE)
            .style(styleRepository.createStyle(styleAnnotation))
            .value(value)
            .build();
    }

    public void fromCell(Cell cell) {
        this.setValue(null);
    }
}
