package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.ColumnWidth;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class ColumnWidthAnnotation {

    private int column;
    private int width;

    public ColumnWidthAnnotation(ColumnWidth columnWidth) {
        this.column = columnWidth.column();
        this.width = columnWidth.width();
    }
}
