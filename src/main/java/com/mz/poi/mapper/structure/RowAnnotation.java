package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Row;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class RowAnnotation extends AbstractRowAnnotation {

    private boolean useRowHeightInPoints = false;
    private int heightInPoints;
    private CellStyleAnnotation defaultStyle;

    public RowAnnotation(Row row, CellStyleAnnotation sheetStyle) {
        this.setRow(row.row());
        this.setRowAfter(row.rowAfter());
        this.setRowAfterOffset(row.rowAfterOffset());

        if (row.heightInPoints().length > 0) {
            this.useRowHeightInPoints = true;
            this.heightInPoints = row.heightInPoints()[0];
        }
        this.defaultStyle = new CellStyleAnnotation(row.defaultStyle(), sheetStyle);
    }
}
