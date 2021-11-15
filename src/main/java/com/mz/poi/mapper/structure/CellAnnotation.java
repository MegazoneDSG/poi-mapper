package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Cell;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class CellAnnotation extends AbstractCellAnnotation {

    public CellAnnotation(Cell cell, CellStyleAnnotation rowStyle) {
        this.setColumn(cell.column());
        this.setCols(cell.cols());
        this.setColumnAfter(cell.columnAfter());
        this.setColumnAfterOffset(cell.columnAfterOffset());
        this.setCellType(cell.cellType());
        this.setIgnoreParse(cell.ignoreParse());
        this.setRequired(cell.required());
        this.setStyle(new CellStyleAnnotation(cell.style(), rowStyle));
        this.setConstraint(new ConstraintAnnotation(cell.constraint()));
    }
}
