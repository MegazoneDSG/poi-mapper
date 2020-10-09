package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Cell;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

@Getter
@Setter
@NoArgsConstructor
public class CellAnnotation {

  private int column;
  private int cols;
  private CellType cellType;
  private boolean ignoreParse;
  private boolean required;
  private CellStyleAnnotation style;

  public CellAnnotation(Cell cell, CellStyleAnnotation rowStyle) {
    this.column = cell.column();
    this.cols = cell.cols();
    this.cellType = cell.cellType();
    this.ignoreParse = cell.ignoreParse();
    this.required = cell.required();
    this.style = new CellStyleAnnotation(cell.style(), rowStyle);
  }
}
