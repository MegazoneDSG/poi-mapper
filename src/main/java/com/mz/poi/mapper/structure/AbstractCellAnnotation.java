package com.mz.poi.mapper.structure;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class AbstractCellAnnotation {

  private int column;
  private int cols;
  private String columnAfter;
  private int columnAfterOffset;
  private CellType cellType;
  private boolean ignoreParse;
  private boolean required;
  private CellStyleAnnotation style;

  public int getColumnSize() {
    if (this instanceof ArrayCellAnnotation) {
      return ((ArrayCellAnnotation) this).getSize();
    } else {
      return 1;
    }
  }
}
