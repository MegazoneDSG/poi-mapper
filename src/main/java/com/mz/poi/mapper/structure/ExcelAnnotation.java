package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Excel;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class ExcelAnnotation {

  private CellStyleAnnotation defaultStyle;

  public ExcelAnnotation(Excel excel) {
    this.defaultStyle = new CellStyleAnnotation(excel.defaultStyle());
  }
}
