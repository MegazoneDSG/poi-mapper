package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.ArrayHeader;
import com.mz.poi.mapper.expression.ArrayHeaderNameExpression;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class ArrayHeaderAnnotation extends AbstractHeaderAnnotation {

  private String simpleNameExpression;
  private String mapping;
  private ArrayHeaderNameExpression arrayHeaderNameExpression;

  public ArrayHeaderAnnotation(ArrayHeader header, CellStyleAnnotation headerDefaultStyle) {
    this.setStyle(new CellStyleAnnotation(header.style(), headerDefaultStyle));
    this.mapping = header.mapping();
    this.simpleNameExpression = header.simpleNameExpression();
  }
}
