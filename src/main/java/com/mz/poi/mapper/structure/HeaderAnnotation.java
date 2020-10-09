package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Header;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class HeaderAnnotation {

  private int column;
  private String name;
  private int cols;
  CellStyleAnnotation style;

  public HeaderAnnotation(Header header, CellStyleAnnotation headerDefaultStyle) {
    this.column = header.column();
    this.name = header.name();
    this.cols = header.cols();
    this.style = new CellStyleAnnotation(header.style(), headerDefaultStyle);
  }
}
