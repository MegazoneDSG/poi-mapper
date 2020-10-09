package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Sheet;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class SheetAnnotation {

  private String name;
  private int index;
  private boolean protect;
  private String protectKey;
  private List<ColumnWidthAnnotation> columnWidths = new ArrayList<>();
  private int defaultRowHeightInPoints;
  private int defaultColumnWidth;
  private CellStyleAnnotation defaultStyle;

  public SheetAnnotation(Sheet sheet, CellStyleAnnotation rootStyle) {
    this.name = sheet.name();
    this.index = sheet.index();
    this.protect = sheet.protect();
    this.protectKey = sheet.protectKey();
    Arrays.asList(sheet.columnWidths())
        .forEach(columnWidth -> this.columnWidths.add(
            new ColumnWidthAnnotation(columnWidth)
        ));
    this.defaultRowHeightInPoints = sheet.defaultRowHeightInPoints();
    this.defaultColumnWidth = sheet.defaultColumnWidth();
    this.defaultStyle = new CellStyleAnnotation(sheet.defaultStyle(), rootStyle);
  }
}
