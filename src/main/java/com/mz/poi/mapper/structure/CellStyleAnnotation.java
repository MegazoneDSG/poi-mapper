package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.CellStyle;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Getter
@Setter
@NoArgsConstructor
public class CellStyleAnnotation {

  private FontAnnotation font = new FontAnnotation();
  private String dataFormat = "General";
  private boolean hidden = false;
  private boolean locked = false;
  private boolean quotePrefixed = false;
  private HorizontalAlignment alignment = HorizontalAlignment.GENERAL;
  private boolean wrapText = false;
  private VerticalAlignment verticalAlignment = VerticalAlignment.BOTTOM;
  private short rotation = 0;
  private short indention = 0;
  private BorderStyle borderLeft = BorderStyle.NONE;
  private BorderStyle borderRight = BorderStyle.NONE;
  private BorderStyle borderTop = BorderStyle.NONE;
  private BorderStyle borderBottom = BorderStyle.NONE;
  private IndexedColors leftBorderColor = IndexedColors.AUTOMATIC;
  private IndexedColors rightBorderColor = IndexedColors.AUTOMATIC;
  private IndexedColors topBorderColor = IndexedColors.AUTOMATIC;
  private IndexedColors bottomBorderColor = IndexedColors.AUTOMATIC;
  private FillPatternType fillPattern = FillPatternType.NO_FILL;
  private IndexedColors fillBackgroundColor = IndexedColors.AUTOMATIC;
  private IndexedColors fillForegroundColor = IndexedColors.AUTOMATIC;
  private boolean shrinkToFit = false;

  public CellStyleAnnotation(CellStyle cellStyle, CellStyleAnnotation InheritanceStyle) {
    this.font = InheritanceStyle.font;
    this.dataFormat = InheritanceStyle.dataFormat;
    this.hidden = InheritanceStyle.hidden;
    this.locked = InheritanceStyle.locked;
    this.quotePrefixed = InheritanceStyle.quotePrefixed;
    this.alignment = InheritanceStyle.alignment;
    this.wrapText = InheritanceStyle.wrapText;
    this.verticalAlignment = InheritanceStyle.verticalAlignment;
    this.rotation = InheritanceStyle.rotation;
    this.indention = InheritanceStyle.indention;
    this.borderLeft = InheritanceStyle.borderLeft;
    this.borderRight = InheritanceStyle.borderRight;
    this.borderTop = InheritanceStyle.borderTop;
    this.borderBottom = InheritanceStyle.borderBottom;
    this.leftBorderColor = InheritanceStyle.leftBorderColor;
    this.rightBorderColor = InheritanceStyle.rightBorderColor;
    this.topBorderColor = InheritanceStyle.topBorderColor;
    this.bottomBorderColor = InheritanceStyle.bottomBorderColor;
    this.fillPattern = InheritanceStyle.fillPattern;
    this.fillBackgroundColor = InheritanceStyle.fillBackgroundColor;
    this.fillForegroundColor = InheritanceStyle.fillForegroundColor;
    this.shrinkToFit = InheritanceStyle.shrinkToFit;
    this.bindAnnotation(cellStyle);
  }

  public CellStyleAnnotation(CellStyle cellStyle) {
    this.bindAnnotation(cellStyle);
  }

  private void bindAnnotation(CellStyle cellStyle) {
    if (cellStyle.font().length > 0) {
      this.font = new FontAnnotation(cellStyle.font()[0], this.font);
    }
    if (cellStyle.dataFormat().length > 0) {
      this.dataFormat = cellStyle.dataFormat()[0];
    }
    if (cellStyle.hidden().length > 0) {
      this.hidden = cellStyle.hidden()[0];
    }
    if (cellStyle.locked().length > 0) {
      this.locked = cellStyle.locked()[0];
    }
    if (cellStyle.quotePrefixed().length > 0) {
      this.quotePrefixed = cellStyle.quotePrefixed()[0];
    }
    if (cellStyle.alignment().length > 0) {
      this.alignment = cellStyle.alignment()[0];
    }
    if (cellStyle.wrapText().length > 0) {
      this.wrapText = cellStyle.wrapText()[0];
    }
    if (cellStyle.verticalAlignment().length > 0) {
      this.verticalAlignment = cellStyle.verticalAlignment()[0];
    }
    if (cellStyle.rotation().length > 0) {
      this.rotation = cellStyle.rotation()[0];
    }
    if (cellStyle.indention().length > 0) {
      this.indention = cellStyle.indention()[0];
    }
    if (cellStyle.borderLeft().length > 0) {
      this.borderLeft = cellStyle.borderLeft()[0];
    }
    if (cellStyle.borderRight().length > 0) {
      this.borderRight = cellStyle.borderRight()[0];
    }
    if (cellStyle.borderTop().length > 0) {
      this.borderTop = cellStyle.borderTop()[0];
    }
    if (cellStyle.borderBottom().length > 0) {
      this.borderBottom = cellStyle.borderBottom()[0];
    }
    if (cellStyle.leftBorderColor().length > 0) {
      this.leftBorderColor = cellStyle.leftBorderColor()[0];
    }
    if (cellStyle.rightBorderColor().length > 0) {
      this.rightBorderColor = cellStyle.rightBorderColor()[0];
    }
    if (cellStyle.topBorderColor().length > 0) {
      this.topBorderColor = cellStyle.topBorderColor()[0];
    }
    if (cellStyle.bottomBorderColor().length > 0) {
      this.bottomBorderColor = cellStyle.bottomBorderColor()[0];
    }
    if (cellStyle.fillPattern().length > 0) {
      this.fillPattern = cellStyle.fillPattern()[0];
    }
    if (cellStyle.fillBackgroundColor().length > 0) {
      this.fillBackgroundColor = cellStyle.fillBackgroundColor()[0];
    }
    if (cellStyle.fillForegroundColor().length > 0) {
      this.fillForegroundColor = cellStyle.fillForegroundColor()[0];
    }
    if (cellStyle.shrinkToFit().length > 0) {
      this.shrinkToFit = cellStyle.shrinkToFit()[0];
    }
  }

  public void applyStyle(
      org.apache.poi.ss.usermodel.CellStyle cellStyle, Font font, XSSFWorkbook workbook) {
    cellStyle.setFont(font);
//    if (isDateType) {
//      workbook.getCreationHelper().createDataFormat().getFormat()
//    } else {
//      cellStyle.setDataFormat(
//          (short) BuiltinFormats.getBuiltinFormat(this.dataFormat)
//      );
//    }
    cellStyle.setDataFormat(
        workbook.getCreationHelper().createDataFormat().getFormat(this.dataFormat)
    );
    cellStyle.setHidden(this.hidden);
    cellStyle.setLocked(this.locked);
    cellStyle.setQuotePrefixed(this.quotePrefixed);
    cellStyle.setAlignment(this.alignment);
    cellStyle.setWrapText(this.wrapText);
    cellStyle.setVerticalAlignment(this.verticalAlignment);
    cellStyle.setRotation(this.rotation);
    cellStyle.setIndention(this.indention);
    cellStyle.setBorderLeft(this.borderLeft);
    cellStyle.setBorderRight(this.borderRight);
    cellStyle.setBorderTop(this.borderTop);
    cellStyle.setBorderBottom(this.borderBottom);
    cellStyle.setLeftBorderColor(this.leftBorderColor.getIndex());
    cellStyle.setRightBorderColor(this.rightBorderColor.getIndex());
    cellStyle.setTopBorderColor(this.topBorderColor.getIndex());
    cellStyle.setBottomBorderColor(this.bottomBorderColor.getIndex());
    cellStyle.setFillPattern(this.fillPattern);
    if (!IndexedColors.AUTOMATIC.equals(this.fillBackgroundColor)) {
      cellStyle.setFillBackgroundColor(this.fillBackgroundColor.getIndex());
    }
    if (!IndexedColors.AUTOMATIC.equals(this.fillForegroundColor)) {
      cellStyle.setFillForegroundColor(this.fillForegroundColor.getIndex());
    }
    this.setShrinkToFit(this.shrinkToFit);
  }
}
