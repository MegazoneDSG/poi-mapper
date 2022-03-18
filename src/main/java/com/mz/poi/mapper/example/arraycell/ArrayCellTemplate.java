package com.mz.poi.mapper.example.arraycell;

import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.ColumnWidth;
import com.mz.poi.mapper.annotation.Excel;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Sheet;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@Excel(
    defaultStyle = @CellStyle(
        font = @Font(fontName = "Arial")
    ),
    dateFormatZoneId = "Asia/Seoul"
)
public class ArrayCellTemplate {

  private ArrayCellSheet sheet = new ArrayCellSheet();

  @Builder
  public ArrayCellTemplate(ArrayCellSheet sheet) {
    this.sheet = sheet;
  }
}
