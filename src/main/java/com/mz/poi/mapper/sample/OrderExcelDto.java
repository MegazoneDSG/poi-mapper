package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.ColumnWidth;
import com.mz.poi.mapper.annotation.Excel;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Sheet;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@Excel(
    defaultStyle = @CellStyle(
        font = @Font(fontName = "나눔고딕")
    )
)
public class OrderExcelDto {

  @Sheet(
      name = "PMS 발주서",
      index = 0,
      columnWidths = {
          @ColumnWidth(column = 0, width = 15),
          @ColumnWidth(column = 11, width = 30)
      },
      protect = false,
      protectKey = "1234",
      defaultColumnWidth = 20,
      defaultRowHeightInPoints = 20
  )
  private OrderSheet sheet = new OrderSheet();
}
