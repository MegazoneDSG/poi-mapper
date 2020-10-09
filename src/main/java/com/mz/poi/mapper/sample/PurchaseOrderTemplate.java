package com.mz.poi.mapper.sample;

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
    )
)
public class PurchaseOrderTemplate {

  @Sheet(
      name = "Order",
      index = 0,
      columnWidths = {
          @ColumnWidth(column = 0, width = 25)
      },
      defaultColumnWidth = 20,
      defaultRowHeightInPoints = 20
  )
  private OrderSheet sheet = new OrderSheet();

  @Builder
  public PurchaseOrderTemplate(OrderSheet sheet) {
    this.sheet = sheet;
  }
}
