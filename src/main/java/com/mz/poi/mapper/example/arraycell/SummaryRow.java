package com.mz.poi.mapper.example.arraycell;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.structure.CellType;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter
@Setter
@NoArgsConstructor
public class SummaryRow {

  @Cell(
      column = 6,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String title = "SUBTOTAL";

  @Cell(
      columnAfter = "title",
      columnAfterOffset = 1,
      cellType = CellType.FORMULA,
      style = @CellStyle(
          fillForegroundColor = IndexedColors.AQUA,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      ),
      ignoreParse = true
  )
  private String formula = "SUM({{itemTable[0].total}}:{{itemTable[last].total}})";
}
