package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.Cell;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

@Getter
@Setter
@NoArgsConstructor
public class SummaryRow {

  @Cell(
      column = 4,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String summaryTitle = "합계";

  @Cell(
      column = 5,
      cellType = CellType.FORMULA,
      ignoreParse = true
  )
  private String formula = "SUM({{items.at(0).formula}}:{{items.last.formula}})";
}
