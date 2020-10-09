package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.Cell;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

@Getter
@Setter
@NoArgsConstructor
public class Info3 {

  @Cell(
      column = 0,
      cols = 2,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String sManagerTitle = "담당자";

  @Cell(
      column = 2,
      cols = 4,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String sManagerValue;

  @Cell(
      column = 6,
      cols = 2,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String managerTitle = "담당MD";

  @Cell(
      column = 8,
      cols = 4,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String managerValue;
}
