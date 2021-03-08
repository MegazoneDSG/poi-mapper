package com.mz.poi.mapper.example.basic;

import com.mz.poi.mapper.annotation.Cell;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import com.mz.poi.mapper.structure.CellType;

@Getter
@Setter
@NoArgsConstructor
public class TitleRow {

  @Cell(
      column = 0,
      cols = 6,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String title = "PURCHASE ORDER";
}
