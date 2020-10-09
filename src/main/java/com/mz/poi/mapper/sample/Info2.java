package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.Cell;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

@Getter
@Setter
@NoArgsConstructor
public class Info2 {

  @Cell(
      column = 0,
      cols = 2,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String supplierNameTile = "거래처명";

  @Cell(
      column = 2,
      cols = 4,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String supplierNameValue;

  @Cell(
      column = 6,
      cols = 2,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String ordererTitle = "발주처명";

  @Cell(
      column = 8,
      cols = 4,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String ordererValue;
}
