package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.Cell;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

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
  private String supplierNameTile = "거래처 정보";

  @Cell(
      column = 6,
      cols = 6,
      cellType = CellType.STRING,
      ignoreParse = true
  )
  private String ordererTitle = "발주처 정보";
}
