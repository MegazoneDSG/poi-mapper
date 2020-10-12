package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.structure.CellType;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class InfoRow {

  @Cell(
      column = 0,
      cellType = CellType.STRING
  )
  private String vendorTitle;

  @Cell(
      column = 1,
      cols = 2,
      cellType = CellType.STRING,
      required = true
  )
  private String vendorValue;

  @Cell(
      column = 3,
      cellType = CellType.STRING
  )
  private String toTitle;

  @Cell(
      column = 4,
      cols = 2,
      cellType = CellType.STRING,
      required = true
  )
  private String toValue;

  @Builder
  public InfoRow(String vendorTitle, String vendorValue, String toTitle, String toValue) {
    this.vendorTitle = vendorTitle;
    this.vendorValue = vendorValue;
    this.toTitle = toTitle;
    this.toValue = toValue;
  }
}
