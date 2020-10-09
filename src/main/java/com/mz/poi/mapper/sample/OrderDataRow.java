package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

@Getter
@Setter
@NoArgsConstructor
public class OrderDataRow {

  @Cell(
      column = 0,
      cellType = CellType.STRING
  )
  private String productName;

  @Cell(
      column = 2,
      cellType = CellType.NUMERIC,
      style = @CellStyle(dataFormat = "#,##0"),
      required = true
  )
  private Long productNumber;

  @Cell(
      column = 3,
      cellType = CellType.NUMERIC,
      style = @CellStyle(dataFormat = "#,##0"),
      required = true
  )
  private Long skuId;

  @Cell(
      column = 4,
      cols = 2,
      cellType = CellType.FORMULA,
      ignoreParse = true
  )
  private String formula = "SUM({{this.productNumber}}:{{this.skuId}})";

  @Builder
  public OrderDataRow(String productName, Long productNumber, Long skuId) {
    this.productName = productName;
    this.productNumber = productNumber;
    this.skuId = skuId;
  }
}
