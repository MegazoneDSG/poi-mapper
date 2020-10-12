package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.structure.CellType;
import java.time.LocalDate;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class ShipRow {

  @Cell(
      column = 0,
      cellType = CellType.STRING
  )
  private String requester;

  @Cell(
      column = 1,
      cellType = CellType.STRING
  )
  private String via;

  @Cell(
      column = 2,
      cellType = CellType.STRING
  )
  private String fob;

  @Cell(
      column = 3,
      cols = 2,
      cellType = CellType.STRING
  )
  private String terms;

  @Cell(
      column = 5,
      cellType = CellType.DATE,
      style = @CellStyle(dataFormat = "yyyy-MM-dd"),
      required = true
  )
  private LocalDate deliveryDate;

  @Builder
  public ShipRow(String requester, String via, String fob, String terms,
      LocalDate deliveryDate) {
    this.requester = requester;
    this.via = via;
    this.fob = fob;
    this.terms = terms;
    this.deliveryDate = deliveryDate;
  }
}
