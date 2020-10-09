package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.Cell;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

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
      cellType = CellType.STRING,
      required = true
  )
  private String deliveryDate;

  @Builder
  public ShipRow(String requester, String via, String fob, String terms,
      String deliveryDate) {
    this.requester = requester;
    this.via = via;
    this.fob = fob;
    this.terms = terms;
    this.deliveryDate = deliveryDate;
  }
}
