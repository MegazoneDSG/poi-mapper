package com.mz.poi.mapper.exception;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class ReadExceptionAddress {
  private Integer sheet;
  private Integer row;
  private Integer column;

  public ReadExceptionAddress(Integer sheet) {
    this.sheet = sheet;
  }

  public ReadExceptionAddress(Integer sheet, Integer row) {
    this.sheet = sheet;
    this.row = row;
  }

  public ReadExceptionAddress(Integer sheet, Integer row, Integer column) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
  }
}
