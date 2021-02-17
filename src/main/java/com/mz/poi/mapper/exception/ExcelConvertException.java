package com.mz.poi.mapper.exception;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ExcelConvertException extends RuntimeException {

  private static final long serialVersionUID = 1L;

  public ExcelConvertException(String message) {
    super(message);
  }

  public ExcelConvertException(String message, Throwable cause) {
    super(message, cause);
  }
}
