package com.mz.poi.mapper.exception;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ExcelGenerateException extends RuntimeException {

  private static final long serialVersionUID = 1L;

  public ExcelGenerateException(String message) {
    super(message);
  }

  public ExcelGenerateException(String message, Throwable cause) {
    super(message, cause);
  }
}
