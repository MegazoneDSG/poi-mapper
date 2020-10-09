package com.mz.poi.mapper.exception;

import java.util.Optional;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ExcelReadException extends RuntimeException {

  private static final long serialVersionUID = 1L;

  private ReadExceptionAddress readExceptionAddress = new ReadExceptionAddress();

  public ExcelReadException(String message, ReadExceptionAddress readExceptionAddress) {
    super(buildExcelAddressMessage(message, readExceptionAddress));
    this.readExceptionAddress = readExceptionAddress;
  }

  public ExcelReadException(String message, Throwable cause,
      ReadExceptionAddress readExceptionAddress) {
    super(buildExcelAddressMessage(message, readExceptionAddress), cause);
    this.readExceptionAddress = readExceptionAddress;
  }

  private static String buildExcelAddressMessage(String message,
      ReadExceptionAddress readExceptionAddress) {
    StringBuilder sb = new StringBuilder();
    sb.append(message);
    if (readExceptionAddress == null) {
      return sb.toString();
    }
    Optional.ofNullable(readExceptionAddress.getSheet())
        .ifPresent(sheet -> sb.append(String.format(" [ sheet: %s ]", sheet)));

    Optional.ofNullable(readExceptionAddress.getRow())
        .ifPresent(row -> sb.append(String.format(" [ row: %s ]", row)));

    Optional.ofNullable(readExceptionAddress.getColumn())
        .ifPresent(column -> sb.append(String.format(" [ column: %s ]", column)));

    return sb.toString();
  }
}
