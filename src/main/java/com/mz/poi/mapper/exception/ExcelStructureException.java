package com.mz.poi.mapper.exception;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ExcelStructureException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public ExcelStructureException(String message) {
        super(message);
    }

    public ExcelStructureException(String message, Throwable cause) {
        super(message, cause);
    }
}
