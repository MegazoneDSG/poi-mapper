package com.mz.poi.mapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import org.apache.poi.ss.usermodel.CellType;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface Cell {

  int column() default 0;

  int cols() default 1;

  CellType cellType();

  boolean ignoreParse() default false;

  boolean required() default false;

  CellStyle style() default @CellStyle;
}
