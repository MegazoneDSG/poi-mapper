package com.mz.poi.mapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface Sheet {

  String name();

  int index();

  boolean protect() default false;

  String protectKey() default "";

  ColumnWidth[] columnWidths() default {};

  int defaultRowHeightInPoints() default 20;

  int defaultColumnWidth() default 20;

  CellStyle defaultStyle() default @CellStyle;
}
