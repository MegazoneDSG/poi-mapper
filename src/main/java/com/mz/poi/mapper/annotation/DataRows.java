package com.mz.poi.mapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface DataRows {

  int row() default 0;

  String rowAfter() default "";

  int rowAfterOffset() default 0;

  Match match() default Match.ALL;

  CellStyle dataStyle() default @CellStyle;

  int[] dataHeightInPoints() default {};

  CellStyle headerStyle() default @CellStyle;

  int[] headerHeightInPoints() default {};

  boolean hideHeader() default false;

  Header[] headers() default {};

  ArrayHeader[] arrayHeaders() default {};
}
