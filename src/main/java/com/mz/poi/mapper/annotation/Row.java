package com.mz.poi.mapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface Row {

  int row() default 0;

  int[] heightInPoints() default {};

  String rowAfter() default "";

  int rowAfterOffset() default 0;

  CellStyle defaultStyle() default @CellStyle;
}
