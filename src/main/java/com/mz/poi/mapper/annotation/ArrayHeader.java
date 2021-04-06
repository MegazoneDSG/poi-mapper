package com.mz.poi.mapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ArrayHeader {

  CellStyle style() default @CellStyle;

  String mapping() default "";

  String simpleNameExpression() default "{{index}}";
}
