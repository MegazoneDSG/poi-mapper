package com.mz.poi.mapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.ANNOTATION_TYPE})
public @interface Constraint {

	boolean suppressDropDownArrow() default true;

	boolean showErrorBox() default true;

	String errorBoxTitle() default "ERROR";

	String errorBoxText() default "ERROR";

	ErrorStyle errorStyle() default ErrorStyle.STOP;

	String[] constraints() default {};
}
