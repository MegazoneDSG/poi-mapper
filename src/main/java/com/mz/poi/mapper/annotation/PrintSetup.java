package com.mz.poi.mapper.annotation;

import static org.apache.poi.ss.usermodel.PrintSetup.LETTER_PAPERSIZE;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.ANNOTATION_TYPE})
public @interface PrintSetup {


	/**
	 * Returns the paper size.
	 *
	 * @return paper size
	 */
	short paperSize() default LETTER_PAPERSIZE;

	/**
	 * Returns the scale.
	 *
	 * @return scale
	 */
	short scale() default 100;

	/**
	 * Returns the page start.
	 *
	 * @return page start
	 */
	short pageStart() default 1;

	/**
	 * Returns the number of pages wide to fit sheet in.
	 *
	 * @return number of pages wide to fit sheet in
	 */
	short fitWidth() default 1;

	/**
	 * Returns the number of pages high to fit the sheet in.
	 *
	 * @return number of pages high to fit the sheet in
	 */
	short fitHeight() default 1;

	/**
	 * Returns the left to right print order.
	 *
	 * @return left to right print order
	 */
	boolean leftToRight() default false;

	/**
	 * Returns the landscape mode.
	 *
	 * @return landscape mode
	 */
	boolean landscape() default false;

	/**
	 * Returns the valid settings.
	 *
	 * @return valid settings
	 */
	boolean validSettings() default true;

	/**
	 * Returns the black and white setting.
	 *
	 * @return black and white setting
	 */
	boolean noColor() default false;

	/**
	 * Returns the draft mode.
	 *
	 * @return draft mode
	 */
	boolean draft() default false;

	/**
	 * Returns the print notes.
	 *
	 * @return print notes
	 */
	boolean notes() default false;

	/**
	 * Returns the no orientation.
	 *
	 * @return no orientation
	 */
	boolean noOrientation() default true;

	/**
	 * Returns the use page numbers.
	 *
	 * @return use page numbers
	 */
	boolean usePage() default false;

	/**
	 * Returns the horizontal resolution.
	 *
	 * @return horizontal resolution
	 */
	short hResolution() default 600;

	/**
	 * Returns the vertical resolution.
	 *
	 * @return vertical resolution
	 */
	short vResolution() default 600;

	/**
	 * Returns the header margin.
	 *
	 * @return header margin
	 */
	double headerMargin() default 0.3;

	/**
	 * Returns the footer margin.
	 *
	 * @return footer margin
	 */
	double footerMargin() default 0.3;

	/**
	 * Returns the number of copies.
	 *
	 * @return number of copies
	 */
	short copies() default 1;

}
