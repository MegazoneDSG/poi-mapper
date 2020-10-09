package com.mz.poi.mapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import org.apache.poi.ss.usermodel.IndexedColors;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.ANNOTATION_TYPE})
public @interface Font {

  /**
   * get the name for the font (i.e. Arial)
   * <p>
   * if empty, default is "Arial"
   *
   * @return String representing the name of the font to use
   */
  String[] fontName() default {};

  /**
   * get the font height
   * <p>
   * if empty, default is 10
   *
   * @return height in the familiar unit of measure - points
   */
  short[] fontHeightInPoints() default {};

  /**
   * get whether to use italics or not
   * <p>
   * if empty, default is false
   *
   * @return italics or not
   */
  boolean[] italic() default {};

  /**
   * get whether to use a strikeout horizontal line through the text or not
   * <p>
   * if empty, default is false
   *
   * @return strikeout or not
   */
  boolean[] strikeout() default {};

  /**
   * get the color for the font
   * <p>
   * if empty, default is IndexedColors.AUTOMATIC
   *
   * @return color to use
   * @see org.apache.poi.hssf.usermodel.HSSFPalette#getColor(short)
   */
  IndexedColors[] color() default {};

  /**
   * get normal,super or subscript. get the color for the font
   * <p>
   * if empty, default is SS_NONE
   *
   * @return offset type to use (SS_NONE,SS_SUPER,SS_SUB)
   */
  short[] typeOffset() default {};

  /**
   * get type of text underlining to use
   * <p>
   * if empty, default is U_NONE
   *
   * @return underlining type (U_NONE,U_SINGLE,U_DOUBLE,U_SINGLE_ACCOUNTING,U_DOUBLE_ACCOUNTING)
   */

  byte[] underline() default {};

  /**
   * get character-set to use.
   * <p>
   * if empty, default is ANSI_CHARSET
   *
   * @return character-set (ANSI_CHARSET,DEFAULT_CHARSET,SYMBOL_CHARSET)
   */
  int[] charSet() default {};

  /**
   * if empty, default is false
   *
   * @return bold
   */
  boolean[] bold() default {};
}
