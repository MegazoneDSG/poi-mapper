package com.mz.poi.mapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.ANNOTATION_TYPE})
public @interface CellStyle {

  /**
   * if empty, default is @Font
   *
   * @return Font
   */
  Font[] font() default {};

  /**
   * get data format. Built in formats are defined at {@link BuiltinFormats}.
   * <p>
   * if empty, default is "General"
   *
   * @see DataFormat
   */
  String[] dataFormat() default {};

  /**
   * get whether the cell's using this style are to be hidden
   * <p>
   * if empty, default is false
   *
   * @return hidden - whether the cell using this style should be hidden
   */
  boolean[] hidden() default {};

  /**
   * get whether the cell's using this style are to be locked
   * <p>
   * if empty, default is false
   *
   * @return hidden - whether the cell using this style should be locked
   */
  boolean[] locked() default {};

  /**
   * Is "Quote Prefix" or "123 Prefix" enabled for the cell? Having this on is somewhat (but not
   * completely, see {@link IgnoredErrorType}) like prefixing the cell value with a ' in Excel
   * <p>
   * if empty, default is false
   */
  boolean[] quotePrefixed() default {};

  /**
   * get the type of horizontal alignment for the cell
   * <p>
   * if empty, default is HorizontalAlignment.GENERAL
   *
   * @return align - the type of alignment
   */
  HorizontalAlignment[] alignment() default {};

  /**
   * get whether the text should be wrapped
   * <p>
   * if empty, default is false
   *
   * @return wrap text or not
   */
  boolean[] wrapText() default {};

  /**
   * get the type of vertical alignment for the cell
   * <p>
   * if empty, default is VerticalAlignment.BOTTOM
   *
   * @return align the type of alignment
   */
  VerticalAlignment[] verticalAlignment() default {};

  /**
   * get the degree of rotation for the text in the cell.
   * <p>
   * Note: HSSF uses values from -90 to 90 degrees, whereas XSSF uses values from 0 to 180 degrees.
   * The implementations of this method will map between these two value-ranges value-range as used
   * by the type of Excel file-format that this CellStyle is applied to.
   * <p>
   * if empty, default is 0
   *
   * @return rotation degrees (see note above)
   */
  short[] rotation() default {};

  /**
   * get the number of spaces to indent the text in the cell.
   * <p>
   * if empty, default is 0
   *
   * @return indent - number of spaces
   */
  short[] indention() default {};

  /**
   * get the type of border to use for the left border of the cell
   * <p>
   * if empty, default is BorderStyle.NONE
   *
   * @return border type
   * @since POI 4.0.0
   */
  BorderStyle[] borderLeft() default {};

  /**
   * get the type of border to use for the right border of the cell
   * <p>
   * if empty, default is BorderStyle.NONE
   *
   * @return border type
   * @since POI 4.0.0
   */
  BorderStyle[] borderRight() default {};

  /**
   * get the type of border to use for the top border of the cell
   * <p>
   * if empty, default is BorderStyle.NONE
   *
   * @return border type
   * @since POI 4.0.0
   */
  BorderStyle[] borderTop() default {};

  /**
   * get the type of border to use for the bottom border of the cell
   * <p>
   * if empty, default is BorderStyle.NONE
   *
   * @return border type
   * @since POI 4.0.0
   */
  BorderStyle[] borderBottom() default {};

  /**
   * get the color to use for the left border
   * <p>
   * if empty, default is IndexedColors.AUTOMATIC
   */
  IndexedColors[] leftBorderColor() default {};

  /**
   * get the color to use for the left border
   * <p>
   * if empty, default is IndexedColors.AUTOMATIC
   *
   * @return the index of the color definition
   */
  IndexedColors[] rightBorderColor() default {};

  /**
   * get the color to use for the top border
   * <p>
   * if empty, default is IndexedColors.AUTOMATIC
   *
   * @return the index of the color definition
   */
  IndexedColors[] topBorderColor() default {};

  /**
   * get the color to use for the left border
   * <p>
   * if empty, default is IndexedColors.AUTOMATIC
   *
   * @return the index of the color definition
   */
  IndexedColors[] bottomBorderColor() default {};

  /**
   * Get the fill pattern
   * <p>
   * if empty, default is FillPatternType.NO_FILL
   *
   * @return the fill pattern, default value is {@link FillPatternType#NO_FILL}
   * @since POI 4.0.0
   */
  FillPatternType[] fillPattern() default {};

  /**
   * get the background fill color, if the fill is defined with an indexed color.
   * <p>
   * if empty, default is IndexedColors.AUTOMATIC
   *
   * @return fill color index, or 0 if not indexed (XSSF only)
   */
  IndexedColors[] fillBackgroundColor() default {};

  /**
   * get the foreground fill color, if the fill is defined with an indexed color.
   * <p>
   * if empty, default is IndexedColors.AUTOMATIC
   *
   * @return fill color, or 0 if not indexed (XSSF only)
   */
  IndexedColors[] fillForegroundColor() default {};

  /**
   * Should the Cell be auto-sized by Excel to shrink it to fit if this text is too long?
   * <p>
   * if empty, default is false
   */
  boolean[] shrinkToFit() default {};
}
