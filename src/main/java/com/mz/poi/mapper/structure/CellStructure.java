package com.mz.poi.mapper.structure;

import java.lang.reflect.Field;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@NoArgsConstructor
public class CellStructure {

  private RowStructure rowStructure;
  private CellAnnotation annotation;
  private Field sheetField;
  private Field rowField;
  private Field field;
  private String fieldName;

  public void setAnnotation(CellAnnotation annotation) {
    this.annotation = annotation;
  }

  @Builder
  public CellStructure(
      RowStructure rowStructure,
      CellAnnotation annotation, Field sheetField, Field rowField, Field field,
      String fieldName) {
    this.rowStructure = rowStructure;
    this.sheetField = sheetField;
    this.rowField = rowField;
    this.annotation = annotation;
    this.field = field;
    this.fieldName = fieldName;
  }
}
