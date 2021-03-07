package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.exception.ExcelStructureException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@NoArgsConstructor
public class SheetStructure {

  private ExcelStructure excelStructure;
  private SheetAnnotation annotation;
  private Field field;
  private String fieldName;
  private List<RowStructure> rows = new ArrayList<>();

  public void setAnnotation(SheetAnnotation annotation) {
    this.annotation = annotation;
  }

  public void setRows(List<RowStructure> rows) {
    this.rows = rows;
  }

  public RowStructure getRow(String fieldName) {
    return this.rows.stream()
        .filter(rowStructure -> rowStructure.getFieldName().equals(fieldName))
        .findFirst()
        .orElseThrow(() -> new ExcelStructureException(
            String.format("No such row of %s fieldName", fieldName)));
  }

  public boolean isAllRowsRead() {
    return this.rows.stream()
        .allMatch(RowStructure::isRead);
  }

  public RowStructure findRowByFieldName(String fieldName) {
    return this.rows.stream()
        .filter(rowStructure -> rowStructure.getFieldName().equals(fieldName))
        .findAny()
        .orElseThrow(() ->
            new ExcelStructureException(String.format("%s row not founded", fieldName))
        );
  }

  public RowStructure findRowByRowIndex(int rowNum) {
    return this.rows.stream()
        .filter(rowStructure ->
            rowStructure.getStartRowNum() <= rowNum && rowStructure.getEndRowNum() >= rowNum)
        .findAny()
        .orElseThrow(() ->
            new ExcelStructureException(String.format("row of index %s not founded", rowNum))
        );
  }

  public RowStructure nextReadRowStructure() {
    return this.rows.stream()
        .filter(rowStructure -> !rowStructure.isRead())
        .filter(rowStructure -> {
          if (rowStructure.isAfterRow()) {
            RowStructure beforeRowStructure = this.findRowByFieldName(
                rowStructure.getAnnotation().getRowAfter()
            );
            if (!beforeRowStructure.isRead()) {
              return false;
            }
            if (rowStructure.isDataRow()) {
              rowStructure.setStartRowNum(
                  beforeRowStructure.getEndRowNum() +
                      rowStructure.getAnnotation().getRowAfterOffset() + 1
              );
            } else {
              rowStructure.setStartRowNum(
                  beforeRowStructure.getEndRowNum() +
                      rowStructure.getAnnotation().getRowAfterOffset() + 1
              );
              rowStructure.setEndRowNum(
                  rowStructure.getStartRowNum()
              );
            }
          } else {
            if (rowStructure.isDataRow()) {
              rowStructure.setStartRowNum(
                  rowStructure.getAnnotation().getRow()
              );
            } else {
              rowStructure.setStartRowNum(
                  rowStructure.getAnnotation().getRow()
              );
              rowStructure.setEndRowNum(
                  rowStructure.getStartRowNum()
              );
            }
          }
          return true;
        })
        .findFirst()
        .orElseThrow(() ->
            new ExcelStructureException("not found nextRowStructure")
        );
  }

  @Builder
  public SheetStructure(ExcelStructure excelStructure,
      SheetAnnotation annotation, Field field, String fieldName) {
    this.excelStructure = excelStructure;
    this.annotation = annotation;
    this.field = field;
    this.fieldName = fieldName;
  }
}
