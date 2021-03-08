package com.mz.poi.mapper.structure;


import com.mz.poi.mapper.exception.ExcelGenerateException;
import com.mz.poi.mapper.exception.ExcelStructureException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@NoArgsConstructor
public class RowStructure {

  private SheetStructure sheetStructure;
  private AbstractRowAnnotation annotation;
  private Field sheetField;
  private Field field;
  private String fieldName;
  private List<CellStructure> cells = new ArrayList<>();

  public void setAnnotation(AbstractRowAnnotation annotation) {
    this.annotation = annotation;
  }

  public void setCells(List<CellStructure> cells) {
    this.cells = cells;
  }

  public CellStructure getCell(String fieldName) {
    return this.cells.stream()
        .filter(cellStructure -> cellStructure.getFieldName().equals(fieldName))
        .findFirst()
        .orElseThrow(() -> new ExcelStructureException(
            String.format("No such cell of %s fieldName", fieldName)));
  }

  public boolean isAfterRow() {
    String rowAfter = this.annotation.getRowAfter();
    return rowAfter != null && rowAfter.length() > 0;
  }

  public boolean isDataRow() {
    return this.annotation instanceof DataRowsAnnotation;
  }

  public boolean isDataRowAndHideHeader() {
    if (this.annotation instanceof DataRowsAnnotation) {
      return ((DataRowsAnnotation) this.annotation).isHideHeader();
    }
    return false;
  }

  private boolean read;
  private boolean calculated;
  private int startRowNum;
  private int endRowNum;

  public void setRead(boolean read) {
    this.read = read;
  }

  public void setCalculated(boolean calculated) {
    this.calculated = calculated;
  }

  public void setStartRowNum(int startRowNum) {
    this.startRowNum = startRowNum;
  }

  public void setEndRowNum(int endRowNum) {
    this.endRowNum = endRowNum;
  }

  public Object findRowData(Object excelDto) {
    try {
      Field sheetField = this.getSheetField();
      Object sheetObj = sheetField.get(excelDto);
      if (sheetObj == null) {
        return null;
      }
      Field rowField = this.getField();
      return rowField.get(sheetObj);
    } catch (IllegalAccessException e) {
      throw new ExcelGenerateException(
          String.format("can not find data row collection, %s", this.getFieldName()), e);
    }
  }

  @SuppressWarnings("unchecked")
  public <T> Collection<T> findRowDataCollection(Object excelDto) {
    try {
      Field sheetField = this.getSheetField();
      Object sheetObj = sheetField.get(excelDto);
      if (sheetObj == null) {
        return new ArrayList<>();
      }
      Field rowField = this.getField();
      return (Collection<T>) rowField.get(sheetObj);
    } catch (IllegalAccessException e) {
      throw new ExcelGenerateException(
          String.format("can not find data row collection, %s", this.getFieldName()), e);
    }
  }

  public int findRowDataCollectionSize(Object excelDto) {
    return findRowDataCollection(excelDto).size();
  }

  public CellStructure findCellByFieldName(String fieldName) {
    return this.cells.stream()
        .filter(cellStructure -> cellStructure.getFieldName().equals(fieldName))
        .findAny()
        .orElseThrow(() ->
            new ExcelStructureException(String.format("%s cell not founded", fieldName))
        );
  }

  public void calculateRowNum(Object excelDto) {
    if (this.isAfterRow()) {
      RowStructure beforeRowStructure = this.sheetStructure.findRowByFieldName(
          this.annotation.getRowAfter()
      );
      if (!beforeRowStructure.calculated) {
        beforeRowStructure.calculateRowNum(excelDto);
      }

      if (this.isDataRow()) {
        this.startRowNum =
            beforeRowStructure.getEndRowNum() + this.getAnnotation().getRowAfterOffset() + 1;
        this.endRowNum =
            this.getStartRowNum() + this.findRowDataCollectionSize(excelDto) +
                (this.isDataRowAndHideHeader() ? -1 : 0);
      } else {
        this.startRowNum =
            beforeRowStructure.getEndRowNum() +
                this.getAnnotation().getRowAfterOffset() + 1;
        this.endRowNum = this.getStartRowNum();
      }

    } else {
      if (this.isDataRow()) {
        this.startRowNum = this.getAnnotation().getRow();
        this.endRowNum =
            this.getStartRowNum() + this.findRowDataCollectionSize(excelDto)
                + (this.isDataRowAndHideHeader() ? -1 : 0);
      } else {
        this.startRowNum = this.getAnnotation().getRow();
        this.endRowNum = this.getStartRowNum();
      }
    }
    this.calculated = true;
  }

  @Builder
  public RowStructure(
      SheetStructure sheetStructure,
      AbstractRowAnnotation annotation, Field sheetField, Field field, String fieldName) {
    this.sheetStructure = sheetStructure;
    this.sheetField = sheetField;
    this.field = field;
    this.annotation = annotation;
    this.fieldName = fieldName;
  }
}
