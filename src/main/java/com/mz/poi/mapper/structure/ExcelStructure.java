package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Excel;
import com.mz.poi.mapper.annotation.Row;
import com.mz.poi.mapper.annotation.Sheet;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class ExcelStructure {

  private Class<?> dtoType;
  private ExcelAnnotation annotation;
  private List<SheetStructure> sheets = new ArrayList<>();


  @Getter
  @Setter
  @NoArgsConstructor
  public static class SheetStructure {

    private SheetAnnotation annotation;
    private Field field;
    private String fieldName;
    private List<RowStructure> rows = new ArrayList<>();

    public boolean isAllRowsGenerated() {
      return this.rows.stream()
          .allMatch(rowStructure -> rowStructure.generated);
    }

    public RowStructure findRowByFieldName(String fieldName) {
      return this.rows.stream()
          .filter(rowStructure -> rowStructure.fieldName.equals(fieldName))
          .findAny()
          .orElseThrow(() ->
              new IllegalArgumentException(String.format("%s row not founded", fieldName))
          );
    }

    public RowStructure findRowByRowIndex(int rowNum) {
      return this.rows.stream()
          .filter(rowStructure -> rowStructure.generated)
          .filter(rowStructure ->
              rowStructure.startRowNum <= rowNum && rowStructure.endRowNum >= rowNum)
          .findAny()
          .orElseThrow(() ->
              new IllegalArgumentException(String.format("row of index %s not founded", rowNum))
          );
    }

    public RowStructure nextRowStructure() {
      return this.rows.stream()
          .filter(rowStructure -> !rowStructure.generated)
          .filter(rowStructure -> {
            if (!rowStructure.isAfterRow()) {
              return true;
            } else {
              RowStructure beforeRowStructure = this.findRowByFieldName(
                  rowStructure.getAnnotation().getRowAfter()
              );
              if (beforeRowStructure.isGenerated()) {
                rowStructure.setStartRowNum(
                    beforeRowStructure.endRowNum +
                        rowStructure.getAnnotation().getRowAfterOffset() + 1
                );
                return true;
              }
              return false;
            }
          })
          .findFirst()
          .orElseThrow(() ->
              new IllegalArgumentException("not found nextRowStructure")
          );
    }

    @Builder
    public SheetStructure(
        SheetAnnotation annotation, Field field, String fieldName) {
      this.annotation = annotation;
      this.field = field;
      this.fieldName = fieldName;
    }
  }

  @Getter
  @Setter
  @NoArgsConstructor
  public static class RowStructure {

    private AbstractRowAnnotation annotation;
    private Field sheetField;
    private Field field;
    private String fieldName;
    private List<CellStructure> cells = new ArrayList<>();

    public boolean isAfterRow() {
      String rowAfter = this.annotation.getRowAfter();
      return rowAfter != null && rowAfter.length() > 0;
    }

    public boolean isDataRow() {
      return this.annotation instanceof DataRowsAnnotation;
    }

    private boolean generated;
    private int startRowNum;
    private int endRowNum;

    @Builder
    public RowStructure(
        AbstractRowAnnotation annotation, Field sheetField, Field field, String fieldName) {
      this.sheetField = sheetField;
      this.field = field;
      this.annotation = annotation;
      this.fieldName = fieldName;
    }
  }

  @Getter
  @Setter
  @NoArgsConstructor
  public static class CellStructure {

    private CellAnnotation annotation;
    private Field sheetField;
    private Field rowField;
    private Field field;
    private String fieldName;

    @Builder
    public CellStructure(
        CellAnnotation annotation, Field sheetField, Field rowField, Field field,
        String fieldName) {
      this.sheetField = sheetField;
      this.rowField = rowField;
      this.annotation = annotation;
      this.field = field;
      this.fieldName = fieldName;
    }
  }

  public ExcelStructure build(Class<?> dtoType) {

    this.dtoType = dtoType;
    Excel excel = dtoType.getAnnotation(Excel.class);
    if (excel == null) {
      throw new IllegalArgumentException("not found excel annotation");
    }
    this.setAnnotation(new ExcelAnnotation(excel));

    this.sheets = Arrays.stream(dtoType.getDeclaredFields())
        .filter(field -> {
          field.setAccessible(true);
          Sheet sheet = field.getAnnotation(Sheet.class);
          return sheet != null;
        })
        .map(field -> {
          SheetStructure sheetStructure = SheetStructure.builder()
              .annotation(
                  new SheetAnnotation(
                      field.getAnnotation(Sheet.class),
                      this.annotation.getDefaultStyle()
                  )
              )
              .field(field)
              .fieldName(field.getName())
              .build();
          this.sheets.add(sheetStructure);
          return sheetStructure;
        })
        .peek(sheetStructure ->
            Arrays.stream(sheetStructure.field.getType().getDeclaredFields())
                .filter(field -> {
                  field.setAccessible(true);
                  Row row = field.getAnnotation(Row.class);
                  DataRows dataRows = field.getAnnotation(DataRows.class);
                  return row != null ||
                      (dataRows != null &&
                          Collection.class.isAssignableFrom(field.getType()));
                })
                .map(field -> {
                  RowStructure rowStructure = RowStructure.builder()
                      .sheetField(sheetStructure.field)
                      .field(field)
                      .fieldName(field.getName())
                      .build();
                  boolean isRow = field.getAnnotation(Row.class) != null;
                  if (isRow) {
                    rowStructure.setAnnotation(
                        new RowAnnotation(
                            field.getAnnotation(Row.class),
                            sheetStructure.annotation.getDefaultStyle()
                        )
                    );
                  } else {
                    rowStructure.setAnnotation(
                        new DataRowsAnnotation(
                            field.getAnnotation(DataRows.class),
                            sheetStructure.annotation.getDefaultStyle()
                        )
                    );
                  }
                  if (!rowStructure.isAfterRow()) {
                    rowStructure.setStartRowNum(rowStructure.getAnnotation().getRow());
                    rowStructure.setEndRowNum(rowStructure.getAnnotation().getRow());
                  }
                  sheetStructure.rows.add(rowStructure);
                  return rowStructure;
                })
                .forEach(rowStructure -> {
                  Field[] fields;
                  boolean isRow = rowStructure.annotation instanceof RowAnnotation;
                  if (isRow) {
                    fields = rowStructure.field.getType().getDeclaredFields();
                  } else {
                    ParameterizedType genericType =
                        (ParameterizedType) rowStructure.field.getGenericType();
                    Class<?> dataRowClass = (Class<?>) genericType.getActualTypeArguments()[0];
                    fields = dataRowClass.getDeclaredFields();
                  }

                  Arrays.stream(fields)
                      .filter(field -> {
                        field.setAccessible(true);
                        Cell cell = field.getAnnotation(Cell.class);
                        return cell != null;
                      })
                      .forEach(field -> {
                        CellStructure cellStructure = CellStructure.builder()
                            .annotation(
                                new CellAnnotation(
                                    field.getAnnotation(Cell.class),
                                    isRow ?
                                        ((RowAnnotation) rowStructure.annotation)
                                            .getDefaultStyle() :
                                        ((DataRowsAnnotation) rowStructure.annotation)
                                            .getDataStyle())
                            )
                            .sheetField(rowStructure.sheetField)
                            .rowField(rowStructure.field)
                            .field(field)
                            .fieldName(field.getName())
                            .build();
                        rowStructure.cells.add(cellStructure);
                      });
                }))
        .collect(Collectors.toList());
    return this;
  }
}
