package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Excel;
import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.annotation.Row;
import com.mz.poi.mapper.annotation.Sheet;
import com.mz.poi.mapper.exception.ExcelStructureException;
import com.mz.poi.mapper.helper.InheritedFieldHelper;
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

@Getter
@NoArgsConstructor
public class ExcelStructure {

  private Class<?> dtoType;
  private ExcelAnnotation annotation;
  private List<SheetStructure> sheets = new ArrayList<>();

  public void setAnnotation(ExcelAnnotation annotation) {
    this.annotation = annotation;
  }

  public SheetStructure getSheet(String fieldName) {
    return this.sheets.stream()
        .filter(sheetStructure -> sheetStructure.getFieldName().equals(fieldName))
        .findFirst()
        .orElseThrow(() -> new ExcelStructureException(
            String.format("No such sheet of %s fieldName", fieldName)));
  }

  public void resetRowGeneratedStatus() {
    this.sheets.forEach(sheetStructure -> {
      sheetStructure.getRows().forEach(rowStructure -> {
        rowStructure.generated = false;
        if (!rowStructure.isAfterRow()) {
          rowStructure.startRowNum = rowStructure.getAnnotation().getRow();
          rowStructure.endRowNum = rowStructure.getAnnotation().getRow();
        } else {
          rowStructure.startRowNum = 0;
          rowStructure.endRowNum = 0;
        }
      });
    });
  }

  @Getter
  @NoArgsConstructor
  public static class SheetStructure {

    private SheetAnnotation annotation;
    private Field field;
    private String fieldName;
    private List<RowStructure> rows = new ArrayList<>();

    public void setAnnotation(SheetAnnotation annotation) {
      this.annotation = annotation;
    }

    public RowStructure getRow(String fieldName) {
      return this.rows.stream()
          .filter(rowStructure -> rowStructure.getFieldName().equals(fieldName))
          .findFirst()
          .orElseThrow(() -> new ExcelStructureException(
              String.format("No such row of %s fieldName", fieldName)));
    }

    public boolean isAllRowsGenerated() {
      return this.rows.stream()
          .allMatch(rowStructure -> rowStructure.generated);
    }

    public RowStructure findRowByFieldName(String fieldName) {
      return this.rows.stream()
          .filter(rowStructure -> rowStructure.fieldName.equals(fieldName))
          .findAny()
          .orElseThrow(() ->
              new ExcelStructureException(String.format("%s row not founded", fieldName))
          );
    }

    public RowStructure findRowByRowIndex(int rowNum) {
      return this.rows.stream()
          .filter(rowStructure -> rowStructure.generated)
          .filter(rowStructure ->
              rowStructure.startRowNum <= rowNum && rowStructure.endRowNum >= rowNum)
          .findAny()
          .orElseThrow(() ->
              new ExcelStructureException(String.format("row of index %s not founded", rowNum))
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
                rowStructure.startRowNum = (beforeRowStructure.endRowNum +
                    rowStructure.getAnnotation().getRowAfterOffset() +
                    (rowStructure.isDataRowAndHideHeader() ? 0 : 1)
                );
                return true;
              }
              return false;
            }
          })
          .findFirst()
          .orElseThrow(() ->
              new ExcelStructureException("not found nextRowStructure")
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
  @NoArgsConstructor
  public static class RowStructure {

    private AbstractRowAnnotation annotation;
    private Field sheetField;
    private Field field;
    private String fieldName;
    private List<CellStructure> cells = new ArrayList<>();

    public void setAnnotation(AbstractRowAnnotation annotation) {
      this.annotation = annotation;
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

    private boolean generated;
    private int startRowNum;
    private int endRowNum;

    public void setGenerated(boolean generated) {
      this.generated = generated;
    }

    public void setStartRowNum(int startRowNum) {
      this.startRowNum = startRowNum;
    }

    public void setEndRowNum(int endRowNum) {
      this.endRowNum = endRowNum;
    }

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
  @NoArgsConstructor
  public static class CellStructure {

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
      throw new ExcelStructureException(
          String.format("not found excel annotation at %s", dtoType.getName()));
    }
    this.setAnnotation(new ExcelAnnotation(excel));

    this.sheets = Arrays.stream(InheritedFieldHelper.getDeclaredFields(dtoType))
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
            Arrays.stream(InheritedFieldHelper.getDeclaredFields(sheetStructure.field.getType()))
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
                    rowStructure.startRowNum = rowStructure.getAnnotation().getRow();
                    rowStructure.endRowNum = rowStructure.getAnnotation().getRow();
                  }
                  sheetStructure.rows.add(rowStructure);
                  return rowStructure;
                })
                .forEach(rowStructure -> {
                  Field[] fields;
                  boolean isRow = rowStructure.annotation instanceof RowAnnotation;
                  if (isRow) {
                    fields = InheritedFieldHelper.getDeclaredFields(rowStructure.field.getType());
                  } else {
                    ParameterizedType genericType =
                        (ParameterizedType) rowStructure.field.getGenericType();
                    Class<?> dataRowClass = (Class<?>) genericType.getActualTypeArguments()[0];
                    fields = InheritedFieldHelper.getDeclaredFields(dataRowClass);
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

                  if (!isRow) {
                    Match match = ((DataRowsAnnotation) rowStructure.annotation).getMatch();
                    if (Match.REQUIRED.equals(match)) {
                      boolean requiredCellPresent = rowStructure.cells.stream()
                          .anyMatch(cellStructure -> cellStructure.getAnnotation().isRequired());
                      if (!requiredCellPresent) {
                        throw new ExcelStructureException(
                            String.format(
                                "%s row match type is %s, but required cell is not founded in structure",
                                rowStructure.getFieldName(), match.toString()));
                      }
                    }
                  }
                }))
        .collect(Collectors.toList());
    return this;
  }
}
