package com.mz.poi.mapper.helper;

import com.mz.poi.mapper.exception.ExcelGenerateException;
import com.mz.poi.mapper.structure.ExcelStructure.RowStructure;
import com.mz.poi.mapper.structure.ExcelStructure.SheetStructure;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;

@NoArgsConstructor
public class FormulaHelper {

  private Map<String, List<FormulaItem>> formulaItemsMap = new HashMap<>();

  public void addFormula(Cell cell, String formula) {
    String sheetName = cell.getRow().getSheet().getSheetName();
    if (!formulaItemsMap.containsKey(sheetName)) {
      formulaItemsMap.put(sheetName, new ArrayList<>());
    }
    this.formulaItemsMap.get(sheetName).add(
        FormulaItem.builder()
            .cell(cell)
            .formula(formula)
            .build()
    );
  }

  public void applySheetFormulas(SheetStructure sheetStructure) {
    String sheetName = sheetStructure.getAnnotation().getName();
    if (!formulaItemsMap.containsKey(sheetName)) {
      return;
    }
    this.formulaItemsMap.get(sheetName).forEach(item -> {
      String addressRx = "\\{\\{(.+?)}}";
      Matcher m = Pattern.compile(addressRx).matcher(item.formula);
      while (m.find()) {
        String addressToReplace = m.group(0);
        String addressStr = m.group(1);
        String[] split = addressStr.split("\\.");
        Optional<String> cellAddress = Optional.empty();
        if (split.length == 2 && "this".equals(split[0])) {
          cellAddress = this.getSelfDataRowCellAddress(sheetStructure, addressStr, item.getCell());
        } else if (split.length == 2) {
          cellAddress = this.getRowCellAddress(sheetStructure, addressStr);
        } else if (split.length == 3) {
          cellAddress = this.getIndexedDataRowCellAddress(sheetStructure, addressStr);
        }
        String finalAddress = cellAddress.orElseThrow(() ->
            new ExcelGenerateException(String.format("Not found cellAddress %s", addressStr)));
        item.formula = item.formula.replace(addressToReplace, finalAddress);
      }
      item.getCell().setCellFormula(item.getFormula());
    });
  }

  private Optional<String> getRowCellAddress(SheetStructure sheetStructure, String addressStr) {
    String[] split = addressStr.split("\\.");
    String rowFieldName = split[0];
    String cellFieldName = split[1];
    RowStructure rowStructure = sheetStructure.findRowByFieldName(rowFieldName);
    return rowStructure.getCells().stream()
        .filter(cellStructure ->
            cellStructure.getFieldName().equals(cellFieldName))
        .map(cellStructure ->
            new CellAddress(
                rowStructure.getStartRowNum(),
                cellStructure.getAnnotation().getColumn()
            ).formatAsString())
        .findFirst();
  }

  private Optional<String> getIndexedDataRowCellAddress(SheetStructure sheetStructure,
      String addressStr) {
    String[] split = addressStr.split("\\.");
    String rowFieldName = split[0];
    String rowNumRex = split[1];
    String cellFieldName = split[2];

    RowStructure rowStructure = sheetStructure.findRowByFieldName(rowFieldName);
    int rowNum = 0;
    if ("last".equals(rowNumRex)) {
      rowNum = rowStructure.getEndRowNum();
    } else {
      String atRex = "at\\(([0-9]+?)\\)";
      Matcher m = Pattern.compile(atRex).matcher(rowNumRex);
      boolean matched = false;
      while (m.find()) {
        matched = true;
        int dataStartNum = rowStructure.getStartRowNum() + 1; // 헤더를 제외하고 시작해야 하기 때문에 1 추가한다.
        int at = Integer.parseInt(m.group(1));
        rowNum = dataStartNum + at;
      }
      if (!matched) {
        throw new ExcelGenerateException(String.format("Invalid formula pattern %s", addressStr));
      }
    }

    int finalRowNum = rowNum;
    return rowStructure.getCells().stream()
        .filter(cellStructure ->
            cellStructure.getFieldName().equals(cellFieldName))
        .map(cellStructure -> new CellAddress(
            finalRowNum, cellStructure.getAnnotation().getColumn()
        ).formatAsString())
        .findFirst();
  }

  private Optional<String> getSelfDataRowCellAddress(SheetStructure sheetStructure,
      String addressStr, Cell cell) {
    String[] split = addressStr.split("\\.");
    String cellFieldName = split[1];

    RowStructure rowStructure = sheetStructure.findRowByRowIndex(cell.getRowIndex());
    return rowStructure.getCells().stream()
        .filter(cellStructure ->
            cellStructure.getFieldName().equals(cellFieldName))
        .map(cellStructure -> new CellAddress(
            cell.getRowIndex(), cellStructure.getAnnotation().getColumn()
        ).formatAsString())
        .findFirst();
  }

  @Getter
  @Setter
  @NoArgsConstructor
  public static class FormulaItem {

    private Cell cell;
    private String formula;

    @Builder
    public FormulaItem(Cell cell, String formula) {
      this.cell = cell;
      this.formula = formula;
    }
  }
}
