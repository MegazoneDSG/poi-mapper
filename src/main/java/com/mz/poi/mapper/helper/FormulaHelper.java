package com.mz.poi.mapper.helper;

import com.mz.poi.mapper.exception.ExcelGenerateException;
import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.RowStructure;
import com.mz.poi.mapper.structure.SheetStructure;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;

@NoArgsConstructor
public class FormulaHelper {

  public void applyFormula(Cell cell, String formulaExpresion, CellStructure cellStructure,
      int rowIndex) {
    String addressRx = "\\{\\{(.+?)}}";
    String formula = formulaExpresion;
    Matcher m = Pattern.compile(addressRx).matcher(formulaExpresion);
    while (m.find()) {
      String addressToReplace = m.group(0);
      String addressStr = m.group(1);
      String[] split = addressStr.split("\\.");
      Optional<String> cellAddress = Optional.empty();
      if (split.length == 2 && "this".equals(split[0])) {
        cellAddress = this.getSelfDataRowCellAddress(cellStructure, addressStr, rowIndex);
      } else if (split.length == 2) {
        cellAddress = this.getRowCellAddress(cellStructure, addressStr);
      } else if (split.length == 3) {
        cellAddress = this.getIndexedDataRowCellAddress(cellStructure, addressStr);
      }
      String finalAddress = cellAddress.orElseThrow(() ->
          new ExcelGenerateException(String.format("Not found cellAddress %s", addressStr)));
      formula = formula.replace(addressToReplace, finalAddress);
    }
    cell.setCellFormula(formula);
  }

  private Optional<String> getRowCellAddress(CellStructure cellStructure, String addressStr) {
    String[] split = addressStr.split("\\.");
    String rowFieldName = split[0];
    String cellFieldName = split[1];
    RowStructure rowStructure = cellStructure.getRowStructure().getSheetStructure()
        .findRowByFieldName(rowFieldName);
    return rowStructure.getCells().stream()
        .filter(_cellStructure ->
            _cellStructure.getFieldName().equals(cellFieldName))
        .map(_cellStructure ->
            new CellAddress(
                rowStructure.getStartRowNum(),
                _cellStructure.getAnnotation().getColumn()
            ).formatAsString())
        .findFirst();
  }

  private Optional<String> getIndexedDataRowCellAddress(CellStructure cellStructure,
      String addressStr) {
    String[] split = addressStr.split("\\.");
    String rowFieldName = split[0];
    String rowNumRex = split[1];
    String cellFieldName = split[2];

    RowStructure rowStructure = cellStructure.getRowStructure().getSheetStructure()
        .findRowByFieldName(rowFieldName);
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
        .filter(_cellStructure ->
            _cellStructure.getFieldName().equals(cellFieldName))
        .map(_cellStructure -> new CellAddress(
            finalRowNum, _cellStructure.getAnnotation().getColumn()
        ).formatAsString())
        .findFirst();
  }

  private Optional<String> getSelfDataRowCellAddress(CellStructure cellStructure,
      String addressStr, int rowIndex) {
    String[] split = addressStr.split("\\.");
    String cellFieldName = split[1];

    RowStructure rowStructure = cellStructure.getRowStructure();
    return rowStructure.getCells().stream()
        .filter(_cellStructure ->
            _cellStructure.getFieldName().equals(cellFieldName))
        .map(_cellStructure -> new CellAddress(
            rowIndex, _cellStructure.getAnnotation().getColumn()
        ).formatAsString())
        .findFirst();
  }
}
