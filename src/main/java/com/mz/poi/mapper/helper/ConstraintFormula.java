package com.mz.poi.mapper.helper;

import com.mz.poi.mapper.structure.CellStructure;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

public class ConstraintFormula {
	private final int startRow;
	private final int endRow;
	private final String formula;
	private final String constraintKey;
	private final DataValidationConstraint dataValidationConstraint;

	public int getStartRow() {
		return startRow;
	}

	public int getEndRow() {
		return endRow;
	}

	public String getFormula() {
		return formula;
	}

	public String getConstraintKey() {
		return constraintKey;
	}

	public DataValidationConstraint getDataValidationConstraint() {
		return dataValidationConstraint;
	}

	public ConstraintFormula(int startRow, int endRow, String formula, String constraintKey, DataValidationConstraint dataValidationConstraint) {
		this.startRow = startRow;
		this.endRow = endRow;
		this.formula = formula;
		this.constraintKey = constraintKey;
		this.dataValidationConstraint = dataValidationConstraint;
	}

	public static ConstraintFormula of(Map<String, ConstraintFormula> constraintFormulaMap, CellStructure cellStructure, Sheet constraintSheet, Sheet validateSheet) {
		List<String> strings = Arrays.asList(cellStructure.getAnnotation().getConstraint().getConstraints());
		int startRow = constraintFormulaMap.values().stream().mapToInt(v -> v.getEndRow() + 1).max().orElse(0);
		int endRow = startRow + strings.size() - 1;
		String formula = constraintSheet.getSheetName() + "!$A$" + (startRow + 1) + ":$A$" + (endRow + 1);

		for (int i = 0; i < strings.size(); i++) {
			constraintSheet.createRow(startRow + i).createCell(0).setCellValue(strings.get(i));
		}
		DataValidationHelper dataValidationHelper = validateSheet.getDataValidationHelper();
		DataValidationConstraint dataValidationConstraint = dataValidationHelper.createFormulaListConstraint(formula);

		String key = getKey(cellStructure);
		ConstraintFormula constraintFormula = new ConstraintFormula(startRow, endRow, formula, key, dataValidationConstraint);
		constraintFormulaMap.put(key, constraintFormula);
		return constraintFormula;
	}

	public static String getKey(CellStructure cellStructure) {
		return cellStructure.getRowStructure().getSheetStructure().getFieldName() + "-" +
			cellStructure.getRowStructure().getFieldName() + "-" +
			cellStructure.getFieldName();
	}
}
