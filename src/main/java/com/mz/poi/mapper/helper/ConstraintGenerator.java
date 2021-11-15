package com.mz.poi.mapper.helper;

import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.ConstraintAnnotation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class ConstraintGenerator {

	private static String SHEET_NAME = "CONSTRAINT_SHEET";
	private Workbook workbook;
	private Map<String, ConstraintFormula> constraintFormulaMap = new ConcurrentHashMap<>();
	private Sheet constraintSheet;

	public ConstraintGenerator(Workbook workbook) {
		this.workbook = workbook;
	}

	public void createConstraint(CellStructure cellStructure, Cell cell, int columnIndex, int cols) {
		ConstraintAnnotation constraint = cellStructure.getAnnotation().getConstraint();
		String[] constraints = constraint.getConstraints();
		if (constraints.length == 0) {
			return;
		}
		if (constraintSheet == null) {
			this.createConstraintSheet();
		}
		ConstraintFormula constraintFormula = this.getConstraintFormula(cellStructure);
		CellRangeAddressList cellRangeAddressList = CellRangeHelper.getCellRangeAddressList(cell, columnIndex, cols);
		DataValidation validation = cell.getRow().getSheet().getDataValidationHelper().createValidation(constraintFormula.getDataValidationConstraint(), cellRangeAddressList);
		validation.setSuppressDropDownArrow(constraint.isSuppressDropDownArrow());
		validation.setShowErrorBox(constraint.isShowErrorBox());
		validation.createErrorBox(constraint.getErrorBoxTitle(), constraint.getErrorBoxText());
		validation.setErrorStyle(constraint.getErrorStyle());
		cell.getRow().getSheet().addValidationData(validation);
	}

	private void createConstraintSheet() {
		constraintSheet = workbook.createSheet(SHEET_NAME);
		workbook.setSheetHidden(workbook.getSheetIndex(SHEET_NAME), true);
	}

	private ConstraintFormula getConstraintFormula(CellStructure cellStructure) {
		String key = ConstraintFormula.getKey(cellStructure);
		if (!constraintFormulaMap.containsKey(key)) {
			String validateSheetName = cellStructure.getRowStructure().getSheetStructure().getAnnotation().getName();
			return ConstraintFormula.of(constraintFormulaMap, cellStructure, constraintSheet, workbook.getSheet(validateSheetName));
		}
		return constraintFormulaMap.get(key);
	}
}

