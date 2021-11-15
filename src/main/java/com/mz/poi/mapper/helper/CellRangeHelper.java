package com.mz.poi.mapper.helper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

public class CellRangeHelper {

	public static CellRangeAddress getCellRangeAddress(Cell cell, int columnIndex, int cols) {
		return new CellRangeAddress(
			cell.getRow().getRowNum(), cell.getRow().getRowNum(),
			columnIndex, (columnIndex + cols - 1));
	}

	public static CellRangeAddressList getCellRangeAddressList(Cell cell, int columnIndex, int cols) {
		return new CellRangeAddressList(cell.getRow().getRowNum(), cell.getRow().getRowNum(), columnIndex, (columnIndex + cols - 1));
	}
}
