package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.ArrayCell;
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
import java.util.Optional;
import java.util.stream.Collectors;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@NoArgsConstructor
public class ExcelStructure {

	private Class<?> dtoType;
	private ExcelAnnotation annotation;
	private List<SheetStructure> sheets = new ArrayList<>();

	public void setSheets(List<SheetStructure> sheets) {
		this.sheets = sheets;
	}

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

	public void prepareReadStructure() {
		this.sheets.forEach(sheetStructure -> {
			sheetStructure.getRows().forEach(rowStructure -> {
				rowStructure.setRead(false);
				rowStructure.setStartRowNum(0);
				rowStructure.setEndRowNum(0);

				rowStructure.getCells().forEach(cellStructure -> {
					cellStructure.setCalculated(false);
					cellStructure.calculateColumn();
				});
			});
		});
	}

	public void prepareGenerateStructure(Object excelDto) {
		this.sheets.forEach(sheetStructure -> {
			sheetStructure.getRows().forEach(rowStructure -> {
				rowStructure.setCalculated(false);
				rowStructure.calculateRowNum(excelDto);

				rowStructure.getCells().forEach(cellStructure -> {
					cellStructure.setCalculated(false);
					cellStructure.calculateColumn();
				});
			});
		});
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
					Sheet sheet = field.getType().getAnnotation(Sheet.class);
					return sheet != null;
				})
				.map(field -> {
					SheetStructure sheetStructure = SheetStructure.builder()
							.excelStructure(this)
							.annotation(
									new SheetAnnotation(
											field.getType().getAnnotation(Sheet.class),
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
						Arrays
								.stream(InheritedFieldHelper.getDeclaredFields(
										sheetStructure.getField().getType()))
								.filter(field -> {
									field.setAccessible(true);
									Row row = field.getType().getAnnotation(Row.class);
									DataRows dataRows = null;
									if (Collection.class.isAssignableFrom(field.getType())) {
										Class<?> genericClass = InheritedFieldHelper.getGenericClass(
												field);
										dataRows = genericClass.getAnnotation(DataRows.class);
									}
									return row != null || dataRows != null;
								})
								.map(field -> {
									RowStructure rowStructure = RowStructure.builder()
											.sheetStructure(sheetStructure)
											.sheetField(sheetStructure.getField())
											.field(field)
											.fieldName(field.getName())
											.build();
									boolean isRow =
											field.getType().getAnnotation(Row.class) != null;
									if (isRow) {
										rowStructure.setAnnotation(
												new RowAnnotation(
														field.getType().getAnnotation(Row.class),
														sheetStructure.getAnnotation()
																.getDefaultStyle()
												)
										);
									} else {
										Class<?> genericClass = InheritedFieldHelper.getGenericClass(
												field);
										rowStructure.setAnnotation(
												new DataRowsAnnotation(
														genericClass.getAnnotation(DataRows.class),
														sheetStructure.getAnnotation()
																.getDefaultStyle()
												)
										);
									}
									sheetStructure.getRows().add(rowStructure);
									return rowStructure;
								})
								.forEach(rowStructure -> {
									Field[] fields;
									boolean isRow = rowStructure.getAnnotation() instanceof RowAnnotation;
									if (isRow) {
										fields = InheritedFieldHelper
												.getDeclaredFields(
														rowStructure.getField().getType());
									} else {
										ParameterizedType genericType =
												(ParameterizedType) rowStructure.getField()
														.getGenericType();
										Class<?> dataRowClass = (Class<?>) genericType.getActualTypeArguments()[0];
										fields = InheritedFieldHelper.getDeclaredFields(
												dataRowClass);
									}

									Arrays.stream(fields)
											.filter(field -> {
												field.setAccessible(true);
												Cell cell = field.getAnnotation(Cell.class);
												ArrayCell arrayCell = field.getAnnotation(
														ArrayCell.class);
												return cell != null ||
														(arrayCell != null &&
																Collection.class.isAssignableFrom(
																		field.getType()));
											})
											.forEach(field -> {
												CellStructure cellStructure = CellStructure.builder()
														.rowStructure(rowStructure)
														.sheetField(rowStructure.getSheetField())
														.rowField(rowStructure.getField())
														.field(field)
														.fieldName(field.getName())
														.build();

												boolean isCell =
														field.getAnnotation(Cell.class) != null;
												if (isCell) {
													cellStructure.setAnnotation(
															new CellAnnotation(
																	field.getAnnotation(Cell.class),
																	isRow ?
																			((RowAnnotation) rowStructure.getAnnotation())
																					.getDefaultStyle()
																			:
																					((DataRowsAnnotation) rowStructure.getAnnotation())
																							.getDataStyle())
													);
												} else {
													cellStructure.setAnnotation(
															new ArrayCellAnnotation(
																	field.getAnnotation(
																			ArrayCell.class),
																	isRow ?
																			((RowAnnotation) rowStructure.getAnnotation())
																					.getDefaultStyle()
																			:
																					((DataRowsAnnotation) rowStructure.getAnnotation())
																							.getDataStyle())
													);
												}
												rowStructure.getCells().add(cellStructure);
											});

									if (!isRow) {
										Match match = ((DataRowsAnnotation) rowStructure.getAnnotation()).getMatch();
										if (Match.REQUIRED.equals(match)) {
											boolean requiredCellPresent = rowStructure.getCells()
													.stream()
													.anyMatch(
															cellStructure -> cellStructure.getAnnotation()
																	.isRequired());
											if (!requiredCellPresent) {
												throw new ExcelStructureException(
														String.format(
																"%s row match type is %s, but required cell is not founded in structure",
																rowStructure.getFieldName(),
																match.toString()));
											}
										}
									}
								}))
				.collect(Collectors.toList());
		return this;
	}
}
