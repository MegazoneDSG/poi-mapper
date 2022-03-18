package com.mz.poi.mapper.example.object;

import com.mz.poi.mapper.annotation.ArrayCell;
import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.ColumnWidth;
import com.mz.poi.mapper.annotation.Excel;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Row;
import com.mz.poi.mapper.annotation.Sheet;
import com.mz.poi.mapper.expression.CellGenerator;
import com.mz.poi.mapper.expression.CustomCellExpression;
import com.mz.poi.mapper.helper.CachedStyleRepository;
import com.mz.poi.mapper.structure.CellStyleAnnotation;
import com.mz.poi.mapper.structure.CellType;
import java.time.LocalDate;
import java.util.List;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@Excel(
		defaultStyle = @CellStyle(
				font = @Font(fontName = "Arial")
		),
		dateFormatZoneId = "Asia/Seoul"
)
public class ObjectTemplate {

	private ObjectSheet sheet = new ObjectSheet();

	@Getter
	@Setter
	@NoArgsConstructor
	@Sheet(
			name = "Order",
			index = 0,
			columnWidths = {
					@ColumnWidth(column = 0, width = 25)
			},
			defaultColumnWidth = 20,
			defaultRowHeightInPoints = 20
	)
	public static class ObjectSheet {

		ObjectRow objectRow = new ObjectRow();

		@Getter
		@Setter
		@NoArgsConstructor
		@Row(
				row = 0
		)
		public static class ObjectRow {

			@Cell(
					column = 0,
					cellType = CellType.NONE,
					style = @CellStyle(dataFormat = "@")
			)
			private Object text;

			@Cell(
					columnAfter = "text",
					cellType = CellType.NONE,
					style = @CellStyle(dataFormat = "0")
			)
			private Object number;

			@ArrayCell(
					columnAfter = "number",
					cellType = CellType.NONE,
					size = 5
			)
			private List<Object> arrayObject;

			@ArrayCell(
					columnAfter = "arrayObject",
					cellType = CellType.NONE,
					size = 5
			)
			private List<TestObjectCellExpression> arrayCustom;

			@Builder
			public ObjectRow(Object text, Object number, List<Object> arrayObject,
					List<TestObjectCellExpression> arrayCustom) {
				this.text = text;
				this.number = number;
				this.arrayObject = arrayObject;
				this.arrayCustom = arrayCustom;
			}
		}

		@Builder
		public ObjectSheet(
				ObjectRow objectRow) {
			this.objectRow = objectRow;
		}
	}

	@Builder
	public ObjectTemplate(
			ObjectSheet sheet) {
		this.sheet = sheet;
	}

	public static class TestObjectCellExpression extends CustomCellExpression<Object> {

		public TestObjectCellExpression(Object value) {
			super(value);
		}

		@Override
		public CellGenerator toCell(CellStyleAnnotation styleAnnotation,
				CachedStyleRepository styleRepository) {
			if (this.getValue() instanceof LocalDate) {
				CellStyleAnnotation copy = styleAnnotation.copy();
				copy.setDataFormat("yyyy-MM-dd");
				copy.setKey("TestObjectCellExpressionLocalDate");
				return CellGenerator.builder()
						.cellType(CellType.DATE)
						.style(styleRepository.createStyle(copy))
						.value(this.getValue())
						.build();

			} else if (this.getValue() instanceof Number) {
				CellStyleAnnotation copy = styleAnnotation.copy();
				copy.setDataFormat("0");
				copy.setKey("TestObjectCellExpressionNumber");
				return CellGenerator.builder()
						.cellType(CellType.NUMERIC)
						.style(styleRepository.createStyle(copy))
						.value(this.getValue())
						.build();
			} else {
				CellStyleAnnotation copy = styleAnnotation.copy();
				copy.setDataFormat("@");
				copy.setKey("TestObjectCellExpressionString");
				return CellGenerator.builder()
						.cellType(CellType.STRING)
						.style(styleRepository.createStyle(copy))
						.value(this.getValue())
						.build();
			}
		}

		@Override
		public void fromCell(org.apache.poi.ss.usermodel.Cell cell) {
			if (org.apache.poi.ss.usermodel.CellType.STRING.equals(cell.getCellType())) {
				this.setValue(cell.getStringCellValue());
			} else if (org.apache.poi.ss.usermodel.CellType.NUMERIC.equals(cell.getCellType())) {
				this.setValue(cell.getNumericCellValue());
			}
		}
	}
}
