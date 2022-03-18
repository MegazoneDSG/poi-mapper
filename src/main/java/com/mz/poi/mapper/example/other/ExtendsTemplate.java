package com.mz.poi.mapper.example.other;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Excel;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Header;
import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.annotation.Row;
import com.mz.poi.mapper.annotation.Sheet;
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
		)
)
public class ExtendsTemplate {

	private TestSheet sheet = new TestSheet();

	@Builder
	public ExtendsTemplate(TestSheet sheet) {
		this.sheet = sheet;
	}

	@Getter
	@Setter
	@NoArgsConstructor
	@Sheet(
			name = "Test",
			index = 0
	)
	public static class TestSheet extends TestSheetSuper {

		public OverrideRow overrideRow = new OverrideRow();

		private List<FirstTableRow> firstTable;

		@Builder
		public TestSheet(List<FirstTableRow> firstTable, List<SecondTableRow> secondTable) {
			this.firstTable = firstTable;
			this.setSecondTable(secondTable);
		}
	}

	@Getter
	@Setter
	@NoArgsConstructor
	public static class TestSheetSuper {

		private List<SecondTableRow> secondTable;
	}

	@Getter
	@Setter
	@NoArgsConstructor
	@Row(row = 0)
	public static class OverrideRow extends OverrideRowSuper {

		@Cell(
				column = 0,
				cellType = CellType.STRING,
				ignoreParse = true
		)
		private String title = "OverrideRow";
	}

	@Getter
	@Setter
	@NoArgsConstructor
	public static class OverrideRowSuper {

		@Cell(
				column = 0,
				cellType = CellType.STRING,
				ignoreParse = true
		)
		private String title = "OverrideRowSuper";
	}

	@Getter
	@Setter
	@NoArgsConstructor
	@DataRows(
			row = 1,
			match = Match.REQUIRED,
			headers = {
					@Header(name = "a", mappings = {"firstValue"}),
					@Header(name = "b", mappings = {"secondValue"})
			}
	)
	public static class FirstTableRow extends FirstTableRowSuper {

		@Cell(
				column = 0,
				cellType = CellType.STRING
		)
		private String firstValue;

		@Builder
		public FirstTableRow(String firstValue, String secondValue) {
			this.firstValue = firstValue;
			this.setSecondValue(secondValue);
		}
	}

	@Getter
	@Setter
	@NoArgsConstructor
	public static class FirstTableRowSuper {

		@Cell(
				column = 1,
				cellType = CellType.STRING,
				required = true
		)
		private String secondValue;
	}

	@Getter
	@Setter
	@NoArgsConstructor
	@DataRows(
			rowAfter = "firstTable",
			match = Match.REQUIRED,
			headers = {
					@Header(name = "c", mappings = {"firstValue"})
			},
			hideHeader = true
	)
	public static class SecondTableRow {

		@Cell(
				column = 0,
				cellType = CellType.DATE,
				style = @CellStyle(
						dataFormat = "yyyy-mm-dd"
				),
				required = true
		)
		private LocalDate firstValue;

		@Builder
		public SecondTableRow(LocalDate firstValue) {
			this.firstValue = firstValue;
		}
	}
}
