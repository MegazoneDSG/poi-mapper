package com.mz.poi.mapper.example.basic;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Header;
import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.structure.CellType;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter
@Setter
@NoArgsConstructor
@DataRows(
		row = 2,
		match = Match.REQUIRED,
		headers = {
				@Header(name = "VENDOR", mappings = {"vendorTitle", "vendorValue"}),
				@Header(name = "SHIP TO", mappings = {"toTitle", "toValue"})
		},
		headerStyle = @CellStyle(
				font = @Font(color = IndexedColors.WHITE),
				fillForegroundColor = IndexedColors.DARK_BLUE,
				fillPattern = FillPatternType.SOLID_FOREGROUND
		)
)
public class InfoRow {

	@Cell(
			column = 0,
			cellType = CellType.STRING
	)
	private String vendorTitle;

	@Cell(
			column = 1,
			cols = 2,
			cellType = CellType.STRING,
			required = true
	)
	private String vendorValue;

	@Cell(
			column = 3,
			cellType = CellType.STRING
	)
	private String toTitle;

	@Cell(
			column = 4,
			cols = 2,
			cellType = CellType.STRING,
			required = true
	)
	private String toValue;

	@Builder
	public InfoRow(String vendorTitle, String vendorValue, String toTitle, String toValue) {
		this.vendorTitle = vendorTitle;
		this.vendorValue = vendorValue;
		this.toTitle = toTitle;
		this.toValue = toValue;
	}
}
