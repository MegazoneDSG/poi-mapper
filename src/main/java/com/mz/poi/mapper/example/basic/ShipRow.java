package com.mz.poi.mapper.example.basic;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Header;
import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.structure.CellType;
import java.time.LocalDate;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter
@Setter
@NoArgsConstructor
@DataRows(
		rowAfter = "infoTable",
		rowAfterOffset = 1,
		match = Match.REQUIRED,
		headers = {
				@Header(name = "REQUESTER", mappings = {"requester"}),
				@Header(name = "SHIP VIA", mappings = {"via"}),
				@Header(name = "F.O.B", mappings = {"fob"}),
				@Header(name = "SHIPPING TERMS", mappings = {"terms"}),
				@Header(name = "DELIVERY DATE", mappings = {"deliveryDate"})
		},
		headerStyle = @CellStyle(
				font = @Font(color = IndexedColors.WHITE),
				fillForegroundColor = IndexedColors.DARK_BLUE,
				fillPattern = FillPatternType.SOLID_FOREGROUND
		),
		dataStyle = @CellStyle(
				borderTop = BorderStyle.THIN,
				borderBottom = BorderStyle.THIN,
				borderLeft = BorderStyle.THIN,
				borderRight = BorderStyle.THIN
		)
)
public class ShipRow {

	@Cell(
			column = 0,
			cellType = CellType.STRING
	)
	private String requester;

	@Cell(
			column = 1,
			cellType = CellType.STRING
	)
	private String via;

	@Cell(
			column = 2,
			cellType = CellType.STRING
	)
	private String fob;

	@Cell(
			column = 3,
			cols = 2,
			cellType = CellType.STRING
	)
	private String terms;

	@Cell(
			column = 5,
			cellType = CellType.DATE,
			style = @CellStyle(dataFormat = "yyyy-MM-dd"),
			required = true
	)
	private LocalDate deliveryDate;

	@Builder
	public ShipRow(String requester, String via, String fob, String terms,
			LocalDate deliveryDate) {
		this.requester = requester;
		this.via = via;
		this.fob = fob;
		this.terms = terms;
		this.deliveryDate = deliveryDate;
	}
}
