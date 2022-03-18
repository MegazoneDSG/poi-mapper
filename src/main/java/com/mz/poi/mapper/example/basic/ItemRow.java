package com.mz.poi.mapper.example.basic;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.Constraint;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Header;
import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.structure.CellType;
import java.math.BigDecimal;
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
		rowAfter = "shipTable",
		rowAfterOffset = 1,
		match = Match.REQUIRED,
		headers = {
				@Header(name = "ITEM", mappings = {"name"}),
				@Header(name = "DESCRIPTION", mappings = {"description"}),
				@Header(name = "QTY", mappings = {"qty"}),
				@Header(name = "UNIT PRICE", mappings = {"unitPrice"}),
				@Header(name = "TOTAL", mappings = {"total"})
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
public class ItemRow {

	@Cell(
			column = 0,
			cellType = CellType.STRING,
			required = true,
			constraint = @Constraint(
					constraints = {"A", "B"},
					errorBoxTitle = "ERROR!",
					errorBoxText = "값을 올바로 선택해 주세요"
			)
	)
	private String name;

	@Cell(
			column = 1,
			cols = 2,
			cellType = CellType.STRING,
			required = true
	)
	private String description;

	@Cell(
			column = 3,
			cellType = CellType.NUMERIC,
			required = true
	)
	private long qty;

	@Cell(
			column = 4,
			cellType = CellType.NUMERIC,
			style = @CellStyle(dataFormat = "#,##0.00"),
			required = true
	)
	private BigDecimal unitPrice;

	@Cell(
			column = 5,
			cellType = CellType.FORMULA,
			style = @CellStyle(
					dataFormat = "#,##0.00",
					fillForegroundColor = IndexedColors.GREY_25_PERCENT,
					fillPattern = FillPatternType.SOLID_FOREGROUND
			),
			ignoreParse = true
	)
	private String total = "product({{this.qty}},{{this.unitPrice}})";

	@Builder
	public ItemRow(String name, String description, long qty, BigDecimal unitPrice) {
		this.name = name;
		this.description = description;
		this.qty = qty;
		this.unitPrice = unitPrice;
	}
}
