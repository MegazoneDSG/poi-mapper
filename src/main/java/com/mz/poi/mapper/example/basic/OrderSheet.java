package com.mz.poi.mapper.example.basic;

import static org.apache.poi.ss.usermodel.PrintSetup.A4_PAPERSIZE;

import com.mz.poi.mapper.annotation.ColumnWidth;
import com.mz.poi.mapper.annotation.PrintSetup;
import com.mz.poi.mapper.annotation.Sheet;
import java.util.List;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

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
		defaultRowHeightInPoints = 20,
		printSetup = @PrintSetup(
				paperSize = A4_PAPERSIZE
		),
		fitToPage = true
)
public class OrderSheet {

	TitleRow titleRow;

	List<InfoRow> infoTable;

	List<ShipRow> shipTable;

	List<ItemRow> itemTable;

	SummaryRow summaryRow;

	@Builder
	public OrderSheet(TitleRow titleRow, List<InfoRow> infoTable,
			List<ShipRow> shipTable, List<ItemRow> itemTable,
			SummaryRow summaryRow) {
		this.titleRow = titleRow;
		this.infoTable = infoTable;
		this.shipTable = shipTable;
		this.itemTable = itemTable;
		this.summaryRow = summaryRow;
	}
}
