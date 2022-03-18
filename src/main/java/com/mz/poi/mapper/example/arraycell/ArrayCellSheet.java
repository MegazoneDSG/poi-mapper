package com.mz.poi.mapper.example.arraycell;

import com.mz.poi.mapper.annotation.ColumnWidth;
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
		defaultRowHeightInPoints = 20
)
public class ArrayCellSheet {

	List<ItemRow> itemTable;

	SummaryRow summaryRow;

	@Builder
	public ArrayCellSheet(List<ItemRow> itemTable,
			SummaryRow summaryRow) {
		this.itemTable = itemTable;
		this.summaryRow = summaryRow;
	}
}
