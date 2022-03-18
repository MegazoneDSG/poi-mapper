package com.mz.poi.mapper.example.basic;

import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.Excel;
import com.mz.poi.mapper.annotation.Font;
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
public class PurchaseOrderTemplate {

	private OrderSheet sheet = new OrderSheet();

	@Builder
	public PurchaseOrderTemplate(OrderSheet sheet) {
		this.sheet = sheet;
	}
}
