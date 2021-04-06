package com.mz.poi.mapper.example.basic;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.structure.CellType;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.math.BigDecimal;

@Getter
@Setter
@NoArgsConstructor
public class ItemRow {

    @Cell(
        column = 0,
        cellType = CellType.STRING,
        required = true
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
