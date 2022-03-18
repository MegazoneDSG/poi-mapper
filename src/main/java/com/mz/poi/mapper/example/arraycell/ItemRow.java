package com.mz.poi.mapper.example.arraycell;

import com.mz.poi.mapper.annotation.ArrayCell;
import com.mz.poi.mapper.annotation.ArrayHeader;
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
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.math.BigDecimal;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@DataRows(
        row = 0,
        match = Match.REQUIRED,
        headers = {
                @Header(name = "ITEM & DESCRIPTION", mappings = {"name", "description"}),
                @Header(name = "UNIT PRICE", mappings = {"unitPrice"}),
                @Header(name = "TOTAL", mappings = {"total"})
        },
        arrayHeaders = {
                @ArrayHeader(simpleNameExpression = "QTY {{index}}", mapping = "qty")
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
      required = true
  )
  private String name;

  @Cell(
      columnAfter = "name",
      cols = 2,
      cellType = CellType.STRING,
      required = true
  )
  private String description;

  @ArrayCell(
      size = 4,
      columnAfter = "description",
      cellType = CellType.NUMERIC,
      required = true
  )
  private List<Integer> qty;

  @Cell(
      columnAfter = "qty",
      cellType = CellType.NUMERIC,
      style = @CellStyle(dataFormat = "#,##0.00"),
      required = true
  )
  private BigDecimal unitPrice;

  @Cell(
      columnAfter = "unitPrice",
      cellType = CellType.FORMULA,
      style = @CellStyle(
          dataFormat = "#,##0.00",
          fillForegroundColor = IndexedColors.GREY_25_PERCENT,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      ),
      ignoreParse = true
  )
  private String total = "product(sum({{this.qty[0]}}:{{this.qty[last]}}),{{this.unitPrice}})";

  @Builder
  public ItemRow(String name, String description, List<Integer> qty, BigDecimal unitPrice) {
    this.name = name;
    this.description = description;
    this.qty = qty;
    this.unitPrice = unitPrice;
  }
}
