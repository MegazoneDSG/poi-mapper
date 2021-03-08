package com.mz.poi.mapper.example.arraycell;

import com.mz.poi.mapper.annotation.ArrayHeader;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Header;
import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.annotation.Row;
import java.util.List;
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
public class ArrayCellSheet {

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
  List<ItemRow> itemTable;

  @Row(rowAfter = "itemTable")
  SummaryRow summaryRow;

  @Builder
  public ArrayCellSheet(List<ItemRow> itemTable,
      SummaryRow summaryRow) {
    this.itemTable = itemTable;
    this.summaryRow = summaryRow;
  }
}
