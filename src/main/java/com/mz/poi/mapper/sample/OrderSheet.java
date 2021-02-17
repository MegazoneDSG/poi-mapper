package com.mz.poi.mapper.sample;

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
public class OrderSheet {

  @Row(
      row = 0,
      defaultStyle = @CellStyle(
          font = @Font(fontHeightInPoints = 20)
      ),
      heightInPoints = 40
  )
  TitleRow titleRow;

  @DataRows(
      row = 2,
      match = Match.REQUIRED,
      headers = {
          @Header(name = "VENDOR", column = 0, cols = 3),
          @Header(name = "SHIP TO", column = 3, cols = 3)
      },
      headerStyle = @CellStyle(
          font = @Font(color = IndexedColors.WHITE),
          fillForegroundColor = IndexedColors.DARK_BLUE,
          fillPattern = FillPatternType.SOLID_FOREGROUND
      )
  )
  List<InfoRow> infoTable;

  @DataRows(
      rowAfter = "infoTable",
      rowAfterOffset = 1,
      match = Match.REQUIRED,
      headers = {
          @Header(name = "REQUESTER", column = 0),
          @Header(name = "SHIP VIA", column = 1),
          @Header(name = "F.O.B", column = 2),
          @Header(name = "SHIPPING TERMS", column = 3, cols = 2),
          @Header(name = "DELIVERY DATE", column = 5)
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
  List<ShipRow> shipTable;

  @DataRows(
      rowAfter = "shipTable",
      rowAfterOffset = 1,
      match = Match.REQUIRED,
      headers = {
          @Header(name = "ITEM", column = 0),
          @Header(name = "DESCRIPTION", column = 1, cols = 2),
          @Header(name = "QTY", column = 3),
          @Header(name = "UNIT PRICE", column = 4),
          @Header(name = "TOTAL", column = 5)
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
