package com.mz.poi.mapper.example.basic;

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
          @Header(name = "VENDOR", mappings = {"vendorTitle", "vendorValue"}),
          @Header(name = "SHIP TO", mappings = {"toTitle", "toValue"})
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
  List<ShipRow> shipTable;

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
