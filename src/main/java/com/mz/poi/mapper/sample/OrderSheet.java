package com.mz.poi.mapper.sample;

import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Header;
import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.annotation.Row;
import java.util.ArrayList;
import java.util.List;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter
@Setter
@NoArgsConstructor
public class OrderSheet {

  @Row(
      row = 1,
      defaultStyle = @CellStyle(
          font = @Font(fontHeightInPoints = 20)
      ),
      heightInPoints = 50
  )
  TitleRow titleRow = new TitleRow();

  @Row(row = 2)
  Info1 info1 = new Info1();
  @Row(row = 3)
  Info2 info2 = new Info2();
  @Row(row = 4)
  Info3 info3 = new Info3();

  @DataRows(
      row = 5,
      match = Match.REQUIRED,
      headerStyle = @CellStyle(
          font = @Font(color = IndexedColors.DARK_RED),
          fillBackgroundColor = IndexedColors.AQUA
      ),
      headerHeightInPoints = 15,
      headers = {
          @Header(column = 0, name = "상품이름"),
          @Header(column = 2, name = "상품번호"),
          @Header(column = 3, name = "상품아이디"),
          @Header(column = 4, cols = 2, name = "합계")
      },
      dataStyle = @CellStyle(
          borderBottom = BorderStyle.DASH_DOT
      )
  )
  List<OrderDataRow> items = new ArrayList<>();

  @Row(
      rowAfter = "items",
      rowAfterOffset = 2
  )
  SummaryRow summaryRow = new SummaryRow();
}
