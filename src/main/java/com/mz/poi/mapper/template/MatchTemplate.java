package com.mz.poi.mapper.template;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.annotation.CellStyle;
import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Excel;
import com.mz.poi.mapper.annotation.Font;
import com.mz.poi.mapper.annotation.Header;
import com.mz.poi.mapper.annotation.Match;
import com.mz.poi.mapper.annotation.Row;
import com.mz.poi.mapper.annotation.Sheet;
import com.mz.poi.mapper.structure.CellType;
import java.util.List;
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
    )
)
public class MatchTemplate {

  @Sheet(
      name = "Test",
      index = 0
  )
  private TestSheet sheet = new TestSheet();

  @Builder
  public MatchTemplate(TestSheet sheet) {
    this.sheet = sheet;
  }

  @Getter
  @Setter
  @NoArgsConstructor
  public static class TestSheet {

    @DataRows(
        row = 0,
        match = Match.STOP_ON_BLANK,
        headers = {
            @Header(name = "a", column = 0),
            @Header(name = "b", column = 1)
        }
    )
    private List<MatchTestDataRow> matchTestTable;

    @Row(
        rowAfter = "matchTestTable",
        rowAfterOffset = 1
    )
    private MatchTestRow matchTestRow = new MatchTestRow();

    @Builder
    public TestSheet(
        List<MatchTestDataRow> matchTestTable) {
      this.matchTestTable = matchTestTable;
    }
  }


  @Getter
  @Setter
  @NoArgsConstructor
  public static class MatchTestDataRow {

    @Cell(
        column = 0,
        cellType = CellType.STRING
    )
    private String firstValue;

    @Cell(
        column = 1,
        cellType = CellType.NUMERIC
    )
    private Integer secondValue;

    @Builder
    public MatchTestDataRow(String firstValue, Integer secondValue) {
      this.firstValue = firstValue;
      this.secondValue = secondValue;
    }
  }

  @Getter
  @Setter
  @NoArgsConstructor
  public static class MatchTestRow {

    @Cell(
        column = 0,
        cellType = CellType.STRING
    )
    private String value = "sample";
  }
}
