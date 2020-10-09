package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.DataRows;
import com.mz.poi.mapper.annotation.Match;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class DataRowsAnnotation extends AbstractRowAnnotation{

  private Match match;
  private CellStyleAnnotation dataStyle;
  private boolean useDataHeightInPoints = false;
  private int dataHeightInPoints;

  private CellStyleAnnotation headerStyle;
  private boolean useHeaderHeightInPoints = false;
  private int headerHeightInPoints;

  private List<HeaderAnnotation> headers = new ArrayList<>();

  public DataRowsAnnotation(DataRows row, CellStyleAnnotation sheetStyle) {
    this.setRow(row.row());
    this.setRowAfter(row.rowAfter());
    this.setRowAfterOffset(row.rowAfterOffset());
    this.match = row.match();
    this.dataStyle = new CellStyleAnnotation(row.dataStyle(), sheetStyle);
    if (row.dataHeightInPoints().length > 0) {
      this.useDataHeightInPoints = true;
      this.dataHeightInPoints = row.dataHeightInPoints()[0];
    }

    this.headerStyle = new CellStyleAnnotation(row.headerStyle(), sheetStyle);
    if (row.headerHeightInPoints().length > 0) {
      this.useHeaderHeightInPoints = true;
      this.headerHeightInPoints = row.headerHeightInPoints()[0];
    }

    Arrays.asList(row.headers())
        .forEach(header -> this.headers.add(
            new HeaderAnnotation(header, headerStyle)
        ));
  }
}
