package com.mz.poi.mapper;


import com.mz.poi.mapper.template.MatchTemplate;
import com.mz.poi.mapper.template.MatchTemplate.MatchTestDataRow;
import com.mz.poi.mapper.template.MatchTemplate.TestSheet;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


public class MatchTemplateSpec {

  private MatchTemplate createModel() {
    return MatchTemplate
        .builder()
        .sheet(
            TestSheet
                .builder()
                .matchTestTable(
                    Stream.of(
                        MatchTestDataRow.builder()
                            .firstValue(null)
                            .secondValue(1)
                            .build(),
                        MatchTestDataRow.builder()
                            .firstValue(null)
                            .secondValue(1)
                            .build(),
                        MatchTestDataRow.builder()
                            .firstValue(null)
                            .secondValue(1)
                            .build(),
                        MatchTestDataRow.builder()
                            .firstValue(null)
                            .secondValue(null)
                            .build()
                    ).collect(Collectors.toList())
                )
                .build()
        ).build();
  }

  @Test
  public void model_to_excel() throws IOException {
    MatchTemplate model = this.createModel();
    SXSSFWorkbook excel = ExcelMapper.toExcel(model);
    File file = new File("match_test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }

  @Test
  public void excel_to_model() {
    MatchTemplate model = this.createModel();
    SXSSFWorkbook excel = ExcelMapper.toExcel(model);
    MatchTemplate fromModel = ExcelMapper.fromExcel(excel, MatchTemplate.class);

    assert fromModel.getSheet().getMatchTestTable().size() == 4;
  }
}
