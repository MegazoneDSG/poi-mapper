package com.mz.poi.mapper;


import com.mz.poi.mapper.example.basic.InfoRow;
import com.mz.poi.mapper.example.basic.ItemRow;
import com.mz.poi.mapper.example.basic.OrderSheet;
import com.mz.poi.mapper.example.basic.PurchaseOrderTemplate;
import com.mz.poi.mapper.example.basic.SummaryRow;
import com.mz.poi.mapper.example.basic.TitleRow;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


public class SXSSFSpec {

  private PurchaseOrderTemplate createModel() {

    ArrayList<ItemRow> itemRows = new ArrayList<>();
    for (int i = 0; i < 200; i++) {
      itemRows.add(
          ItemRow.builder()
              .name("#228839221").description("Product ABC").qty(getRandomNumber(1, 10))
              .unitPrice(BigDecimal.valueOf(getRandomNumber(1, 200))).build());
    }

    return PurchaseOrderTemplate
        .builder()
        .sheet(
            OrderSheet
                .builder()
                .titleRow(new TitleRow())
                .infoTable(
                    Stream.of(
                        InfoRow.builder()
                            .vendorTitle("Name").vendorValue("S.Park")
                            .toTitle("Name").toValue("John").build()
                    ).collect(Collectors.toList())
                )
                .shipTable(
                    new ArrayList<>()
                )
                .itemTable(itemRows)
                .summaryRow(new SummaryRow())
                .build()
        ).build();
  }

  public int getRandomNumber(int min, int max) {
    return (int) ((Math.random() * (max - min)) + min);
  }

  @Test
  public void model_to_stream_excel() throws IOException {
    PurchaseOrderTemplate model = this.createModel();
    Workbook excel = ExcelMapper.toExcel(model, new SXSSFWorkbook(50));
    File file = new File("sxssf_test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }

  @Test
  public void excel_to_model() throws IOException {
    FileInputStream fis = new FileInputStream("sxssf_test.xlsx");
    XSSFWorkbook excel = new XSSFWorkbook(fis);
    PurchaseOrderTemplate fromModel = ExcelMapper.fromExcel(excel, PurchaseOrderTemplate.class);
    assert fromModel.getSheet().getItemTable().size() == 200;
  }

}
