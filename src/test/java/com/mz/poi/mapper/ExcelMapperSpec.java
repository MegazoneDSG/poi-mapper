package com.mz.poi.mapper;


import com.mz.poi.mapper.sample.InfoRow;
import com.mz.poi.mapper.sample.ItemRow;
import com.mz.poi.mapper.sample.OrderSheet;
import com.mz.poi.mapper.sample.PurchaseOrderTemplate;
import com.mz.poi.mapper.sample.ShipRow;
import com.mz.poi.mapper.sample.SummaryRow;
import com.mz.poi.mapper.sample.TitleRow;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


public class ExcelMapperSpec {

  private PurchaseOrderTemplate createModel() {
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
                            .toTitle("Name").toValue("John").build(),
                        InfoRow.builder()
                            .vendorTitle("Company Name").vendorValue("Megazone")
                            .toTitle("Company Name").toValue("DSG").build(),
                        InfoRow.builder()
                            .vendorTitle("Address")
                            .vendorValue("MEGAZONE B/D Yeoksam-dong Gangnam-gu")
                            .toTitle("Address").toValue("DSG B/D Yeoksam-dong Gangnam-gu").build(),
                        InfoRow.builder()
                            .vendorTitle("CT, ST ZIP").vendorValue("SEOUL 06235 KOREA")
                            .toTitle("CT, ST ZIP").toValue("SEOUL 12345 KOREA").build(),
                        InfoRow.builder()
                            .vendorTitle("Phone").vendorValue("T.82(0)2 2108 9105")
                            .toTitle("Phone").toValue("T. 82 (0)2 2109 2500").build()
                    ).collect(Collectors.toList())
                )
                .shipTable(
                    Stream.of(
                        ShipRow.builder()
                            .requester("L.J")
                            .via("Purchase Part")
                            .fob("")
                            .terms("")
                            .deliveryDate("2020-11-01")
                            .build()
                    ).collect(Collectors.toList())
                )
                .itemTable(
                    Stream.of(
                        ItemRow.builder()
                            .name("#228839221").description("Product ABC").qty(1)
                            .unitPrice(BigDecimal.valueOf(150L)).build(),
                        ItemRow.builder()
                            .name("#428832121").description("Product EFG").qty(15)
                            .unitPrice(BigDecimal.valueOf(12L)).build(),
                        ItemRow.builder()
                            .name("#339884344").description("Product XYZ").qty(78)
                            .unitPrice(BigDecimal.valueOf(1.75)).build()
                    ).collect(Collectors.toList())
                )
                .summaryRow(new SummaryRow())
                .build()
        ).build();
  }

  @Test
  public void model_to_excel() throws IOException {
    PurchaseOrderTemplate model = this.createModel();
    XSSFWorkbook excel = ExcelMapper.toExcel(model);
    File file = new File("test.xlsx");
    FileOutputStream fos = new FileOutputStream(file);
    excel.write(fos);
    fos.close();
  }

  @Test
  public void excel_to_model() {
    PurchaseOrderTemplate model = this.createModel();
    XSSFWorkbook excel = ExcelMapper.toExcel(model);
    PurchaseOrderTemplate fromModel = ExcelMapper.fromExcel(excel, PurchaseOrderTemplate.class);

    assert fromModel.getSheet().getItemTable().size() == 3;
  }

}
