package com.mz.poi.mapper;


import com.mz.poi.mapper.example.other.ExtendsTemplate.FirstTableRow;
import com.mz.poi.mapper.example.other.ExtendsTemplate.SecondTableRow;
import com.mz.poi.mapper.example.other.ExtendsTemplate.TestSheet;
import com.mz.poi.mapper.example.other.InheritedExtendsTemplate;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.stream.Collectors;
import java.util.stream.Stream;


public class InheritedExtendsTemplateSpec extends MapperTestSupport {

    private InheritedExtendsTemplate createModel() {
        TestSheet sheet = TestSheet
            .builder()
            .firstTable(
                Stream.of(
                    FirstTableRow.builder()
                        .firstValue("a")
                        .secondValue("b")
                        .build()
                ).collect(Collectors.toList())
            )
            .secondTable(
                Stream.of(
                    SecondTableRow.builder()
                        .firstValue(LocalDate.now().minusDays(3))
                        .build()
                ).collect(Collectors.toList())
            )
            .build();
        return new InheritedExtendsTemplate(sheet);
    }

    @Test
    public void model_to_excel() throws IOException {
        InheritedExtendsTemplate model = this.createModel();
        Workbook excel = ExcelMapper.toExcel(model);
        File file = new File(testDir + "/inherited_extends_test.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        excel.write(fos);
        fos.close();
    }

    @Test
    public void excel_to_model() {
        InheritedExtendsTemplate model = this.createModel();
        Workbook excel = ExcelMapper.toExcel(model);
        InheritedExtendsTemplate fromModel = ExcelMapper.fromExcel(excel, InheritedExtendsTemplate.class);

        assert fromModel.getSheet().getFirstTable().size() == 1;
        assert fromModel.getSheet().getSecondTable().size() == 1;
    }
}
