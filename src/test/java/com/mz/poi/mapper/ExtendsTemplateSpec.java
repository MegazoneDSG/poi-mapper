package com.mz.poi.mapper;


import com.mz.poi.mapper.example.other.ExtendsTemplate;
import com.mz.poi.mapper.example.other.ExtendsTemplate.FirstTableRow;
import com.mz.poi.mapper.example.other.ExtendsTemplate.SecondTableRow;
import com.mz.poi.mapper.example.other.ExtendsTemplate.TestSheet;
import com.mz.poi.mapper.structure.CellStructure;
import com.mz.poi.mapper.structure.ExcelStructure;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.stream.Collectors;
import java.util.stream.Stream;


public class ExtendsTemplateSpec extends MapperTestSupport {

    private ExtendsTemplate createModel() {
        return ExtendsTemplate
            .builder()
            .sheet(
                TestSheet
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
                    .build()
            ).build();
    }

    @Test
    public void model_to_excel() throws IOException {
        ExtendsTemplate model = this.createModel();
        Workbook excel = ExcelMapper.toExcel(model);
        File file = new File(testDir + "/extends_test.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        excel.write(fos);
        fos.close();
    }

    @Test
    public void excel_to_model() {
        ExtendsTemplate model = this.createModel();
        Workbook excel = ExcelMapper.toExcel(model);
        ExtendsTemplate fromModel = ExcelMapper.fromExcel(excel, ExtendsTemplate.class);

        assert fromModel.getSheet().getFirstTable().size() == 1;
        assert fromModel.getSheet().getSecondTable().size() == 1;
    }

    @Test
    public void model_to_excel_modify_cell_structure_column() throws IOException {
        ExcelStructure structure = new ExcelStructure().build(ExtendsTemplate.class);
        CellStructure cellStructure =
            structure.getSheet("sheet").getRow("firstTable").getCell("secondValue");
        cellStructure.getAnnotation().setColumn(2);
        cellStructure.getAnnotation().setCols(2);

        ExtendsTemplate model = this.createModel();
        Workbook excel = ExcelMapper.toExcel(model, structure);
        File file = new File(testDir + "/modify_cell_structure_column.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        excel.write(fos);
        fos.close();
    }

    @Test
    public void excel_to_model_modify_cell_structure_column() {
        ExcelStructure structure = new ExcelStructure().build(ExtendsTemplate.class);
        CellStructure cellStructure =
            structure.getSheet("sheet").getRow("firstTable").getCell("secondValue");
        cellStructure.getAnnotation().setColumn(2);
        cellStructure.getAnnotation().setCols(2);

        ExtendsTemplate model = this.createModel();
        Workbook excel = ExcelMapper.toExcel(model, structure);
        ExtendsTemplate fromModel = ExcelMapper.fromExcel(excel, ExtendsTemplate.class, structure);

        assert fromModel.getSheet().getFirstTable().size() == 1;
        assert fromModel.getSheet().getSecondTable().size() == 1;
    }
}
