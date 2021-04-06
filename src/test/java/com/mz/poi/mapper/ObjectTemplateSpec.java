package com.mz.poi.mapper;


import com.mz.poi.mapper.example.object.ObjectTemplate;
import com.mz.poi.mapper.example.object.ObjectTemplate.ObjectSheet;
import com.mz.poi.mapper.example.object.ObjectTemplate.ObjectSheet.ObjectRow;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

public class ObjectTemplateSpec extends MapperTestSupport {

    private ObjectTemplate createModel() {
        List<Object> arrayObjects = Arrays.asList(10, 10, "123", null, LocalDate.now());
        List<ObjectTemplate.TestObjectCellExpression> testObjectCellExpressions = Arrays.asList(
            new ObjectTemplate.TestObjectCellExpression(10),
            new ObjectTemplate.TestObjectCellExpression(10),
            new ObjectTemplate.TestObjectCellExpression("123"),
            new ObjectTemplate.TestObjectCellExpression(null),
            new ObjectTemplate.TestObjectCellExpression(LocalDate.now())
        );

        return ObjectTemplate
            .builder()
            .sheet(
                ObjectSheet
                    .builder()
                    .objectRow(
                        ObjectRow.builder()
                            .text("Text")
                            .number(new BigDecimal("10.12"))
                            .arrayObject(arrayObjects)
                            .arrayCustom(testObjectCellExpressions)
                            .build()
                    )
                    .build()
            ).build();
    }

    @Test
    public void model_to_excel() throws IOException {
        ObjectTemplate model = this.createModel();
        Workbook excel = ExcelMapper.toExcel(model);
        File file = new File(testDir + "/object_cell_test.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        excel.write(fos);
        fos.close();
    }

    @Test
    public void excel_to_model() throws IOException {
        FileInputStream fis = new FileInputStream(testDir + "/object_cell_test.xlsx");
        XSSFWorkbook excel = new XSSFWorkbook(fis);
        ObjectTemplate fromModel = ExcelMapper.fromExcel(excel, ObjectTemplate.class);
        assert "Text".equals(fromModel.getSheet().getObjectRow().getText());
        assert fromModel.getSheet().getObjectRow().getNumber().equals(Double.valueOf("10.12"));
        assert fromModel.getSheet().getObjectRow().getArrayObject().get(0).equals(Double.valueOf("10"));
        assert fromModel.getSheet().getObjectRow().getArrayCustom().get(0).getValue().equals(Double.valueOf("10"));
    }
}
