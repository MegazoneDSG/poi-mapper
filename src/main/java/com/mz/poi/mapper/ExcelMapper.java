package com.mz.poi.mapper;


import com.mz.poi.mapper.exception.ExcelReadException;
import com.mz.poi.mapper.structure.ExcelStructure;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMapper {

    private ExcelMapper() {
        throw new UnsupportedOperationException();
    }

    public static <T> T fromExcel(
        Workbook workbook, final Class<T> excelDtoType) {
        if (workbook instanceof SXSSFWorkbook) {
            throw new ExcelReadException("Not support read operation for SXSSFWorkbook");
        }
        return new ExcelReader(workbook).read(excelDtoType);
    }

    public static <T> T fromExcel(
        Workbook workbook, final Class<T> excelDtoType, final ExcelStructure excelStructure) {
        if (workbook instanceof SXSSFWorkbook) {
            throw new ExcelReadException("Not support read operation for SXSSFWorkbook");
        }
        return new ExcelReader(workbook).read(excelDtoType, excelStructure);
    }

    public static Workbook toExcel(Object excelDto) {
        return new ExcelGenerator(excelDto, new XSSFWorkbook()).generate();
    }

    public static Workbook toExcel(Object excelDto, final ExcelStructure excelStructure) {
        return new ExcelGenerator(excelDto, new XSSFWorkbook()).generate(excelStructure);
    }

    public static Workbook toExcel(Object excelDto, Workbook workbook) {
        return new ExcelGenerator(excelDto, workbook).generate();
    }

    public static Workbook toExcel(Object excelDto, final ExcelStructure excelStructure,
                                   Workbook workbook) {
        return new ExcelGenerator(excelDto, workbook).generate(excelStructure);
    }
}
