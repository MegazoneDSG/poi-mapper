package com.mz.poi.mapper.helper;

import com.mz.poi.mapper.structure.CellStyleAnnotation;
import com.mz.poi.mapper.structure.FontAnnotation;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

public class CachedStyleRepository {
    private Workbook workbook;
    private Map<String, CellStyle> cachedDataRowStyle = new HashMap<>();

    public CachedStyleRepository(Workbook workbook) {
        this.workbook = workbook;
    }

    public CellStyle createStyle(CellStyleAnnotation annotation) {
        return Optional.ofNullable(this.getStyle(annotation.getKey()))
            .orElse(this.createNewStyle(annotation));
    }

    private CellStyle createNewStyle(CellStyleAnnotation annotation) {
        FontAnnotation fontAnnotation = annotation.getFont();
        Font font = this.workbook.createFont();
        fontAnnotation.applyFont(font);

        CellStyle cellStyle = this.workbook.createCellStyle();
        annotation.applyStyle(cellStyle, font, this.workbook);
        cachedDataRowStyle.put(annotation.getKey(), cellStyle);
        return cellStyle;
    }

    public CellStyle getStyle(String key) {
        return cachedDataRowStyle.get(key);
    }

    public boolean hasStyle(String key) {
        return cachedDataRowStyle.containsKey(key);
    }

    public long getSize() {
        return cachedDataRowStyle.size();
    }
}
