package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Sheet;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
public class SheetAnnotation {

    private String name;
    private int index;
    private boolean protect;
    private String protectKey;
    private boolean fitToPage;
    private PrintSetupAnnotation printSetup;
    private List<ColumnWidthAnnotation> columnWidths = new ArrayList<>();
    private int defaultRowHeightInPoints;
    private int defaultColumnWidth;
    private CellStyleAnnotation defaultStyle;

    public SheetAnnotation(Sheet sheet, CellStyleAnnotation rootStyle) {
        this.name = sheet.name();
        this.index = sheet.index();
        this.protect = sheet.protect();
        this.protectKey = sheet.protectKey();
        this.fitToPage = sheet.fitToPage();
        this.printSetup = new PrintSetupAnnotation(sheet.printSetup());
        Arrays.asList(sheet.columnWidths())
            .forEach(columnWidth -> this.columnWidths.add(
                new ColumnWidthAnnotation(columnWidth)
            ));
        this.defaultRowHeightInPoints = sheet.defaultRowHeightInPoints();
        this.defaultColumnWidth = sheet.defaultColumnWidth();
        this.defaultStyle = new CellStyleAnnotation(sheet.defaultStyle(), rootStyle);
    }
}
