package com.mz.poi.mapper.expression;

import com.mz.poi.mapper.structure.CellType;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellStyle;

@Setter
@Getter
@NoArgsConstructor
public class CellGenerator {
    private CellType cellType;
    private Object value;
    private CellStyle style;

    @Builder
    public CellGenerator(CellType cellType, Object value, CellStyle style) {
        this.cellType = cellType;
        this.value = value;
        this.style = style;
    }
}
