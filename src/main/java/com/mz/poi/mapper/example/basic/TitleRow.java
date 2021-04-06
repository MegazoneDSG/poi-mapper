package com.mz.poi.mapper.example.basic;

import com.mz.poi.mapper.annotation.Cell;
import com.mz.poi.mapper.structure.CellType;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class TitleRow {

    @Cell(
        column = 0,
        cols = 6,
        cellType = CellType.STRING,
        ignoreParse = true
    )
    private String title = "PURCHASE ORDER";
}
