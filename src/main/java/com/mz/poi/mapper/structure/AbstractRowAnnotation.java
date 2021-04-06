package com.mz.poi.mapper.structure;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public abstract class AbstractRowAnnotation {

    private int row;
    private String rowAfter;
    private int rowAfterOffset;
}
