package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Header;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
public class HeaderAnnotation extends AbstractHeaderAnnotation {

    private String name;
    private List<String> mappings = new ArrayList<>();

    public HeaderAnnotation(Header header, CellStyleAnnotation headerDefaultStyle) {
        this.setStyle(new CellStyleAnnotation(header.style(), headerDefaultStyle));
        this.name = header.name();
        this.mappings = Arrays.asList(header.mappings());
    }
}
