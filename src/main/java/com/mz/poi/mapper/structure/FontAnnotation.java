package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Font;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import static org.apache.poi.ss.usermodel.Font.ANSI_CHARSET;
import static org.apache.poi.ss.usermodel.Font.SS_NONE;
import static org.apache.poi.ss.usermodel.Font.U_NONE;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter
@Setter
@NoArgsConstructor
public class FontAnnotation {

    private String fontName = "Arial";
    private short fontHeightInPoints = 10;
    private boolean italic = false;
    private boolean strikeout = false;
    private IndexedColors color = IndexedColors.AUTOMATIC;
    private short typeOffset = SS_NONE;
    private byte underline = U_NONE;
    private int charSet = ANSI_CHARSET;
    private boolean bold = false;

    public FontAnnotation(Font font, FontAnnotation InheritanceFont) {
        this.fontName = InheritanceFont.fontName;
        this.fontHeightInPoints = InheritanceFont.fontHeightInPoints;
        this.italic = InheritanceFont.italic;
        this.strikeout = InheritanceFont.strikeout;
        this.color = InheritanceFont.color;
        this.typeOffset = InheritanceFont.typeOffset;
        this.underline = InheritanceFont.underline;
        this.charSet = InheritanceFont.charSet;
        this.bold = InheritanceFont.bold;

        if (font.fontName().length > 0) {
            this.fontName = font.fontName()[0];
        }
        if (font.fontHeightInPoints().length > 0) {
            this.fontHeightInPoints = font.fontHeightInPoints()[0];
        }
        if (font.italic().length > 0) {
            this.italic = font.italic()[0];
        }
        if (font.strikeout().length > 0) {
            this.strikeout = font.strikeout()[0];
        }
        if (font.color().length > 0) {
            this.color = font.color()[0];
        }
        if (font.typeOffset().length > 0) {
            this.typeOffset = font.typeOffset()[0];
        }
        if (font.underline().length > 0) {
            this.underline = font.underline()[0];
        }
        if (font.charSet().length > 0) {
            this.charSet = font.charSet()[0];
        }
        if (font.bold().length > 0) {
            this.bold = font.bold()[0];
        }
    }

    public void applyFont(org.apache.poi.ss.usermodel.Font font) {
        font.setFontName(this.fontName);
        font.setFontHeightInPoints(this.fontHeightInPoints);
        font.setItalic(this.italic);
        font.setStrikeout(this.strikeout);
        font.setColor(this.color.getIndex());
        font.setTypeOffset(this.typeOffset);
        font.setUnderline(this.underline);
        font.setCharSet(this.charSet);
        font.setBold(this.bold);
    }

    public FontAnnotation copy() {
        FontAnnotation annotation = new FontAnnotation();
        annotation.fontName = this.fontName;
        annotation.fontHeightInPoints = this.fontHeightInPoints;
        annotation.italic = this.italic;
        annotation.strikeout = this.strikeout;
        annotation.color = this.color;
        annotation.typeOffset = this.typeOffset;
        annotation.underline = this.underline;
        annotation.charSet = this.charSet;
        annotation.bold = this.bold;
        return annotation;
    }
}
