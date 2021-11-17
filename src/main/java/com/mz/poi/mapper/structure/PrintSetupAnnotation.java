package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.PrintSetup;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class PrintSetupAnnotation {

	private short paperSize;
	private short scale;
	private short pageStart;
	private short fitWidth;
	private short fitHeight;
	private boolean leftToRight;
	private boolean landscape;
	private boolean validSettings;
	private boolean noColor;
	private boolean draft;
	private boolean notes;
	private boolean noOrientation;
	private boolean usePage;
	private short hResolution;
	private short vResolution;
	private double headerMargin;
	private double footerMargin;
	private short copies;


	public PrintSetupAnnotation(PrintSetup printSetup) {
		paperSize = printSetup.paperSize();
		scale = printSetup.scale();
		pageStart = printSetup.pageStart();
		fitWidth = printSetup.fitWidth();
		fitHeight = printSetup.fitHeight();
		leftToRight = printSetup.leftToRight();
		landscape = printSetup.landscape();
		validSettings = printSetup.validSettings();
		noColor = printSetup.noColor();
		draft = printSetup.draft();
		notes = printSetup.notes();
		noOrientation = printSetup.noOrientation();
		usePage = printSetup.usePage();
		hResolution = printSetup.hResolution();
		vResolution = printSetup.vResolution();
		headerMargin = printSetup.headerMargin();
		footerMargin = printSetup.footerMargin();
		copies = printSetup.copies();
	}
}
