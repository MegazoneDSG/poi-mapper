package com.mz.poi.mapper.structure;

import com.mz.poi.mapper.annotation.Constraint;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.DataValidation;

@Getter
@Setter
@NoArgsConstructor
public class ConstraintAnnotation {

	private boolean suppressDropDownArrow;
	private boolean showErrorBox;
	private String errorBoxTitle;
	private String errorBoxText;
	private int errorStyle;
	private String[] constraints;

	public ConstraintAnnotation(Constraint constraint) {
		this.suppressDropDownArrow = constraint.suppressDropDownArrow();
		this.showErrorBox = constraint.showErrorBox();
		this.errorBoxTitle = constraint.errorBoxTitle();
		this.errorBoxText = constraint.errorBoxText();
		switch (constraint.errorStyle()) {
			case INFO:
				this.errorStyle = DataValidation.ErrorStyle.INFO;
				break;
			case STOP:
				this.errorStyle = DataValidation.ErrorStyle.STOP;
				break;
			case WARNING:
				this.errorStyle = DataValidation.ErrorStyle.WARNING;
				break;
		}
		this.constraints = constraint.constraints();
	}
}
