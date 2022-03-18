package com.mz.poi.mapper.example.csv;

import com.opencsv.bean.CsvBindByName;
import com.opencsv.bean.CsvDate;
import com.opencsv.bean.CsvNumber;
import java.time.LocalDate;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class BasicCsv {

	@CsvBindByName(column = "name1")
	private String name;

	@CsvBindByName(column = "date1")
	@CsvDate("yyyy-MM-dd")
	private LocalDate date;

	@CsvBindByName(column = "number")
	@CsvNumber("#,##0")
	private int number;
}
