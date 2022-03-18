package com.mz.poi.mapper;


import com.mz.poi.mapper.example.csv.BasicCsv;
import com.opencsv.bean.CsvToBeanBuilder;
import com.opencsv.bean.StatefulBeanToCsv;
import com.opencsv.bean.StatefulBeanToCsvBuilder;
import com.opencsv.exceptions.CsvDataTypeMismatchException;
import com.opencsv.exceptions.CsvRequiredFieldEmptyException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import org.junit.Test;


public class CsvBasicSpec extends MapperTestSupport {

	@Test
	public void model_to_csv()
			throws IOException, CsvRequiredFieldEmptyException, CsvDataTypeMismatchException {
		Writer writer = new FileWriter(testDir + "/csv_basic.csv");
		StatefulBeanToCsv beanToCsv = new StatefulBeanToCsvBuilder(writer).build();
		List<BasicCsv> list = new ArrayList<>();
		list.add(new BasicCsv("Test1", LocalDate.now(), 100000));
		list.add(new BasicCsv("Test1", LocalDate.now(), 100000));
		beanToCsv.write(list);
		writer.close();
	}

	@Test
	public void csv_to_model() throws IOException {
		List<BasicCsv> beans = new CsvToBeanBuilder<BasicCsv>(
				new FileReader(testDir + "/csv_basic.csv"))
				.withType(BasicCsv.class).build().parse();
		assert beans.size() == 2;
	}

}
