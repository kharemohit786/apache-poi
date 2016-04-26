package com.tal.poi;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.SpringApplicationConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.tal.poi.service.WorkbookWriter;

@RunWith(SpringJUnit4ClassRunner.class)
@SpringApplicationConfiguration(classes = DemoApplication.class)
public class DemoApplicationTests {
	@Test
	public void contextLoads() {
		WorkbookWriter workbookWriter = new WorkbookWriter();
		workbookWriter.WriteExcel();
	}
}

