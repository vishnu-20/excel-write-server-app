package com.vcs.outlook.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Collectors;

import javax.annotation.PostConstruct;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.stereotype.Service;

@Service
public class ExcelWithFilePath {

	public void writeToW() throws Exception {
		String filePath = "D:\\Book1.xlsx";

		String sheetName = "InstrumentDetails";

		List<InstrumentDetails> instrumentList = new ArrayList<>();
		instrumentList.add(new InstrumentDetails(5027, "Guitar", "Vishnu", 299.99, LocalDate.of(2024, 1, 1)));
		instrumentList.add(new InstrumentDetails(5028, "Piano", "Keyboard", 499.99, LocalDate.of(2024, 2, 1)));
		instrumentList.add(new InstrumentDetails(5029, "Flute", "Wind", 149.99, LocalDate.of(2024, 3, 1)));

		Map<String, InstrumentDetails> studentData = instrumentList.stream()
				.collect(Collectors.toMap(instrument -> String.valueOf(instrument.getIndexId()),
						instrument -> instrument, (existing, replacement) -> existing, // Handle duplicate keys
						TreeMap::new));
		try {
			File file = new File(filePath);

			XSSFWorkbook workbook;
			Sheet sheet;

			if (file.exists()) {
				try (FileInputStream fis = new FileInputStream(file)) {
					workbook = new XSSFWorkbook(fis);
					sheet = workbook.getSheet(sheetName);
					if (sheet == null) {
						sheet = workbook.createSheet(sheetName);
					} else {
						// Clear the sheet
						for (int i = sheet.getLastRowNum(); i > 0; i--) {
							sheet.removeRow(sheet.getRow(i));
						}
					}
				}
			} else {
				workbook = new XSSFWorkbook();
				sheet = workbook.createSheet(sheetName);
			}

			// Create a CellStyle for formatting dates
			CreationHelper createHelper = workbook.getCreationHelper();
			CellStyle dateCellStyle = workbook.createCellStyle();
			dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd"));

			int rowid = 1;

			Set<String> keyid = studentData.keySet();
			for (String key : keyid) {
				Row row = sheet.createRow(rowid++);
				InstrumentDetails objectArr = studentData.get(key);
				int cellid = 0;

				Cell cell = row.createCell(cellid++);
				cell.setCellValue(objectArr.getIndexId());

				cell = row.createCell(cellid++);
				cell.setCellValue(objectArr.getName());

				cell = row.createCell(cellid++);
				cell.setCellValue(objectArr.getType());

				cell = row.createCell(cellid++);
				cell.setCellValue(objectArr.getPrice());

				cell = row.createCell(cellid++);
				LocalDate localDate = objectArr.getDate();
				Date date = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
				cell.setCellValue(date);
				cell.setCellStyle(dateCellStyle); // Apply date format
			}

			// Auto-size columns to fit the contents
			for (int i = 0; i < 5; i++) {
				sheet.autoSizeColumn(i);
			}

			try (FileOutputStream out = new FileOutputStream(file)) {
				workbook.write(out);
			}
			workbook.close();

			System.out.println("Excel file written successfully.");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
