package com.vcs.outlook.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
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
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;

import org.springframework.stereotype.Service;

@Service
public class WriteDataToExcel {

	@Autowired
	private ResourceLoader resourceLoader;

	
//	@PostConstruct
	public void pub() throws Exception {
        String sheetName = "InstrumentDetails";
//		writeToW(sheetName);
//		writeToW("InstrumentDetails1");
	}

	public void writeToW(String sheetName , List<InstrumentDetails> instrumentList) throws Exception {
        Resource resource = resourceLoader.getResource("classpath:Book1.xlsx");

        // Copy the resource file to a temporary location
//        File tempFile = resource.getFile();
        
        File tempFile = new File("src/main/resources/Book1.xlsx");





        Map<String, InstrumentDetails> studentData = instrumentList.stream()
                .collect(Collectors.toMap(instrument -> String.valueOf(instrument.getIndexId()),
                        instrument -> instrument, (existing, replacement) -> existing, // Handle duplicate keys
                        TreeMap::new));

        try {
            XSSFWorkbook workbook;
            Sheet sheet;

            try (FileInputStream fis = new FileInputStream(tempFile)) {
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
            File originalFile = new File("src/main/resources/Book1.xlsx");

            try (FileOutputStream out = new FileOutputStream(originalFile)) {
                workbook.write(out);
                out.close();
                
            }
            workbook.close();

            System.out.println("Excel file written successfully.");

//            // Replace the original file in the resources directory
//            Files.copy(tempFile.toPath(), originalFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Clean up the temporary file if necessary
            tempFile.deleteOnExit();
        }
    }

}
