package com.vcs.outlook.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import javax.annotation.PostConstruct;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class WriteDataToExcel {

    @PostConstruct
    public void writeToW() throws Exception {
        String filePath = "D:\\Book1.xlsx";
        String sheetName = "InstrumentDetails";
        
        Map<String, InstrumentDetails> studentData = new TreeMap<>();
        studentData.put("1", new InstrumentDetails(5027));
        studentData.put("2", new InstrumentDetails(5027));
        studentData.put("3", new InstrumentDetails(5027));

        try {
            File file = new File(filePath);
            XSSFWorkbook workbook;
            Sheet sheet;
            
            if (file.exists()) {
                FileInputStream fis = new FileInputStream(file);
                workbook = new XSSFWorkbook(fis);
                sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    sheet = workbook.createSheet(sheetName);
                } else {
                    // Clear the sheet
                    for (int i = sheet.getLastRowNum(); i >= 0; i--) {
                        sheet.removeRow(sheet.getRow(i));
                    }
                }
                fis.close();
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet(sheetName);
            }

            int rowid = 0;

            Set<String> keyid = studentData.keySet();
            for (String key : keyid) {
                Row row = sheet.createRow(rowid++);
                InstrumentDetails objectArr = studentData.get(key);
                int cellid = 0;

                Cell cell = row.createCell(cellid++);
                cell.setCellValue(objectArr.getIndexId());
            }

            FileOutputStream out = new FileOutputStream(file);
            workbook.write(out);
            out.close();
            workbook.close();

            System.out.println("Excel file written successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

