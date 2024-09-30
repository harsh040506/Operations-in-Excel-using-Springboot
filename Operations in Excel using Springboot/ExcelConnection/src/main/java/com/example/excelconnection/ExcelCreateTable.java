package com.example.excelconnection;

import org.apache.poi.ss.usermodel.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SpringBootApplication
public class ExcelCreateTable {

    public static void main(String[] args) {
        SpringApplication.run(ExcelCreateTable.class, args);
    }

    @RestController
    @RequestMapping("/api/excel")
    public static class ExcelController {

        @GetMapping("/create")
        public String createExcelFile() {
            try {
                createExcel("sample.xlsx");
                return "Excel file created successfully!";
            } catch (IOException e) {
                e.printStackTrace();
                return "Error occurred while creating Excel file.";
            }
        }

        private void createExcel(String filePath) throws IOException {
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Sample Sheet");

                // Define column headers
                List<String> headers = Arrays.asList("ID", "Name", "Email");
                Row headerRow = sheet.createRow(0);

                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers.get(i));
                }

                // Add sample data
                List<List<String>> data = Arrays.asList(
                        Arrays.asList("1", "Harsh Chhajer", "harsh@example.com"),
                        Arrays.asList("2", "Jane Smith", "jane@example.com")
                );

                int rowNum = 1;
                for (List<String> rowData : data) {
                    Row row = sheet.createRow(rowNum++);
                    for (int i = 0; i < rowData.size(); i++) {
                        Cell cell = row.createCell(i);
                        cell.setCellValue(rowData.get(i));
                    }
                }

                // Auto-size columns
                for (int i = 0; i < headers.size(); i++) {
                    sheet.autoSizeColumn(i);
                }

                // Write the output to a file
                try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                    workbook.write(fileOut);
                }
            }
        }
    }
}

class CreateStyledTable {
    public static void main(String[] args) {
        SpringApplication.run(CreateStyledTable.class, args);
    }

    @RestController
    @RequestMapping("/api/excel")
    public static class ExcelController {

        @GetMapping("/create")
        public String createExcelFile() {
            try {
                createExcel("sample.xlsx");
                return "Excel file created successfully!";
            } catch (IOException e) {
                e.printStackTrace();
                return "Error occurred while creating Excel file.";
            }
        }

        private void createExcel(String filePath) throws IOException {
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Sample Sheet");

                // Create a cell style with thin borders
                CellStyle thinBorderStyle = workbook.createCellStyle();
                thinBorderStyle.setBorderBottom(BorderStyle.THIN);
                thinBorderStyle.setBorderLeft(BorderStyle.THIN);
                thinBorderStyle.setBorderTop(BorderStyle.THIN);
                thinBorderStyle.setBorderRight(BorderStyle.THIN);

                // Create a cell style with thick borders
                CellStyle thickBorderStyle = workbook.createCellStyle();
                thickBorderStyle.setBorderBottom(BorderStyle.THICK);
                thickBorderStyle.setBorderLeft(BorderStyle.THICK);
                thickBorderStyle.setBorderTop(BorderStyle.THICK);
                thickBorderStyle.setBorderRight(BorderStyle.THICK);

                // Define column headers
                List<String> headers = Arrays.asList("ID", "Name", "Email");
                Row headerRow = sheet.createRow(0);

                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers.get(i));
                    cell.setCellStyle(thickBorderStyle); // Apply thick border style to header cells
                }

                // Add sample data
                List<List<String>> data = Arrays.asList(
                        Arrays.asList("1", "Harsh Chhajer", "harsh@example.com"),
                        Arrays.asList("2", "Jane Smith", "jane@example.com")
                );

                int rowNum = 1;
                for (List<String> rowData : data) {
                    Row row = sheet.createRow(rowNum++);
                    for (int i = 0; i < rowData.size(); i++) {
                        Cell cell = row.createCell(i);
                        cell.setCellValue(rowData.get(i));
                        cell.setCellStyle(thinBorderStyle); // Apply thin border style to data cells
                    }
                }

                // Auto-size columns
                for (int i = 0; i < headers.size(); i++) {
                    sheet.autoSizeColumn(i);
                }

                // Write the output to a file
                try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                    workbook.write(fileOut);
                }
            }
        }
    }
}