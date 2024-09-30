package com.example.excelconnection;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * This program illustrates how to crete a Microsoft Excel document.
 * @author Harsh Chhajer
 */

public class CreateExcelFile {
    public static void main(String[] args) {

        // Create a workbook object
        Workbook workbook = new XSSFWorkbook();

        // Create a sheet
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create a row and put some cells in it
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);

        // Set values in cells
        cell1.setCellValue("Hello");
        cell2.setCellValue("World!");

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream("Book.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        System.out.println("Excel file created successfully.");
    }
}

class CreateExcelFileStyled {
    public static void main(String[] args) {

        // Create a workbook object
        Workbook workbook = new XSSFWorkbook();

        // Create a sheet
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create a font and set it to bold
        Font font = workbook.createFont();
        font.setFontName("JetBrains Mono SemiBold"); // Set the font name to JetBrains Mono
        font.setBold(true); // Set the font style to bold
        font.setFontHeightInPoints((short) 11); // Set font size (optional)

        // Create a cell style and set the font
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);

        // Create a row and put some cells in it
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);

        // Apply the style to the cells and set values
        cell1.setCellStyle(cellStyle);
        cell1.setCellValue("Perseus");

        cell2.setCellStyle(cellStyle);
        cell2.setCellValue("Jackson");

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream("Book.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        System.out.println("Excel file created successfully with bold text in JetBrains Mono SemiBold font.");
    }
}