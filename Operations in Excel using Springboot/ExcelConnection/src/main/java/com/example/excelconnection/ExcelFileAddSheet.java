package com.example.excelconnection;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Create a new sheet.
 * @author www.codejava.net
 */

public class ExcelFileAddSheet {
    public static void main(String[] args) {
        String excelFilePath = "Book.xlsx";

        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet newSheet = workbook.createSheet("Contacts");
            Object[][] bookComments = {
                    {"Name: ", "Birthdate: "},
                    {"Harsh Chhajer", "4th May,2006"},
                    {"Percy Jackson", "7th August,2006"},
            };

            int rowCount = 0;

            for (Object[] aBook : bookComments) {
                Row row = newSheet.createRow(++rowCount);

                int columnCount = 0;

                for (Object field : aBook) {
                    Cell cell = row.createCell(++columnCount);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
            }

            FileOutputStream outputStream = new FileOutputStream("Book.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException | EncryptedDocumentException
                 | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }
}