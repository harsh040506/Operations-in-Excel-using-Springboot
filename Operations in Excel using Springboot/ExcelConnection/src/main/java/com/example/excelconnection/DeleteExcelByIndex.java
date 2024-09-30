package com.example.excelconnection;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Remove an existing sheet.
 * @author Harsh Chhajer
 */

public class DeleteExcelByIndex {
    public static void main(String[] args) {
        String excelFilePath = "Book.xlsx";

        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            workbook.removeSheetAt(0);

            FileOutputStream outputStream = new FileOutputStream("Book1.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }
}