package com.example.excelconnection;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**This program illustrates how to update an existing Microsoft Excel document.
 * Remove a sheet by its name.
 * @author Harsh Chhajer
 */

public class DeleteExcelByName {
    public static void main(String[] args) {
        String excelFilePath = "Book.xlsx"; // Path to the Excel file
        String sheetNameToDelete = "Sheet1"; // Name of the sheet to delete

        try {
            // Open the Excel file
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            // Find the sheet index by its name
            int sheetIndex = -1;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (workbook.getSheetName(i).equalsIgnoreCase(sheetNameToDelete)) {
                    sheetIndex = i;
                    break;
                }
            }

            if (sheetIndex != -1) {
                // Remove the sheet
                workbook.removeSheetAt(sheetIndex);
                System.out.println("Sheet '" + sheetNameToDelete + "' removed successfully.");
            } else {
                System.out.println("Sheet with name '" + sheetNameToDelete + "' not found.");
            }

            // Write the changes to the file
            FileOutputStream outputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }
}