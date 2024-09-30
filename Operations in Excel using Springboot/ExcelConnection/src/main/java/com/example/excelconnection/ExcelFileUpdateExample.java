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
 * Append new rows to an existing sheet.
 * @author Harsh Chhajer
 */

public class ExcelFileUpdateExample {
    public static void main(String[] args) {
        String excelFilePath = "Book.xlsx";

        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(0);

            Object[][] bookData = {
                    {"The Alchemist", "Paulo Coelho", 1},
                    {"Brief Answers to the Big Questions", "Stephen Hawking", 2},
                    {"Percy Jackson Sea of Monsters", "Rick Riordan", 3},
                    {"To Sir With Love", "E.R.Braithwaite", 4},
            };

            int rowCount = sheet.getLastRowNum();

            for (Object[] aBook : bookData) {
                Row row = sheet.createRow(++rowCount);

                int columnCount = 1;

                Cell cell = row.createCell(columnCount);
                cell.setCellValue(rowCount);

                for (Object field : aBook) {
                    cell = row.createCell(++columnCount);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
            }

            inputStream.close();
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

class ExcelFileUpdateExample2 {
    public static void main(String[] args) {
        String excelFilePath = "Book.xlsx";

        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(0);

            Object[][] bookData = {
                    {"Atomic Number: ", "Element Name: ","Symbol: ","Cost(INR per Kilogram): " },
                    {"43", "Technetium", "Tc", "₹120,000,000"},
                    {"44", "Ruthenium", "Ru", "₹24,000"},
                    {"45", "Rhodium", "Rh", "₹1,800,000"},
                    {"46", "Palladium", "Pd", "₹5,200,000"},
                    {"47", "Silver", "Ag", "₹73,000"},
                    {"48", "Cadmium", "Cd", "₹1,800"},
                    {"49", "Indium", "In", "₹8,000"},
                    {"50", "Tin", "Sn", "₹2,500"},
                    {"51", "Antimony", "Sb", "₹6,000"},
                    {"52", "Tellurium", "Te", "₹6,000"},
                    {"53", "Iodine", "I", "₹4,000"},
                    {"54", "Xenon", "Xe", "₹75,000,000"},
                    {"55", "Cesium", "Cs", "₹1,500,000"},
                    {"56", "Barium", "Ba", "₹1,500"},
                    {"57", "Lanthanum", "La", "₹5,000"},
                    {"58", "Cerium", "Ce", "₹5,000"},
                    {"59", "Praseodymium", "Pr", "₹10,000"},
                    {"60", "Neodymium", "Nd", "₹8,000"},
                    {"61", "Promethium", "Pm", "₹100,000,000"},
                    {"62", "Samarium", "Sm", "₹10,000"},
                    {"63", "Europium", "Eu", "₹60,000"},
                    {"64", "Gadolinium", "Gd", "₹7,500"},
                    {"65", "Terbium", "Tb", "₹40,000"},
                    {"66", "Dysprosium", "Dy", "₹22,000"},
                    {"67", "Holmium", "Ho", "₹15,000"},
                    {"68", "Erbium", "Er", "₹10,000"},
                    {"69", "Thulium", "Tm", "₹50,000"},
                    {"70", "Ytterbium", "Yb", "₹12,000"},
                    {"71", "Lutetium", "Lu", "₹35,000"},
                    {"72", "Hafnium", "Hf", "₹4,000"},
                    {"73", "Tantalum", "Ta", "₹20,000"},
                    {"74", "Tungsten", "W", "₹3,000"},
                    {"75", "Rhenium", "Re", "₹7,000,000"},
                    {"76", "Osmium", "Os", "₹32,000"},
                    {"77", "Iridium", "Ir", "₹1,600,000"},
                    {"78", "Platinum", "Pt", "₹4,500,000"},
                    {"79", "Gold", "Au", "₹6,000,000"},
                    {"80", "Mercury", "Hg", "₹2,000"},
                    {"81", "Thallium", "Tl", "₹8,000"},
                    {"82", "Lead", "Pb", "₹2,000"},
                    {"83", "Bismuth", "Bi", "₹1,500"},
                    {"84", "Polonium", "Po", "₹500,000,000"},
                    {"85", "Astatine", "At", "₹100,000,000,000"}
            };

            int rowCount = sheet.getLastRowNum();

            for (Object[] aBook : bookData) {
                Row row = sheet.createRow(rowCount++);

                int columnCount = 0;

                Cell cell = row.createCell(columnCount);
                cell.setCellValue(rowCount);

                for (Object field : aBook) {
                    cell = row.createCell(++columnCount);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
            }

            inputStream.close();
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