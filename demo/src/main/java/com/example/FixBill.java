package com.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FixBill {

    public static void main(String[] args) throws IOException {
        // Path to your Excel file
        String excelFilePath = "C:\\Users\\CalvinYuen\\OneDrive - American Bear Logistics\\CLEAN_TOOL.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Check if the workbook contains at least two sheets
            if (workbook.getNumberOfSheets() < 2) {
                System.err.println("The workbook does not contain the required number of sheets.");
                return;
            }

            // Get the first two sheets
            Sheet sheet1 = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.getSheetAt(1);

            // Process the sheets
            processSheets(sheet1, sheet2);

            // Write the changes back to the file
            try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath);
                 SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook((XSSFWorkbook) workbook)) {
                sxssfWorkbook.write(fileOutputStream);
            }
        }
    }

    public static void processSheets(Sheet sheet1, Sheet sheet2) {
        // Create a map to store the row index of each code in sheet2
        Map<String, Integer> codeIndexMap = new HashMap<>();
        for (int j = 0; j <= sheet2.getLastRowNum(); j++) {
            Row row = sheet2.getRow(j);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null) {
                    String code = getCellStringValue(cell);
                    if (code != null) {
                        codeIndexMap.put(code, j);
                        System.out.println("Added to map: " + code);
                    }
                }
            }
        }

        // Iterate through each row in sheet1 and find matching codes in sheet2
        for (int i = 0; i <= sheet1.getLastRowNum(); i++) {
            Row row1 = sheet1.getRow(i);
            if (row1 != null) {
                Cell cell1 = row1.getCell(0);
                if (cell1 != null) {
                    String code1 = getCellStringValue(cell1);
                    if (code1 != null && codeIndexMap.containsKey(code1)) {
                        int matchingRowIndex = codeIndexMap.get(code1);
                        Row row2 = sheet2.getRow(matchingRowIndex);
                        if (row2 != null) {
                            Cell sourceCell = row1.getCell(1);
                            if (sourceCell != null) {
                                Cell destinationCell = row2.createCell(8); // 9th column is index 8
                                destinationCell.setCellValue(getCellStringValue(sourceCell));
                                System.out.println("Copied value from sheet1 to sheet2: " + getCellStringValue(sourceCell));
                            } else {
                                System.out.println("Source cell in sheet1 is null at row: " + i);
                            }
                        } else {
                            System.out.println("Row in sheet2 is null for matching row index: " + matchingRowIndex);
                        }
                    } else {
                        System.out.println("No matching code in sheet2 for code: " + code1);
                    }
                } else {
                    System.out.println("Cell in sheet1 is null at row: " + i);
                }
            } else {
                System.out.println("Row in sheet1 is null at index: " + i);
            }
        }
    }

    private static String getCellStringValue(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((int) cell.getNumericCellValue()).trim();
        }
        return null;
    }
}
