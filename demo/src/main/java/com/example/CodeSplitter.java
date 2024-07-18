package com.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CodeSplitter {

    public static void main(String[] args) throws IOException {
        // Path to your Excel file
        String excelFilePath = "C:\\Users\\CalvinYuen\\OneDrive - American Bear Logistics\\CLEAN_TOOL.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Check if the workbook contains at least one sheet
            if (workbook.getNumberOfSheets() < 1) {
                System.err.println("The workbook does not contain any sheets.");
                return;
            }

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Process the sheet
            processSheet(sheet);

            // Write the changes back to the file
            try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath);
                 SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook((XSSFWorkbook) workbook)) {
                sxssfWorkbook.write(fileOutputStream);
            }
        }
    }

    private static void processSheet(Sheet sheet) {
        // Regular expression to match the valid codes
        Pattern codePattern = Pattern.compile("\\b[A-Z]{4}\\d{7}\\b");

        // List to keep track of new rows to add
        List<RowData> newRowsData = new ArrayList<>();

        // Iterate over the rows of the sheet
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            // Get the cell in the second column (index 1)
            Cell cell = row.getCell(1);

            if (cell != null && cell.getCellType() == CellType.STRING) {
                String cellValue = cell.getStringCellValue();
                Matcher matcher = codePattern.matcher(cellValue);

                // List to store valid codes found in the cell
                List<String> validCodes = new ArrayList<>();

                // Find all matches and add them to the list
                while (matcher.find()) {
                    validCodes.add(matcher.group());
                }

                // If multiple codes are found, prepare new rows for each additional code
                if (validCodes.size() > 1) {
                    for (int j = 1; j < validCodes.size(); j++) {
                        newRowsData.add(new RowData(i, validCodes.get(j)));
                    }
                }

                // Set the cell value to the first code
                if (!validCodes.isEmpty()) {
                    cell.setCellValue(validCodes.get(0));
                }
            }
        }

        // Shift rows and add the new rows to the sheet
        for (RowData rowData : newRowsData) {
            sheet.shiftRows(rowData.rowNum + 1, sheet.getLastRowNum(), 1);
            Row newRow = sheet.createRow(rowData.rowNum + 1);
            copyRowData(sheet.getRow(rowData.rowNum), newRow);
            newRow.getCell(1).setCellValue(rowData.code);
        }
    }

    private static void copyRowData(Row sourceRow, Row newRow) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell oldCell = sourceRow.getCell(i);
            if (oldCell != null) {
                Cell newCell = newRow.createCell(i);
                copyCell(oldCell, newCell);
            }
        }
    }

    private static void copyCell(Cell oldCell, Cell newCell) {
        newCell.setCellStyle(oldCell.getCellStyle());
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            case BLANK:
                newCell.setBlank();
                break;
            default:
                break;
        }
    }

    private static class RowData {
        int rowNum;
        String code;

        RowData(int rowNum, String code) {
            this.rowNum = rowNum;
            this.code = code;
        }
    }
}
