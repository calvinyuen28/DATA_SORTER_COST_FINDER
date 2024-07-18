package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Invoice {

    public static void main(String[] args) throws IOException {
        // Path to your Excel file
        String excelFilePath = "C:\\Users\\CalvinYuen\\OneDrive - American Bear Logistics\\CLEAN_TOOL.xlsx";

        FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet1 = workbook.getSheetAt(0);
        Sheet sheet2 = workbook.getSheetAt(1);

        // Pattern to match valid codes
        Pattern pattern = Pattern.compile("\\b[A-Z]{4}\\d{7}\\b");

        // Map to index rows in the second sheet by code
        Map<String, Integer> codeIndexMap = new HashMap<>();
        for (int j = 0; j <= sheet2.getLastRowNum(); j++) {
            Row row2 = sheet2.getRow(j);
            if (row2 != null) {
                Cell cell2 = row2.getCell(1); // Assuming the codes are in the second column of the second sheet
                if (cell2 != null && cell2.getCellType() == CellType.STRING) {
                    String cellValue = cell2.getStringCellValue();
                    Matcher matcher = pattern.matcher(cellValue);
                    while (matcher.find()) {
                        String code = matcher.group();
                        codeIndexMap.put(code, j);
                    }
                }
            }
        }

        // List to store row insertions
        List<RowInsertion> rowsToInsert = new ArrayList<>();

        // Iterate through the rows in the first sheet
        for (int i = 0; i <= sheet1.getLastRowNum(); i++) {
            Row row1 = sheet1.getRow(i);
            if (row1 != null) {
                Cell cell = row1.getCell(1); // Look in the second column
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    Matcher matcher = pattern.matcher(cellValue);
                    while (matcher.find()) {
                        String code = matcher.group();
                        // Look for the matching code in the second sheet using the index map
                        if (codeIndexMap.containsKey(code)) {
                            int insertIndex = codeIndexMap.get(code) + 1;
                            rowsToInsert.add(new RowInsertion(insertIndex, row1));
                        }
                    }
                }
            }
        }

        // Perform the row insertions in bulk to minimize row shifts
        rowsToInsert.sort((a, b) -> Integer.compare(a.rowIndex, b.rowIndex));

        int shiftOffset = 0;
        for (RowInsertion rowInsertion : rowsToInsert) {
            int targetIndex = rowInsertion.rowIndex + shiftOffset;
            sheet2.shiftRows(targetIndex, sheet2.getLastRowNum() + 1, 1);
            Row newRow = sheet2.createRow(targetIndex);
            copyRow(rowInsertion.sourceRow, newRow);
            shiftOffset++;
        }

        fileInputStream.close();

        FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();
    }

    private static void copyRow(Row sourceRow, Row destinationRow) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = destinationRow.createCell(i);

            if (oldCell == null) {
                continue;
            }

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
    }

    private static class RowInsertion {
        int rowIndex;
        Row sourceRow;

        RowInsertion(int rowIndex, Row sourceRow) {
            this.rowIndex = rowIndex;
            this.sourceRow = sourceRow;
        }
    }
}
