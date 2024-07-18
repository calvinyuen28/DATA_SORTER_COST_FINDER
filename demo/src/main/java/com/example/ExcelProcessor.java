package com.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelProcessor {

    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\CalvinYuen\\Downloads\\TEST_EXCEL\\EXAMPLE_TESTER.xlsx";
        String newExcelFilePath = "C:\\Users\\CalvinYuen\\Downloads\\TEST_EXCEL\\EXAMPLE_TESTER_UPDATED.xlsx";
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Reading the sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Create a map to store IDS and tx_memo_in
            Map<String, String> idTxMemoMap = new HashMap<>();
            ArrayList<String> codes = new ArrayList<>();

            // Populate the map and array
            populateMapAndArray(sheet, idTxMemoMap, codes);

            // Process the LINE_DSCRP and update as needed
            updateLineDescriptions(sheet, idTxMemoMap, codes);

            // Write changes to a new excel file
            try (FileOutputStream fos = new FileOutputStream(newExcelFilePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void populateMapAndArray(Sheet sheet, Map<String, String> idTxMemoMap, ArrayList<String> codes) {
        int rowIndex = 0;
        while (true) {
            Row row = sheet.getRow(rowIndex++);
            if (row == null) break;
            
            Cell idCell = row.getCell(2); // Third column
            Cell txMemoCell = row.getCell(3); // Fourth column
            Cell codeCell = row.getCell(4); // Fifth column

            if (idCell == null || txMemoCell == null) break;
            
            String id = getCellValueAsString(idCell);
            String txMemo = getCellValueAsString(txMemoCell);
            
            if (id.isEmpty() && txMemo.isEmpty()) break;
            
            idTxMemoMap.put(id, txMemo);

            if (codeCell != null && !getCellValueAsString(codeCell).isEmpty()) {
                codes.add(getCellValueAsString(codeCell));
            }
        }
    }

    private static void updateLineDescriptions(Sheet sheet, Map<String, String> idTxMemoMap, ArrayList<String> codes) {
        int rowIndex = 0;
        while (true) {
            Row row = sheet.getRow(rowIndex++);
            if (row == null) break;

            Cell lineDscrpCell = row.getCell(0); // First column
            Cell txnQbIdCell = row.getCell(1); // Second column

            if (txnQbIdCell == null) continue; // Skip if txnQbIdCell is null

            String lineDscrp = (lineDscrpCell != null) ? getCellValueAsString(lineDscrpCell) : "";
            String txnQbId = getCellValueAsString(txnQbIdCell);

            boolean containsCode = false;
            for (String code : codes) {
                if (lineDscrp.contains(code)) {
                    containsCode = true;
                    break;
                }
            }

            if ((lineDscrp.isEmpty() || !containsCode) && idTxMemoMap.containsKey(txnQbId)) {
                String txMemo = idTxMemoMap.get(txnQbId);
                if (lineDscrpCell == null) {
                    lineDscrpCell = row.createCell(0); // Create cell if it doesn't exist
                }
                lineDscrpCell.setCellValue(lineDscrp + " " + txMemo);
            }

            // Debugging statement
            System.out.println("Row " + rowIndex + ": " + lineDscrpCell.getStringCellValue());
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
