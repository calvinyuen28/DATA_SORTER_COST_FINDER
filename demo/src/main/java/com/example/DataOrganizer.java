package com.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataOrganizer {
    public static void main(String[] args) throws IOException {
        List<String> codes = new ArrayList<>();

        String inputFile = "C:\\Users\\CalvinYuen\\Downloads\\TEST_EXCEL\\EXAMPLE_TESTER.xlsx"; // path to your input file
        String outputFile = "output.xlsx"; // path to your output file

        try (FileInputStream fis = new FileInputStream(inputFile);
             XSSFWorkbook inputWorkbook = new XSSFWorkbook(fis);
             XSSFWorkbook outputWorkbook = new XSSFWorkbook()) {

            XSSFSheet inputSheet = inputWorkbook.getSheetAt(0);
            XSSFSheet outputSheet = outputWorkbook.createSheet("Output");

            // Create headers in the output file
            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("CODE");
            headerRow.createCell(1).setCellValue("BILL");
            headerRow.createCell(2).setCellValue("INVOICE");
            headerRow.createCell(3).setCellValue("PROFIT");

            // Read the first column of the input file and populate the codes list
            for (Row row : inputSheet) {
                Cell cell = row.getCell(0); // First column
                if (cell != null) {
                    String code = getCellValueAsString(cell);
                    codes.add(code);
                }
            }

            // Populate the CODE column in the output file
            for (int i = 0; i < codes.size(); i++) {
                Row row = outputSheet.createRow(i + 1);
                row.createCell(0).setCellValue(codes.get(i));
                row.createCell(1).setCellValue(0); // Initialize BILL with 0
                row.createCell(2).setCellValue(0); // Initialize INVOICE with 0
                row.createCell(3).setCellValue(0); // Initialize PROFIT with 0
            }

            // Process the input file, focusing on the 2nd and 4th columns
            for (Row row : inputSheet) {
                Cell billCodeCell = row.getCell(1); // 2nd column
                Cell invoiceCodeCell = row.getCell(3); // 4th column

                if (billCodeCell != null) {
                    String billCode = getCellValueAsString(billCodeCell);
                    for (String code : codes) {
                        if (billCode.contains(code)) {
                            int codeRowIndex = codes.indexOf(code) + 1;
                            Row codeRow = outputSheet.getRow(codeRowIndex);
                            Cell billCell = codeRow.getCell(1);
                            billCell.setCellValue(billCell.getNumericCellValue() + getNumericCellValue(row.getCell(2)));
                        }
                    }
                }

                if (invoiceCodeCell != null) {
                    String invoiceCode = getCellValueAsString(invoiceCodeCell);
                    for (String code : codes) {
                        if (invoiceCode.contains(code)) {
                            int codeRowIndex = codes.indexOf(code) + 1;
                            Row codeRow = outputSheet.getRow(codeRowIndex);
                            Cell invoiceCell = codeRow.getCell(2);
                            invoiceCell.setCellValue(invoiceCell.getNumericCellValue() + getNumericCellValue(row.getCell(4)));
                        }
                    }
                }
            }

            // Calculate profit as invoice minus bill
            for (int i = 1; i <= codes.size(); i++) {
                Row row = outputSheet.getRow(i);
                Cell billCell = row.getCell(1);
                Cell invoiceCell = row.getCell(2);
                Cell profitCell = row.getCell(3);

                double bill = billCell.getNumericCellValue();
                double invoice = invoiceCell.getNumericCellValue();
                double profit = invoice - bill;

                profitCell.setCellValue(profit);
            }

            // Write the output file
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }
        }
    }

    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private static double getNumericCellValue(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue());
            } catch (NumberFormatException e) {
                return 0;
            }
        }
        return 0;
    }
}
