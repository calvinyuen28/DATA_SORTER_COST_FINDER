package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ForkJoinPool;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BillInvoice {

    public static void main(String[] args) throws IOException {
        // Path to your Excel file
        String excelFilePath = "C:\\Users\\CalvinYuen\\Downloads\\CLEAN_TOOL.xlsx";

        FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet1 = workbook.getSheetAt(0);
        Sheet sheet2 = workbook.getSheetAt(1);

        Pattern pattern = Pattern.compile("\\b[A-Z]{4}\\d{7}( DFW)?\\b");
        
        Map<String, List<Row>> rowMap = new ConcurrentHashMap<>();

        ForkJoinPool customThreadPool = new ForkJoinPool(Runtime.getRuntime().availableProcessors());

        customThreadPool.submit(() ->
            IntStream.range(0, sheet1.getPhysicalNumberOfRows()).parallel().forEach(i -> {
                Row row = sheet1.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0);
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        Matcher matcher = pattern.matcher(cellValue);
                        while (matcher.find()) {
                            String code = matcher.group();
                            rowMap.computeIfAbsent(code, k -> new ArrayList<>()).add(row);
                        }
                    }
                }
            })
        ).join();

        fileInputStream.close();

        for (Map.Entry<String, List<Row>> entry : rowMap.entrySet()) {
            String code = entry.getKey();
            for (Row sourceRow : entry.getValue()) {
                Row newRow = sheet2.createRow(sheet2.getLastRowNum() + 1);
                copyRow(sourceRow, newRow);
                newRow.getCell(0).setCellValue(code);
            }
        }

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
                newCell = null;
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
}
