package com.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NumberFixer {

    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\CalvinYuen\\Downloads\\NUMBER_FIXER_7_18_2024.xlsx";
        String outputFilePath = "C:\\Users\\CalvinYuen\\Downloads\\NUMBER_FIXER_OUTPUT.xlsx";
        convertNumbersToDates(inputFilePath, outputFilePath);
    }

    public static void convertNumbersToDates(String inputFilePath, String outputFilePath) {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook inputWorkbook = new XSSFWorkbook(fis);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet inputSheet = inputWorkbook.getSheetAt(0); // Get the first sheet from the input file
            Sheet outputSheet = outputWorkbook.createSheet("Converted Dates"); // Create a new sheet in the output file

            // Create a cell style for dates in the output file
            CellStyle dateCellStyle = outputWorkbook.createCellStyle();
            CreationHelper createHelper = outputWorkbook.getCreationHelper();
            dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));

            for (Row row : inputSheet) {
                Row outputRow = outputSheet.createRow(row.getRowNum());

                for (Cell cell : row) {
                    Cell outputCell = outputRow.createCell(cell.getColumnIndex());

                    if (cell.getCellType() == CellType.NUMERIC) {
                        double numericValue = cell.getNumericCellValue();
                        if (DateUtil.isCellDateFormatted(cell)) {
                            outputCell.setCellValue(cell.getDateCellValue());
                            outputCell.setCellStyle(dateCellStyle);
                        } else {
                            LocalDate date = LocalDate.of(1899, 12, 30).plusDays((long) numericValue);
                            Date javaDate = Date.from(date.atStartOfDay(ZoneId.systemDefault()).toInstant());
                            outputCell.setCellValue(javaDate);
                            outputCell.setCellStyle(dateCellStyle);
                        }
                    } else {
                        switch (cell.getCellType()) {
                            case STRING:
                                outputCell.setCellValue(cell.getStringCellValue());
                                break;
                            case BOOLEAN:
                                outputCell.setCellValue(cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                outputCell.setCellFormula(cell.getCellFormula());
                                break;
                            case BLANK:
                                outputCell.setBlank();
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
