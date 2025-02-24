package com.example.demo.controller;

import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpStatus;
import org.springframework.core.io.ClassPathResource;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.opencsv.CSVWriter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
public class FileUploadController {

    private static final Logger logger = LoggerFactory.getLogger(FileUploadController.class);

    @PostMapping("/submit")
    public ResponseEntity<String> handleFileUpload() {
        try {
            String excelFilePath = "files/feedback.xlsx";
            ClassPathResource excelResource = new ClassPathResource(excelFilePath);
            File excelFile = excelResource.getFile();
            String csvFilePath = excelFile.getAbsolutePath().replace(".xlsx", ".csv");
            File csvFile = new File(csvFilePath);

            logger.info("Excel file absolute path: {}", excelFile.getAbsolutePath());
            logger.info("CSV file absolute path: {}", csvFile.getAbsolutePath());

            FileInputStream fis = new FileInputStream(excelFile);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            fis.close();
            Sheet sheet = workbook.getSheetAt(0);

            if (csvFile.exists() && csvFile.delete()) {
                logger.info("Existing CSV file deleted successfully: {}", csvFile.getAbsolutePath());
            } else {
                logger.warn("Failed to delete existing CSV file: {}", csvFile.getAbsolutePath());
            }

            exportToCsv(workbook, csvFile);

            logger.info("Excel file read and exported to CSV successfully!");
            return ResponseEntity.ok("Excel file read and exported to CSV successfully!");

        } catch (IOException e) {
            logger.error("Error processing files: ", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Error processing files: " + e.getMessage());
        }
    }

    private void exportToCsv(Workbook workbook, File csvFile) throws IOException {
        try (CSVWriter writer = new CSVWriter(new FileWriter(csvFile))) {
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            int rowCount = sheet.getPhysicalNumberOfRows();
            logger.debug("Exporting {} rows to CSV...", rowCount);

            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    logger.debug("Processing row: {}", i + 1);
                    List<String> csvRowData = new ArrayList<>();
                    short minColIx = row.getFirstCellNum();
                    short maxColIx = row.getLastCellNum();
                    logger.debug("  Row {} - First Column Index: {}, Last Column Index: {}", i + 1, minColIx, maxColIx);

                    if (maxColIx == -1) continue;

                    for (short colIx = minColIx; colIx < maxColIx; colIx++) {
                        Cell cell = row.getCell(colIx);
                        if (cell == null) {
                            logger.debug("  Row {} Col {} - Cell is NULL", i + 1, colIx + 1);
                            csvRowData.add("");
                            continue;
                        }

                        CellType cellType = cell.getCellType();
                        String cellValueToExport;

                        if (cellType == CellType.FORMULA) {
                            try {
                                CellValue evaluatedCellValue = evaluator.evaluate(cell);
                                switch (evaluatedCellValue.getCellType()) {
                                    case STRING:
                                        cellValueToExport = evaluatedCellValue.getStringValue();
                                        break;
                                    case NUMERIC:
                                        cellValueToExport = formatter.formatCellValue(cell);
                                        break;
                                    case BOOLEAN:
                                        cellValueToExport = String.valueOf(evaluatedCellValue.getBooleanValue());
                                        break;
                                    case ERROR:
                                        cellValueToExport = "ERROR: " + evaluatedCellValue.getErrorValue();
                                        break;
                                    case BLANK:
                                    case _NONE:
                                        cellValueToExport = "";
                                        break;
                                    default:
                                        cellValueToExport = "";
                                        logger.warn("  Row {} Col {} - Unexpected CellType after formula evaluation (using empty string): {}", i + 1, colIx + 1, evaluatedCellValue.getCellType());
                                }
                                logger.debug("  Row {} Col {} - Cell Type: FORMULA, Evaluated Value: {}, Formula: {}",
                                        i + 1, colIx + 1, cellValueToExport, cell.getCellFormula());
                            } catch (Exception formulaEx) {
                                logger.warn("  Row {} Col {} - Formula evaluation error: {}", i + 1, colIx + 1, formulaEx.getMessage());
                                cellValueToExport = "FORMULA_ERROR";
                            }
                        } else {
                            cellValueToExport = formatter.formatCellValue(cell);
                            logger.debug("  Row {} Col {} - Cell Type: {}, Formatted Value: {}, Formula (if any): N/A",
                                    i + 1, colIx + 1, cellType, cellValueToExport);
                        }
                        csvRowData.add(cellValueToExport);
                    }
                    writer.writeNext(csvRowData.toArray(new String[0]));
                } else {
                    logger.debug("Row {} is NULL", i + 1);
                }
            }
        } catch (IOException csvEx) {
            logger.error("IO Error during CSV export: {}", csvEx.getMessage());
            throw csvEx;
        }
        logger.info("Excel data exported to CSV file (with formula evaluation): {}", csvFile.getAbsolutePath());
    }
}
