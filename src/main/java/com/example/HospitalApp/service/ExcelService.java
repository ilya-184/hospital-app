package com.example.HospitalApp.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
@Deprecated
/**
 * Не работает
 */
public class ExcelService {

    public ByteArrayInputStream processExcelFiles(List<MultipartFile> files) throws IOException {
        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Result");

        for (MultipartFile file : files) {
            Workbook workbook = new XSSFWorkbook(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0);

            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Row resultRow = resultSheet.getRow(rowIndex);

                if (resultRow == null) {
                    resultRow = resultSheet.createRow(rowIndex);
                }

                if (row != null) {
                    copyRow(row, resultRow, resultWorkbook);
                }
            }
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        resultWorkbook.write(outputStream);
        resultWorkbook.close();

        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    private void copyRow(Row sourceRow, Row targetRow, Workbook workbook) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        for (int cellIndex = 0; cellIndex < sourceRow.getLastCellNum(); cellIndex++) {
            Cell sourceCell = sourceRow.getCell(cellIndex);
            Cell targetCell = targetRow.getCell(cellIndex);

            if (targetCell == null) {
                targetCell = targetRow.createCell(cellIndex);
            }

            if (sourceCell != null) {
                switch (sourceCell.getCellType()) {
                    case NUMERIC:
                        double cellValue = sourceCell.getNumericCellValue();
                        double resultValue = targetCell.getNumericCellValue();
                        targetCell.setCellValue(resultValue + cellValue);
                        break;
                    case STRING:
                        try {
                            double numericValue = Double.parseDouble(sourceCell.getStringCellValue());
                            double currentResultValue = targetCell.getNumericCellValue();
                            targetCell.setCellValue(currentResultValue + numericValue);
                        } catch (NumberFormatException e) {
                            if (targetCell.getCellType() == CellType.BLANK) {
                                targetCell.setCellValue(sourceCell.getStringCellValue());
                            }
                        }
                        break;
                    case FORMULA:
                        CellValue cellValueEvaluated = evaluator.evaluate(sourceCell);
                        if (cellValueEvaluated != null) {
                            if (cellValueEvaluated.getCellType() == CellType.NUMERIC) {
                                double numericValue = cellValueEvaluated.getNumberValue();
                                double currentResultValue = targetCell.getNumericCellValue();
                                targetCell.setCellValue(currentResultValue + numericValue);
                            } else if (cellValueEvaluated.getCellType() == CellType.STRING) {
                                try {
                                    double numericValue = Double.parseDouble(cellValueEvaluated.getStringValue());
                                    double currentResultValue = targetCell.getNumericCellValue();
                                    targetCell.setCellValue(currentResultValue + numericValue);
                                } catch (NumberFormatException e) {
                                    if (targetCell.getCellType() == CellType.BLANK) {
                                        targetCell.setCellValue(cellValueEvaluated.getStringValue());
                                    }
                                }
                            }
                        }
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case ERROR:
                        targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                        break;
                    default:
                        if (targetCell.getCellType() == CellType.BLANK) {
                            targetCell.setCellValue(sourceCell.toString());
                        }
                        break;
                }
            }
        }
    }
}